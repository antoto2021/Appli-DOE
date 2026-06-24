// script.js

let rootDirectoryHandle = null;
let folderMapping = {}; 
let averifierHandle = null;
let excelMapping = {}; // { code: titre }

// UI Elements
const btnSelectDir = document.getElementById('btnSelectDir');
const statusDir = document.getElementById('statusDir');
const inputDocx = document.getElementById('inputDocx');
const statusDocx = document.getElementById('statusDocx');
const inputZip = document.getElementById('inputZip');
const statusZip = document.getElementById('statusZip');
const inputExcel = document.getElementById('inputExcel');
const statusExcel = document.getElementById('statusExcel');
const logConsole = document.getElementById('logConsole');
const btnClearLog = document.getElementById('btnClearLog');

const step2 = document.getElementById('step-2');
const step3 = document.getElementById('step-3');
const step4 = document.getElementById('step-4');

// Utility to write to console
function log(message, type = 'info') {
    const entry = document.createElement('div');
    entry.className = `log-entry log-${type}`;
    
    // Format timestamp
    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour12: false, hour: '2-digit', minute:'2-digit', second:'2-digit' });
    
    let typeIcon = '';
    if (type === 'success') typeIcon = '✓';
    if (type === 'error') typeIcon = '✖';
    if (type === 'warning') typeIcon = '⚠';
    if (type === 'info') typeIcon = 'ℹ';
    if (type === 'system') typeIcon = '⚡';

    entry.textContent = `[${timeStr}] ${typeIcon} ${message}`;
    logConsole.appendChild(entry);
    logConsole.scrollTop = logConsole.scrollHeight;
}

btnClearLog.addEventListener('click', () => {
    logConsole.innerHTML = '<div class="log-entry log-system">Console effacée.</div>';
});

// Étape 1 : Sélection du dossier
btnSelectDir.addEventListener('click', async () => {
    try {
        rootDirectoryHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
        
        statusDir.textContent = `Dossier actif : ${rootDirectoryHandle.name}`;
        statusDir.className = "status-indicator success";
        
        // Création du dossier "A vérifier" à la racine
        averifierHandle = await rootDirectoryHandle.getDirectoryHandle("A vérifier", { create: true });
        log("Dossier de repli 'A vérifier' configuré à la racine.", 'system');

        // Réactivation de l'étape 2
        step2.classList.remove('disabled');
        inputDocx.disabled = false;
        statusDocx.textContent = "Prêt. Sélectionnez la trame .docx.";
        statusDocx.className = "status-indicator";

        // Scan des dossiers existants
        await scanExistingFolders(rootDirectoryHandle);
        
    } catch (err) {
        log(`Accès refusé ou annulé : ${err.message}`, 'error');
        statusDir.textContent = "Sélection annulée ou refusée.";
        statusDir.className = "status-indicator error";
    }
});

async function scanExistingFolders(currentHandle) {
    let count = 0;
    for await (const entry of currentHandle.values()) {
        if (entry.kind === 'directory') {
            const match = entry.name.match(/^(\d+(\.\d+)*)\.?\s+(.*)/);
            if (match) {
                folderMapping[match[1]] = entry;
                count++;
            }
            // On peut explorer récursivement si besoin, mais ici on scanne le premier niveau principalement
            // Ou récursivement pour les sous-dossiers :
            await scanSubFolders(entry);
        }
    }
    
    if (Object.keys(folderMapping).length > 0) {
        log(`Arborescence détectée en mémoire (${Object.keys(folderMapping).length} dossiers indexés).`, 'success');
        // L'étape 2 (Trame) devient OBLIGATOIRE, on n'active donc pas l'étape 3 ici.
        // On ne modifie pas non plus le statut de la case 2.
    }
}

async function scanSubFolders(currentHandle) {
    for await (const entry of currentHandle.values()) {
        if (entry.kind === 'directory') {
            const match = entry.name.match(/^(\d+(\.\d+)*)\.?\s+(.*)/);
            if (match) {
                folderMapping[match[1]] = entry;
            }
            await scanSubFolders(entry);
        }
    }
}

// Étape 2 : Parsing DOCX
inputDocx.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    statusDocx.textContent = "Analyse de la trame en cours...";
    statusDocx.className = "status-indicator";
    log(`Analyse du fichier trame : ${file.name}`, 'info');

    try {
        const zip = await JSZip.loadAsync(await file.arrayBuffer());
        const xmlText = await zip.file("word/document.xml").async("text");
        const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
        const paragraphs = xmlDoc.getElementsByTagName("w:p");
        let tempStructure = {};
        let folderCount = 0;

        for (let p of paragraphs) {
            let pText = Array.from(p.getElementsByTagName("w:t")).map(t => t.textContent).join('').trim();
            const match = pText.match(/^(\d+(\.\d+)*)\.?\s+(.*)/);
            
            if (match) {
                const numPrefix = match[1];
                const cleanName = `${numPrefix}. ${match[3]}`.replace(/[\\/*?:"<>|]/g, "_").trim();
                const parts = numPrefix.split('.');
                
                let parentHandle = parts.length === 1 ? rootDirectoryHandle : (tempStructure[parts.slice(0, -1).join('.')] || rootDirectoryHandle);
                let dirHandle;
                
                // Si le dossier existe déjà (trouvé à l'étape 1), on le réutilise directement sans solliciter OneDrive
                if (folderMapping[numPrefix]) {
                    dirHandle = folderMapping[numPrefix];
                } else {
                    // Sinon, on le crée physiquement
                    try {
                        dirHandle = await parentHandle.getDirectoryHandle(cleanName, { create: true });
                        folderCount++; // On compte seulement les NOUVEAUX dossiers créés
                    } catch (err) {
                        if (err.name === 'InvalidStateError') {
                            // OneDrive a invalidé le dossier parent en arrière-plan car on vient d'y créer un sous-dossier !
                            // On va regénérer un handle tout neuf depuis la racine.
                            try {
                                let freshHandle = rootDirectoryHandle;
                                let currentPrefix = "";
                                for (let i = 0; i < parts.length - 1; i++) {
                                    currentPrefix = currentPrefix ? `${currentPrefix}.${parts[i]}` : parts[i];
                                    let deadHandle = folderMapping[currentPrefix];
                                    if (deadHandle) {
                                        freshHandle = await freshHandle.getDirectoryHandle(deadHandle.name);
                                    }
                                }
                                parentHandle = freshHandle;
                                
                                // On met à jour le cache avec ce handle tout neuf
                                const parentPrefix = parts.slice(0, -1).join('.');
                                tempStructure[parentPrefix] = parentHandle;
                                folderMapping[parentPrefix] = parentHandle;

                                // On retente la création
                                dirHandle = await parentHandle.getDirectoryHandle(cleanName, { create: true });
                                folderCount++;
                            } catch (retryErr) {
                                log(`Échec création ${cleanName} après restauration : ${retryErr.message}`, 'warning');
                            }
                        } else {
                            log(`Échec création ${cleanName} : ${err.name} - ${err.message}`, 'warning');
                        }
                    }
                }
                
                if (dirHandle) {
                    tempStructure[numPrefix] = dirHandle;
                    folderMapping[numPrefix] = dirHandle;
                }
            }
        }
        
        // REVÉRIFICATION: Comme suggéré, on lance un scan complet de la racine pour
        // rattraper tous les dossiers qui ont été créés malgré l'alerte rouge !
        log("Revérification des dossiers présents à la racine...", 'info');
        try {
            await scanSubFolders(rootDirectoryHandle);
        } catch (scanErr) {
            log(`Impossible de revérifier la racine : ${scanErr.message}`, 'warning');
        }
        
        statusDocx.textContent = "Trame chargée et vérifiée avec succès.";
        statusDocx.className = "status-indicator success";
        log(`Arborescence prête. Nouveaux dossiers créés: ${folderCount}`, 'success');
        
        step3.classList.remove('disabled');
        inputExcel.disabled = false;
        statusExcel.textContent = "Prêt. Importez l'Excel.";
        statusExcel.className = "status-indicator";

    } catch (err) { 
        log(`Erreur DOCX : ${err.message}`, 'error'); 
        statusDocx.textContent = err.message.includes("OneDrive") ? err.message : "Erreur lors de la lecture. Vérifiez la console.";
        statusDocx.className = "status-indicator error";
    }
    
    // Reset l'input pour permettre de re-sélectionner le même fichier si besoin
    e.target.value = '';
});

// Étape 3 : Parsing de l'Excel (Export GED)
inputExcel.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    statusExcel.textContent = "Analyse de l'Excel en cours...";
    log(`Analyse du fichier Excel : ${file.name}`, 'info');

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to array of arrays
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        excelMapping = [];
        let count = 0;

        // Skip header row
        for (let i = 1; i < json.length; i++) {
            const row = json[i];
            if (!row || row.length < 3) continue;
            
            const code = row[0] ? row[0].toString().trim() : '';
            const titre = row[1] ? row[1].toString().trim() : '';
            const statut = row[2] ? row[2].toString().trim().toUpperCase() : '';
            const chrono = row[7] ? row[7].toString().trim() : '';
            const fichierPrincipal = row[10] ? row[10].toString().trim() : '';

            if (statut.includes('BPE') && code && titre) {
                excelMapping.push({ code, titre, chrono, fichierPrincipal });
                count++;
            }
        }

        statusExcel.textContent = "Excel chargé avec succès.";
        statusExcel.className = "status-indicator success";
        log(`Mapping Excel chargé : ${count} documents BPE indexés.`, 'success');

        step4.classList.remove('disabled');
        inputZip.disabled = false;
        statusZip.textContent = "Prêt. Importez le ZIP.";
        statusZip.className = "status-indicator";
    } catch (err) {
        log(`Erreur Excel : ${err.message}`, 'error');
        statusExcel.textContent = "Erreur lors de la lecture.";
        statusExcel.className = "status-indicator error";
    }

    e.target.value = '';
});

// Étape 4 : Traitement du ZIP
inputZip.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    statusZip.textContent = "Extraction et analyse en cours...";
    log(`Début du traitement du fichier : ${file.name}`, 'info');

    try {
        const zip = await JSZip.loadAsync(await file.arrayBuffer());
        let bpeRegistry = {};
        let fileCount = 0;

        // Phase 1 : Filtre Récence sur le mot "BPE" ou récupération de tout
        for (const [path, entry] of Object.entries(zip.files)) {
            if (entry.dir) continue;
            const parts = path.split('/');
            const filename = parts[parts.length - 1];

            // Ignore les fichiers cachés (comme .DS_Store)
            if (filename.startsWith('.')) continue;

            // Ne traiter que les fichiers PDF
            if (!filename.toLowerCase().endsWith('.pdf')) continue;

            // IGNORER STRICTEMENT LES CARTOUCHES (vérifier dans tout le chemin)
            if (path.toLowerCase().includes('cartouche')) {
                log(`Fichier ignoré (Cartouche) : ${filename}`, 'warning');
                continue;
            }

            // IGNORER LES ANCIENS INDICES (s'il n'y a pas "BPE" dans le nom)
            const fnLower = filename.toLowerCase();
            const isBPE = fnLower.includes('bpe');
            if (!isBPE && fnLower.match(/\[ind\s*[a-z]\]|\[[a-z]\]|-a\.|-b\.|-c\./i)) {
                log(`Fichier ignoré (Ancien indice) : ${filename}`, 'warning');
                continue;
            }

            // Filtre de récence pour les fichiers du même nom
            if (!bpeRegistry[filename] || entry.date > bpeRegistry[filename].date) {
                bpeRegistry[filename] = { entry: entry, date: entry.date, parent: parts.length > 1 ? parts[parts.length - 2] : "" };
            }
        }

        // Phase 2 : Extraction et routage
        for (const [filename, data] of Object.entries(bpeRegistry)) {
            // Ne plus ajouter le dossier parent pour garder un nom propre
            let finalName = filename;
            
            // Check against Excel mapping
            let matchedTitre = null;
            let usedCode = null;
            
            for (const mapItem of excelMapping) {
                const fnLower = filename.toLowerCase();
                // Match exact du fichier principal OU inclusion du chrono OU inclusion du code
                if (
                    (mapItem.fichierPrincipal && fnLower === mapItem.fichierPrincipal.toLowerCase()) ||
                    (mapItem.chrono && filename.includes(mapItem.chrono)) ||
                    (mapItem.code && filename.includes(mapItem.code))
                ) {
                    matchedTitre = mapItem.titre;
                    usedCode = mapItem.code;
                    break;
                }
            }

            // Determine target handle based on mapping Titre or filename fallback
            let targetHandle;
            if (matchedTitre) {
                targetHandle = getTargetHandle(matchedTitre);
            } else {
                // Si pas de match Excel ET pas de "BPE" dans le nom -> direction "A vérifier"
                if (!filename.toLowerCase().includes("bpe")) {
                    targetHandle = averifierHandle;
                } else {
                    targetHandle = getTargetHandle(filename);
                }
            }

            // --- STANDARDISATION DU NOM DE FICHIER ---
            const lastDotIndex = finalName.lastIndexOf('.');
            let baseName = lastDotIndex !== -1 ? finalName.substring(0, lastDotIndex) : finalName;
            const extPart = lastDotIndex !== -1 ? finalName.substring(lastDotIndex) : '.pdf';

            // 1. Enlever les anciens indices entre crochets (ex: [Ind A], [A], [C])
            baseName = baseName.replace(/^\[(Ind\s+[a-zA-Z]|[a-zA-Z])\]\s*/i, '');

            // 2. Enlever toutes les mentions BPE existantes pour éviter les doublons
            // ex: [BPE], BPE, - BPE, _BPE
            baseName = baseName.replace(/\[BPE\]/ig, '');
            baseName = baseName.replace(/[\s\-_]*BPE[\s\-_]*/ig, ' ');

            // 3. Nettoyer les espaces résiduels
            baseName = baseName.trim().replace(/\s{2,}/g, ' ');

            // 4. Ajouter proprement _BPE à la fin
            let cleanFinalName = `${baseName}_BPE${extPart}`;

            // SANITIZATION: Remove characters that are invalid in Windows file names
            cleanFinalName = cleanFinalName.replace(/[\\/*?:"<>|]/g, "_");

            try {
                const blob = await data.entry.async("blob");
                const fileHandle = await targetHandle.getFileHandle(cleanFinalName, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(blob);
                await writable.close();
                
                let statusLog = targetHandle === averifierHandle ? 'warning' : 'success';
                let destName = targetHandle === averifierHandle ? '⚠️ A vérifier' : targetHandle.name;
                
                if (matchedTitre) {
                    log(`Routage (Code: ${usedCode}) : ${cleanFinalName} ➡️ ${destName}`, statusLog);
                } else {
                    log(`Routage (Sans match) : ${cleanFinalName} ➡️ ${destName}`, statusLog);
                }
                
                fileCount++;
            } catch (err) {
                log(`Erreur sur le fichier "${cleanFinalName}" : ${err.message}`, 'error');
            }
        }
        
        log(`✅ Importation terminée : ${fileCount} fichiers classés.`, 'success');
        statusZip.textContent = "Traitement terminé avec succès.";
        statusZip.className = "status-indicator success";
        
    } catch (err) { 
        log(`Erreur ZIP globale : ${err.message}`, 'error'); 
        statusZip.textContent = "Erreur de traitement.";
        statusZip.className = "status-indicator error";
    }
    
    // Reset l'input
    e.target.value = '';
});

// NOUVEL ALGORITHME DE ROUTAGE AVANCÉ
function getTargetHandle(filename) {
    const fn = filename.toLowerCase();
    
    // 5. CONTRÔLE QUALITÉ, ESSAIS ET TRAÇABILITÉ
    if (/\b(bl|livraison|bon)\b/.test(fn)) return folderMapping["5.4"] || averifierHandle;
    if (/\b(essai|ecrasement|eprouvette|labo)\b/.test(fn)) return folderMapping["5.1"] || averifierHandle;
    if (/\b(fnc|non conformite|ecart|fad|fqr|fdr)\b/.test(fn)) return folderMapping["5.3"] || averifierHandle;
    
    // 4. FICHES TECHNIQUES ET CERTIFICATS
    // Bétons
    if (/\b(formulation|mix|beton|convenance)\b/.test(fn)) return folderMapping["4.1"] || averifierHandle;
    // Fournitures (DAF, DAG, Fiches techniques diverses)
    if (/\b(daf|dag|fourniture|fiche technique|ft)\b/.test(fn)) return folderMapping["4.2"] || folderMapping["4.1"] || averifierHandle;
    
    // 3. NOTES DE CALCUL ET DOCUMENTS TECHNIQUES MÉTHODES (NDC)
    if (/\b(pex|prc|fiche de tache)\b/.test(fn)) return folderMapping["3.4"] || folderMapping["3"] || averifierHandle;
    if (/\b(soutenement|fondation)\b/.test(fn)) return folderMapping["3.3"] || folderMapping["3"] || averifierHandle;
    if (/\b(geotechnique|sol)\b/.test(fn)) return folderMapping["3.2"] || folderMapping["3"] || averifierHandle;
    if (/\b(ndc|calcul|note de calcul)\b/.test(fn)) return folderMapping["3.1"] || folderMapping["3"] || averifierHandle;

    // 2. PLANS DE RÉCOLEMENT (PLA)
    if (/\b(coffrage)\b/.test(fn)) return folderMapping["2.2"] || folderMapping["2"] || averifierHandle;
    if (/\b(ferraillage|armature)\b/.test(fn)) return folderMapping["2.3"] || folderMapping["2"] || averifierHandle;
    if (/\b(etancheite)\b/.test(fn)) return folderMapping["2.4"] || folderMapping["2"] || averifierHandle;
    if (/\b(masse|situation|implantation)\b/.test(fn)) return folderMapping["2.1"] || folderMapping["2"] || averifierHandle;
    if (/\b(pla|plan)\b/.test(fn)) return folderMapping["2"] || averifierHandle;

    // FALLBACK : "A vérifier"
    return averifierHandle;
}
