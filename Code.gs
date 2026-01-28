const SHEET_NAME = 'Matériel_exposants';
const CACHE_DURATION = 300; // 5 minutes en secondes

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gestion Exposants')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🔹 Retourne tous les exposants en JSON avec CACHE
function getExposants() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('exposants_data');
    
    if (cachedData) {
      Logger.log("Données récupérées depuis le cache");
      return JSON.parse(cachedData);
    }
    
    Logger.log("Récupération depuis le Sheet");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return [];
    }

    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = allData[0];
    
    let idxNom = -1;
    for (let j = 0; j < headers.length; j++) {
      if (headers[j] && headers[j].toString().trim() === 'Nom Exposant') {
        idxNom = j;
        break;
      }
    }
    
    if (idxNom < 0) {
      return [];
    }
    
    const exposants = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const nom = row[idxNom];
      
      if (!nom || nom.toString().trim() === '') {
        continue;
      }
      
      const obj = {};
      
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j] ? headers[j].toString().trim() : '';
        if (!header) continue;
        
        const value = row[j];
        
        if (value === null || value === undefined || value === '') {
          obj[header] = '';
        } else if (value instanceof Date) {
          obj[header] = value.toISOString();
        } else if (typeof value === 'number') {
          obj[header] = value;
        } else {
          obj[header] = value.toString();
        }
      }
      
      exposants.push(obj);
    }
    
    try {
      cache.put('exposants_data', JSON.stringify(exposants), CACHE_DURATION);
      Logger.log("Données mises en cache pour " + CACHE_DURATION + " secondes");
    } catch (e) {
      Logger.log("Erreur lors de la mise en cache: " + e.toString());
    }
    
    return exposants;
    
  } catch (error) {
    Logger.log("ERREUR: " + error.toString());
    return [];
  }
}

// 🔹 Enregistrer ARRIVÉE ou DÉPART avec HEURE UNIQUEMENT (sans date)
function saveAction(nom, type, signature, commentaire) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return false;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = allData[0];

    // Trouver les index
    let idxNom = -1;
    let idxArrCom = -1;
    let idxArrDate = -1;
    let idxDepCom = -1;
    let idxDepDate = -1;
    
    for (let j = 0; j < headers.length; j++) {
      const h = headers[j] ? headers[j].toString().trim() : '';
      if (h === 'Nom Exposant') idxNom = j;
      else if (h === 'Commentaire_Arrivee') idxArrCom = j;
      else if (h === 'Heure_Arrivee') idxArrDate = j;
      else if (h === 'Commentaire_Depart') idxDepCom = j;
      else if (h === 'Heure_Depart') idxDepDate = j;
    }

    if (idxNom < 0) {
      return false;
    }

    const now = new Date();
    const nomTrim = nom.toString().trim();

    // Trouver la ligne
    for (let i = 1; i < allData.length; i++) {
      const rowNom = allData[i][idxNom];
      
      if (rowNom && rowNom.toString().trim() === nomTrim) {
        Logger.log("Exposant trouvé à la ligne " + (i + 1));
        
        if (type === 'arrivee') {
          // Commentaire
          if (idxArrCom >= 0) {
            sheet.getRange(i + 1, idxArrCom + 1).setValue(commentaire || '');
          }
          // Heure - IMPORTANT: Écrire la valeur ET formater la cellule
          if (idxArrDate >= 0) {
            const cell = sheet.getRange(i + 1, idxArrDate + 1);
            cell.setValue(now);
            // FORCER le format HEURE SEULEMENT (HH:MM)
            cell.setNumberFormat('HH:mm');
          }
        } else if (type === 'depart') {
          // Commentaire
          if (idxDepCom >= 0) {
            sheet.getRange(i + 1, idxDepCom + 1).setValue(commentaire || '');
          }
          // Heure - IMPORTANT: Écrire la valeur ET formater la cellule
          if (idxDepDate >= 0) {
            const cell = sheet.getRange(i + 1, idxDepDate + 1);
            cell.setValue(now);
            // FORCER le format HEURE SEULEMENT (HH:MM)
            cell.setNumberFormat('HH:mm');
          }
        }
        
        // Invalider le cache
        const cache = CacheService.getScriptCache();
        cache.remove('exposants_data');
        
        Logger.log("Enregistrement réussi avec format HH:mm");
        return true;
      }
    }

    Logger.log("Exposant non trouvé");
    return false;
    
  } catch (error) {
    Logger.log("ERREUR saveAction: " + error.toString());
    return false;
  }
}

// 🔹 Fonction pour formater TOUTES les colonnes d'heure existantes
function formatAllTimeColumns() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log("Feuille non trouvée");
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Trouver les colonnes d'heure
    let idxArrDate = -1;
    let idxDepDate = -1;
    
    for (let j = 0; j < headers.length; j++) {
      const h = headers[j] ? headers[j].toString().trim() : '';
      if (h === 'Heure_Arrivee') idxArrDate = j;
      else if (h === 'Heure_Depart') idxDepDate = j;
    }
    
    let formatted = 0;
    
    // Formater la colonne Heure_Arrivee
    if (idxArrDate >= 0) {
      const range = sheet.getRange(2, idxArrDate + 1, lastRow - 1, 1);
      range.setNumberFormat('HH:mm');
      formatted++;
      Logger.log("Colonne Heure_Arrivee formatée");
    }
    
    // Formater la colonne Heure_Depart
    if (idxDepDate >= 0) {
      const range = sheet.getRange(2, idxDepDate + 1, lastRow - 1, 1);
      range.setNumberFormat('HH:mm');
      formatted++;
      Logger.log("Colonne Heure_Depart formatée");
    }
    
    Logger.log("=== FORMATAGE TERMINÉ ===");
    Logger.log(formatted + " colonne(s) formatée(s) en HH:mm");
    Logger.log("Toutes les heures existantes et futures seront au format HH:mm");
    
  } catch (error) {
    Logger.log("ERREUR: " + error.toString());
  }
}

// 🔹 Fonction pour vider manuellement le cache
function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('exposants_data');
  Logger.log("Cache vidé");
}

// 🔹 Fonction de test
function testGetExposants() {
  clearCache();
  
  const start = new Date().getTime();
  const exposants = getExposants();
  const end = new Date().getTime();
  
  Logger.log("=== TEST ===");
  Logger.log("Temps d'exécution: " + (end - start) + " ms");
  Logger.log("Nombre: " + exposants.length);
  
  const start2 = new Date().getTime();
  const exposants2 = getExposants();
  const end2 = new Date().getTime();
  
  Logger.log("Temps avec cache: " + (end2 - start2) + " ms");
}
