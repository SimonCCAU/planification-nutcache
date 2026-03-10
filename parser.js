/**
 * parser.js — Parse les rapports Nutcache (CSV ou XLSX)
 * 
 * Format CSV Nutcache (dénormalisé) : chaque ligne contient toutes les infos
 * Colonnes par ligne :
 *   0: "Grand total"
 *   1: heures grand total
 *   2: heures facturables grand total  
 *   3: 0
 *   4: coûts grand total
 *   5: montants grand total
 *   6: "Client - Code - Nom du projet"
 *   7: heures projet
 *   8: "Total pour le projet"
 *   9: coûts projet
 *   10: montants projet
 *   11: heures facturables projet
 *   12: 0
 *   13: Nom du membre
 *   14-19: headers ("Service", "Heures travaillées", ...)
 *   20: "Total pour le membre"
 *   21: heures membre
 *   22: heures facturées membre
 *   23: heures facturables membre
 *   24: coûts membre
 *   25: montants membre
 *   26: Nom du service
 *   27: heures service
 *   28: heures facturées service
 *   29: heures facturables service
 *   30: coûts service
 *   31: montants service
 */

function parseNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  // Gérer format français: "1 234,56" → 1234.56
  let s = String(val).trim();
  s = s.replace(/\s/g, "");     // retirer espaces (séparateur milliers)
  s = s.replace(",", ".");       // virgule → point décimal
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseNutcacheReport(workbookOrData) {
  let rows;
  
  if (workbookOrData._csvRows) {
    // CSV parsé manuellement
    rows = workbookOrData._csvRows;
  } else if (workbookOrData.SheetNames) {
    // Format XLSX via SheetJS
    const sheet = workbookOrData.Sheets[workbookOrData.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  } else {
    rows = workbookOrData;
  }

  if (!rows || rows.length === 0) {
    return { periode: { debut: null, fin: null }, projets: [] };
  }

  // Détecter le format : CSV dénormalisé ou XLSX hiérarchique
  const firstRow = rows[0];
  const isFlat = String(firstRow[0] || "").trim() === "Grand total";

  if (isFlat) {
    return parseFlatCSV(rows);
  } else {
    return parseHierarchicalXLSX(rows);
  }
}

/**
 * Parse le format CSV dénormalisé de Nutcache
 */
function parseFlatCSV(rows) {
  const projetsMap = {};
  
  for (const row of rows) {
    if (!row || row.length < 27) continue;
    
    const projetTitle = String(row[6] || "").trim();
    const membreNom = String(row[13] || "").trim();
    const serviceNom = String(row[26] || "").trim().replace(/-$/, "").trim();
    
    if (!projetTitle || !membreNom || !serviceNom) continue;
    
    // Parser le titre du projet
    const projetMatch = projetTitle.match(/^(.+?)\s*-\s*(\d{5})\s*-\s*(.+)$/);
    if (!projetMatch) continue;
    
    const client = projetMatch[1].trim();
    const code = projetMatch[2].trim();
    const nom = projetMatch[3].trim();
    
    // Détection type de projet
    const fullTitle = projetTitle.toLowerCase();
    let categorie = "Facturable";
    let tabColor = "#077C79";
    if (fullTitle.includes("consultant")) {
      categorie = "Consultant municipal";
      tabColor = "#EEDB4C";
    } else if (fullTitle.includes("interne") || client.toLowerCase() === "interne") {
      categorie = "Non-facturable";
      tabColor = "#EB5E63";
    }
    
    // Données du service
    const serviceHeures = parseNumber(row[27]);
    const serviceFacturees = parseNumber(row[28]);
    const serviceFacturables = parseNumber(row[29]);
    const serviceCouts = parseNumber(row[30]);
    const serviceMontants = parseNumber(row[31]);
    
    // Données du membre (totaux)
    const membreHeures = parseNumber(row[21]);
    const membreCouts = parseNumber(row[24]);
    const membreMontants = parseNumber(row[25]);
    
    // Données du projet (totaux)
    const projetHeures = parseNumber(row[7]);
    const projetCouts = parseNumber(row[9]);
    const projetMontants = parseNumber(row[10]);
    
    // Créer ou mettre à jour le projet
    if (!projetsMap[code]) {
      projetsMap[code] = {
        client: client,
        code: code,
        nom: nom,
        categorie: categorie,
        tabColor: tabColor,
        totalHeures: projetHeures,
        totalCouts: projetCouts,
        totalFacturables: projetMontants,
        membresMap: {}
      };
    }
    
    const projet = projetsMap[code];
    
    // Créer ou mettre à jour le membre
    if (!projet.membresMap[membreNom]) {
      projet.membresMap[membreNom] = {
        nom: membreNom,
        heures: membreHeures,
        couts: membreCouts,
        montantsFacturables: membreMontants,
        tauxDeduit: membreHeures > 0 ? Math.round((membreCouts / membreHeures) * 100) / 100 : 0,
        services: []
      };
    }
    
    // Ajouter le service
    projet.membresMap[membreNom].services.push({
      nom: serviceNom,
      heures: serviceHeures,
      couts: serviceCouts,
      facturables: serviceMontants
    });
  }
  
  // Convertir en tableau
  const projets = Object.values(projetsMap).map(p => ({
    client: p.client,
    code: p.code,
    nom: p.nom,
    categorie: p.categorie,
    tabColor: p.tabColor,
    totalHeures: p.totalHeures,
    totalCouts: p.totalCouts,
    totalFacturables: p.totalFacturables,
    membres: Object.values(p.membresMap)
  }));
  
  return {
    periode: { debut: null, fin: null },  // Le CSV ne contient pas la période
    projets: projets
  };
}

/**
 * Parse le format XLSX hiérarchique de Nutcache (ancien format)
 */
function parseHierarchicalXLSX(rows) {
  let reportDateDebut = null;
  let reportDateFin = null;
  const periodCell = String(rows[1] && rows[1][3] || "").trim();
  const periodMatch = periodCell.match(/(\d{4}-\d{2}-\d{2})\s*(?:à|to|-)\s*(\d{4}-\d{2}-\d{2})/);
  if (periodMatch) {
    reportDateDebut = periodMatch[1];
    reportDateFin = periodMatch[2];
  }

  const projets = [];
  let currentProjet = null;
  let currentMembre = null;
  let inServiceHeader = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const cellA = String(row[0] || "").trim();

    if (!cellA && !row[1] && !row[2]) continue;
    if (cellA.startsWith("Devise utilisée")) continue;
    if (cellA === "Grand total") break;

    const projetMatch = cellA.match(/^(.+?)\s*-\s*(\d{5})\s*-\s*(.+)$/);
    if (projetMatch) {
      const client = projetMatch[1].trim();
      const nom = projetMatch[3].trim();
      const fullTitle = cellA.toLowerCase();
      let categorie = "Facturable";
      let tabColor = "#077C79";
      if (fullTitle.includes("consultant")) {
        categorie = "Consultant municipal"; tabColor = "#EEDB4C";
      } else if (fullTitle.includes("interne") || client.toLowerCase() === "interne") {
        categorie = "Non-facturable"; tabColor = "#EB5E63";
      }
      currentProjet = {
        client: client, code: projetMatch[2].trim(), nom: nom,
        categorie: categorie, tabColor: tabColor,
        totalHeures: 0, totalCouts: 0, totalFacturables: 0, membres: []
      };
      projets.push(currentProjet);
      currentMembre = null; inServiceHeader = false;
      continue;
    }

    if (cellA.startsWith("Total pour le projet") && currentProjet) {
      currentProjet.totalHeures = parseFloat(row[1]) || 0;
      currentProjet.totalCouts = parseFloat(row[2]) || 0;
      currentProjet.totalFacturables = parseFloat(row[4]) || 0;
      currentMembre = null;
      continue;
    }

    if (cellA.startsWith("Total pour le membre") && currentMembre) {
      currentMembre.heures = parseFloat(row[1]) || 0;
      currentMembre.couts = parseFloat(row[2]) || 0;
      currentMembre.montantsFacturables = parseFloat(row[4]) || 0;
      currentMembre.tauxDeduit = currentMembre.heures > 0
        ? Math.round((currentMembre.couts / currentMembre.heures) * 100) / 100 : 0;
      inServiceHeader = false;
      continue;
    }

    if (cellA === "Service") { inServiceHeader = true; continue; }

    if (inServiceHeader && currentMembre && cellA && row[1] !== "") {
      currentMembre.services.push({
        nom: cellA.replace(/-$/, "").trim(),
        heures: parseFloat(row[1]) || 0,
        couts: parseFloat(row[2]) || 0,
        facturables: parseFloat(row[4]) || 0
      });
      continue;
    }

    if (currentProjet && cellA && !cellA.startsWith("Total") && cellA !== "Service") {
      const nextRow = rows[i + 1] || [];
      if (String(nextRow[0] || "").trim() === "Service") {
        currentMembre = {
          nom: cellA, heures: 0, couts: 0, montantsFacturables: 0, tauxDeduit: 0, services: []
        };
        currentProjet.membres.push(currentMembre);
        inServiceHeader = false;
        continue;
      }
    }
  }

  return { periode: { debut: reportDateDebut, fin: reportDateFin }, projets: projets };
}

/**
 * Lit un fichier (XLSX ou CSV) et retourne les données parsées
 */
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const isCSV = file.name.toLowerCase().endsWith(".csv");
    
    if (isCSV) {
      // Parser CSV manuellement — pas de SheetJS pour éviter les problèmes de conversion
      const textReader = new FileReader();
      textReader.onload = function(e) {
        try {
          const text = e.target.result;
          const rows = parseCSVText(text);
          // Retourner un faux workbook compatible avec parseNutcacheReport
          resolve({ _csvRows: rows });
        } catch(err) {
          reject(new Error("Erreur CSV: " + err.message));
        }
      };
      textReader.onerror = () => reject(new Error("Erreur lecture CSV"));
      textReader.readAsText(file, "utf-8");
    } else {
      // XLSX via SheetJS
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", raw: true });
          resolve(workbook);
        } catch(err) {
          reject(new Error("Erreur XLSX: " + err.message));
        }
      };
      reader.onerror = () => reject(new Error("Erreur lecture fichier"));
      reader.readAsArrayBuffer(file);
    }
  });
}

/**
 * Parse un texte CSV en respectant les guillemets et virgules dans les valeurs
 */
function parseCSVText(text) {
  const rows = [];
  const lines = text.split(/\r?\n/);
  
  for (const line of lines) {
    if (!line.trim()) continue;
    const fields = [];
    let current = "";
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          fields.push(current);
          current = "";
        } else {
          current += ch;
        }
      }
    }
    fields.push(current);
    rows.push(fields);
  }
  
  return rows;
}
