/**
 * updater.js — v2 — Marqueurs texte + plages dynamiques
 * Cherche les sections par leur label au lieu d'index de lignes fixes.
 */

const TEMPLATE_SHEET = "_TEMPLATE";
const INDEX_SHEET = "Index Projets";
const DASHBOARD_SHEET = "Tableau de bord";

// Marqueurs texte à chercher dans les onglets projets
const MARKERS = {
  AVANCEMENT: "AVANCEMENT",
  ALLOCATION: "ALLOCATION DES RESSOURCES",
  PLANIF_CAP: "PLANIFICATION CAPACITAIRE DU PROJET",
  PHASES: "DÉTAIL PAR PHASE / SERVICE",
  NOTES: "NOTES DE SUIVI",
  TOTAL: "TOTAL",
  PERIODE: "Période Nutcache"
};

// ============================================================
// UTILITAIRES
// ============================================================

function log(msg) {
  const el = document.getElementById("log");
  if (el) {
    const line = document.createElement("div");
    line.textContent = `[${new Date().toLocaleTimeString("fr-CA")}] ${msg}`;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }
  console.log(msg);
}

function logError(msg) {
  const el = document.getElementById("log");
  if (el) {
    const line = document.createElement("div");
    line.className = "log-error";
    line.textContent = `⚠ ${msg}`;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }
  console.error(msg);
}

/**
 * Cherche un marqueur texte dans la colonne A d'un worksheet.
 * Retourne le numéro de ligne (1-based) ou null si non trouvé.
 */
async function findMarker(ws, context, markerText) {
  const usedRange = ws.getUsedRange();
  usedRange.load("values, rowCount, columnCount");
  await context.sync();

  for (let r = 0; r < usedRange.values.length; r++) {
    const cellA = String(usedRange.values[r][0] || "").trim();
    if (cellA === markerText || cellA.startsWith(markerText)) {
      return r + 1; // 1-based
    }
  }
  return null;
}

/**
 * Cherche un marqueur texte partiel dans la colonne A.
 */
async function findMarkerPartial(ws, context, partialText) {
  const usedRange = ws.getUsedRange();
  usedRange.load("values");
  await context.sync();

  for (let r = 0; r < usedRange.values.length; r++) {
    const cellA = String(usedRange.values[r][0] || "").trim();
    if (cellA.includes(partialText)) {
      return r + 1;
    }
  }
  return null;
}

/**
 * Lit toutes les valeurs de la colonne A pour un worksheet.
 * Retourne un objet avec les positions de tous les marqueurs.
 */
async function findAllMarkers(ws, context) {
  // Scanner les 100 premières lignes de la colonne A
  const scanRange = ws.getRangeByIndexes(0, 0, 100, 1);
  scanRange.load("values");
  await context.sync();

  const positions = {};
  for (let r = 0; r < scanRange.values.length; r++) {
    const cellA = String(scanRange.values[r][0] || "").trim();
    if (!cellA) continue;
    for (const [key, text] of Object.entries(MARKERS)) {
      if (cellA === text || cellA.startsWith(text)) {
        if (!positions[key]) positions[key] = r + 1;
      }
    }
  }
  return positions;
}

/**
 * Trouve la ligne "TOTAL" après une position donnée dans la colonne A.
 */
async function findTotalAfter(ws, context, afterRow) {
  const usedRange = ws.getUsedRange();
  usedRange.load("values");
  await context.sync();

  for (let r = afterRow; r < usedRange.values.length; r++) {
    if (String(usedRange.values[r][0] || "").trim() === "TOTAL") {
      return r + 1;
    }
  }
  return null;
}

/**
 * Récupère les codes projets existants dans l'Index.
 */
async function getExistingProjectCodes(context) {
  const indexSheet = context.workbook.worksheets.getItemOrNullObject(INDEX_SHEET);
  await context.sync();
  if (indexSheet.isNullObject) return [];

  const usedRange = indexSheet.getUsedRange();
  usedRange.load("values");
  await context.sync();

  const codes = [];
  for (let i = 0; i < usedRange.values.length; i++) {
    const code = String(usedRange.values[i][0] || "").trim();
    // Ignorer headers et labels
    if (code && /^\d{5}$|^[A-Z]{2,3}-\d{3}$/.test(code)) {
      codes.push(code);
    }
  }
  return codes;
}

async function getExistingSheetNames(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  return sheets.items.map(s => s.name);
}

// ============================================================
// LECTURE STRUCTURÉE D'UN ONGLET PROJET
// ============================================================

/**
 * Lit la structure d'un onglet projet en se basant sur les marqueurs.
 * Retourne un objet avec toutes les positions clés.
 */
async function readProjectStructure(ws, context) {
  const markers = await findAllMarkers(ws, context);
  
  // Position de la période — essayer avec et sans accent
  let periodeRow = await findMarkerPartial(ws, context, "Période Nutcache");
  if (!periodeRow) periodeRow = await findMarkerPartial(ws, context, "riode Nutcache");
  if (!periodeRow) periodeRow = await findMarkerPartial(ws, context, "Periode");
  
  // Position de l'allocation : header = ALLOCATION row + 1, data = header + 1
  let allocHeaderRow = null;
  let allocDataStart = null;
  let allocTotalRow = null;
  if (markers.ALLOCATION) {
    allocHeaderRow = markers.ALLOCATION + 1;
    allocDataStart = markers.ALLOCATION + 2;
    // Trouver le TOTAL de l'allocation (premier TOTAL après ALLOCATION)
    allocTotalRow = await findTotalAfter(ws, context, markers.ALLOCATION);
  }

  // Planification capacitaire
  let planifHeaderRow = null;
  let planifDataStart = null;
  let planifTotalRow = null;
  if (markers.PLANIF_CAP) {
    // Header est 2 lignes après le titre (titre, sous-titre, header)
    planifHeaderRow = markers.PLANIF_CAP + 2;
    planifDataStart = markers.PLANIF_CAP + 3;
    planifTotalRow = await findTotalAfter(ws, context, markers.PLANIF_CAP);
  }

  // Phases
  let phasesHeaderRow = null;
  let phasesDataStart = null;
  if (markers.PHASES) {
    phasesHeaderRow = markers.PHASES + 1;
    phasesDataStart = markers.PHASES + 2;
  }

  return {
    markers,
    periodeRow,
    allocHeaderRow,
    allocDataStart,
    allocTotalRow,
    allocMaxRows: allocTotalRow ? allocTotalRow - allocDataStart : 20,
    planifHeaderRow,
    planifDataStart,
    planifTotalRow,
    phasesHeaderRow,
    phasesDataStart
  };
}

// ============================================================
// MISE À JOUR D'UN PROJET EXISTANT
// ============================================================

async function updateExistingProject(context, projet, reportPeriode) {
  log(`↻ Mise à jour : ${projet.code} — ${projet.nom}`);

  const ws = context.workbook.worksheets.getItem(projet.code);
  const struct = await readProjectStructure(ws, context);

  if (!struct.allocDataStart) {
    logError(`Structure non trouvée pour ${projet.code} — marqueur ALLOCATION manquant`);
    return;
  }

  // --- LIRE LA PÉRIODE EXISTANTE ---
  let existingDebut = "";
  let existingFin = "";
  if (struct.periodeRow) {
    const periodeRange = ws.getRangeByIndexes(struct.periodeRow - 1, 1, 1, 2);
    periodeRange.load("values");
    await context.sync();
    existingDebut = String(periodeRange.values[0][0] || "").trim();
    existingFin = String(periodeRange.values[0][1] || "").trim();
  }

  // --- DÉTERMINER LE MODE ---
  let mode = "ECRASEMENT";
  let newDebut = reportPeriode.debut || "";
  let newFin = reportPeriode.fin || "";

  if (existingDebut && existingFin && reportPeriode.debut && reportPeriode.fin) {
    const exDeb = new Date(existingDebut);
    const exFin = new Date(existingFin);
    const rpDeb = new Date(reportPeriode.debut);
    const rpFin = new Date(reportPeriode.fin);

    if (rpDeb <= exDeb && rpFin >= exFin) {
      mode = "ECRASEMENT";
      newDebut = reportPeriode.debut;
      newFin = reportPeriode.fin;
    } else if (rpDeb > exFin) {
      mode = "AJOUT";
      newDebut = existingDebut;
      newFin = reportPeriode.fin;
    } else if (rpFin < exDeb) {
      mode = "AJOUT";
      newDebut = reportPeriode.debut;
      newFin = existingFin;
    } else {
      mode = "ECRASEMENT";
      newDebut = rpDeb < exDeb ? reportPeriode.debut : existingDebut;
      newFin = rpFin > exFin ? reportPeriode.fin : existingFin;
      log(`  ⚠ Chevauchement partiel — écrasement appliqué`);
    }
  }

  log(`  Mode: ${mode} | Période: ${newDebut} → ${newFin}`);

  // --- METTRE À JOUR LA PÉRIODE ---
  if (struct.periodeRow) {
    ws.getRangeByIndexes(struct.periodeRow - 1, 1, 1, 1).values = [[newDebut]];
    ws.getRangeByIndexes(struct.periodeRow - 1, 2, 1, 1).values = [[newFin]];
  }

  // --- LIRE LES RESSOURCES EXISTANTES ---
  const maxRows = struct.allocMaxRows;
  const allocRange = ws.getRangeByIndexes(struct.allocDataStart - 1, 0, maxRows, 5);
  allocRange.load("values");
  await context.sync();

  const existingNames = {};
  let nextEmptyRow = struct.allocDataStart;
  for (let i = 0; i < allocRange.values.length; i++) {
    const name = String(allocRange.values[i][0] || "").trim();
    if (name && name !== "TOTAL") {
      existingNames[name] = {
        row: struct.allocDataStart + i,
        heures: parseFloat(allocRange.values[i][2]) || 0,
        couts: parseFloat(allocRange.values[i][3]) || 0
      };
      nextEmptyRow = struct.allocDataStart + i + 1;
    }
  }

  if (mode === "ECRASEMENT") {
    for (let i = 0; i < maxRows; i++) {
      const r = struct.allocDataStart - 1 + i;
      ws.getRangeByIndexes(r, 2, 1, 1).values = [[0]]; // C
      ws.getRangeByIndexes(r, 3, 1, 1).values = [[0]]; // D
      ws.getRangeByIndexes(r, 4, 1, 1).values = [[0]]; // E
    }
  }

  // --- INJECTER DONNÉES PAR MEMBRE ---
  for (const membre of projet.membres) {
    const existing = existingNames[membre.nom];
    let targetRow;

    if (existing) {
      targetRow = existing.row;
    } else {
      if (nextEmptyRow >= (struct.allocTotalRow || struct.allocDataStart + 20)) {
        logError(`Plus de place pour ${membre.nom} dans ${projet.code}`);
        continue;
      }
      targetRow = nextEmptyRow;
      ws.getRangeByIndexes(targetRow - 1, 0, 1, 1).values = [[membre.nom]];
      nextEmptyRow++;
    }

    const rowIdx = targetRow - 1;
    if (mode === "AJOUT" && existing) {
      const nh = existing.heures + membre.heures;
      const nc = existing.couts + membre.couts;
      const nt = nh > 0 ? Math.round((nc / nh) * 100) / 100 : 0;
      ws.getRangeByIndexes(rowIdx, 2, 1, 1).values = [[nh]];
      ws.getRangeByIndexes(rowIdx, 3, 1, 1).values = [[nc]];
      ws.getRangeByIndexes(rowIdx, 4, 1, 1).values = [[nt]];
      log(`  + ${membre.nom}: ${existing.heures}h + ${membre.heures}h = ${nh}h`);
    } else {
      ws.getRangeByIndexes(rowIdx, 2, 1, 1).values = [[membre.heures]];
      ws.getRangeByIndexes(rowIdx, 3, 1, 1).values = [[membre.couts]];
      ws.getRangeByIndexes(rowIdx, 4, 1, 1).values = [[membre.tauxDeduit]];
    }
  }

  // --- PHASES ---
  await updatePhases(context, ws, projet, struct);

  await context.sync();
  log(`✓ ${projet.code} mis à jour [${mode}] (${projet.membres.length} membres)`);
}

// ============================================================
// CRÉATION D'UN NOUVEAU PROJET
// ============================================================

async function createNewProject(context, projet, reportPeriode) {
  log(`✚ Nouveau projet : ${projet.code} — ${projet.nom}`);

  const templateSheet = context.workbook.worksheets.getItem(TEMPLATE_SHEET);
  const newSheet = templateSheet.copy("End");
  await context.sync();

  newSheet.name = projet.code;
  newSheet.visibility = "Visible";
  newSheet.tabColor = projet.tabColor || "#077C79";

  // Lire la structure du template copié
  const struct = await readProjectStructure(newSheet, context);

  // Titre — chercher {{CODE}} dans col A et remplacer
  let titleRow = await findMarkerPartial(newSheet, context, "{{CODE}}");
  if (!titleRow) titleRow = await findMarkerPartial(newSheet, context, "CODE");
  if (!titleRow) titleRow = 4; // fallback ligne 4
  
  log(`  Titre sur ligne ${titleRow}`);
  newSheet.getRangeByIndexes(titleRow - 1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];
  // Appliquer couleur de fond du titre selon le type
  const titleRange = newSheet.getRangeByIndexes(titleRow - 1, 0, 1, 7);
  titleRange.format.fill.color = projet.tabColor || "#077C79";
  const fontColor = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";
  titleRange.format.font.color = fontColor;
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  // Effacer R1 si le titre s'y est copié par erreur
  const r1val = newSheet.getRangeByIndexes(0, 0, 1, 1);
  r1val.load("values");
  await context.sync();
  if (String(r1val.values[0][0] || "").includes("—") || String(r1val.values[0][0] || "").includes("CODE")) {
    newSheet.getRangeByIndexes(0, 0, 1, 1).values = [[""]];
  }
  // En-tête — chercher chaque label dans col A et écrire la valeur en col B
  const fieldMap = {
    "Client": projet.client,
    "Catégorie": projet.categorie || "Facturable",
  };
  
  const usedRange = newSheet.getUsedRange();
  usedRange.load("values");
  await context.sync();
  
  for (let r = 0; r < usedRange.values.length; r++) {
    const cellA = String(usedRange.values[r][0] || "").trim();
    if (fieldMap[cellA] !== undefined) {
      newSheet.getRangeByIndexes(r, 1, 1, 1).values = [[fieldMap[cellA]]];
    }
  }

  // Période
  if (struct.periodeRow) {
    newSheet.getRangeByIndexes(struct.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut || ""]];
    newSheet.getRangeByIndexes(struct.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin || ""]];
  }

  // Appliquer couleurs des en-têtes de tableaux selon le type de projet
  await applyProjectStyling(newSheet, context, struct, projet);

  // Membres dans allocation
  if (struct.allocDataStart) {
    for (let i = 0; i < projet.membres.length && i < 20; i++) {
      const membre = projet.membres[i];
      const r = struct.allocDataStart - 1 + i;
      newSheet.getRangeByIndexes(r, 0, 1, 1).values = [[membre.nom]];
      newSheet.getRangeByIndexes(r, 1, 1, 1).values = [[0]];
      newSheet.getRangeByIndexes(r, 2, 1, 1).values = [[membre.heures]];
      newSheet.getRangeByIndexes(r, 3, 1, 1).values = [[membre.couts]];
      newSheet.getRangeByIndexes(r, 4, 1, 1).values = [[membre.tauxDeduit]];
    }

    if (struct.planifDataStart) {
      for (let i = 0; i < projet.membres.length && i < 20; i++) {
        newSheet.getRangeByIndexes(struct.planifDataStart - 1 + i, 0, 1, 1).values = [[projet.membres[i].nom]];
      }
    }
  }

  // Phases
  await updatePhases(context, newSheet, projet, struct);

  // Index et Dashboard
  await addToIndex(context, projet);
  await addToDashboard(context, projet);

  await context.sync();
  log(`✓ ${projet.code} créé avec succès`);
}

// ============================================================
// APPLIQUER LE STYLE COULEUR PAR TYPE DE PROJET
// ============================================================

async function applyProjectStyling(ws, context, struct, projet) {
  const accent = projet.tabColor || "#077C79";
  const fontOnAccent = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";

  // Couleur des en-têtes d'avancement (Heures / $)
  if (struct.markers.AVANCEMENT) {
    const avRow = struct.markers.AVANCEMENT + 1;
    const avHeader = ws.getRangeByIndexes(avRow - 1, 1, 1, 2);
    avHeader.format.fill.color = accent;
    avHeader.format.font.color = fontOnAccent;
  }

  // Couleur en-tête allocation
  if (struct.allocHeaderRow) {
    const allocHeader = ws.getRangeByIndexes(struct.allocHeaderRow - 1, 0, 1, 7);
    allocHeader.format.fill.color = accent;
    allocHeader.format.font.color = fontOnAccent;
  }

  // Couleur en-tête planification capacitaire
  if (struct.planifHeaderRow) {
    const planifHeader = ws.getRangeByIndexes(struct.planifHeaderRow - 1, 0, 1, 14);
    planifHeader.format.fill.color = accent;
    planifHeader.format.font.color = fontOnAccent;
  }

  // Couleur en-tête phases
  if (struct.phasesHeaderRow) {
    const phasesHeader = ws.getRangeByIndexes(struct.phasesHeaderRow - 1, 0, 1, 5);
    phasesHeader.format.fill.color = accent;
    phasesHeader.format.font.color = fontOnAccent;
  }

  // Couleur en-tête notes
  const notesRow = await findMarkerPartial(ws, context, "NOTES DE SUIVI");
  if (notesRow) {
    const notesHeader = ws.getRangeByIndexes(notesRow, 0, 1, 3);
    notesHeader.format.fill.color = accent;
    notesHeader.format.font.color = fontOnAccent;
  }

  await context.sync();
  log(`  → Style ${projet.categorie} appliqué (${accent})`);
}

// ============================================================
// PHASES
// ============================================================

async function updatePhases(context, ws, projet, struct) {
  const phaseRows = [];
  for (const membre of projet.membres) {
    for (const service of membre.services) {
      phaseRows.push([service.nom, membre.nom, service.heures, service.couts, service.facturables]);
    }
  }

  if (!struct.phasesDataStart) return;

  // Effacer 30 lignes à partir du début des données phases
  const clearRange = ws.getRangeByIndexes(struct.phasesDataStart - 1, 0, 30, 5);
  clearRange.clear("Contents");

  if (phaseRows.length > 0) {
    const phaseRange = ws.getRangeByIndexes(struct.phasesDataStart - 1, 0, phaseRows.length, 5);
    phaseRange.values = phaseRows;
  }
}

// ============================================================
// INDEX PROJETS
// ============================================================

async function addToIndex(context, projet) {
  const indexSheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  const usedRange = indexSheet.getUsedRange();
  usedRange.load("rowCount");
  await context.sync();

  const newRow = usedRange.rowCount;
  const data = [
    projet.code, projet.nom, projet.client,
    projet.categorie || "Facturable", "En cours", "",
    "", "", 0, 0,
    "Importé depuis Nutcache"
  ];
  indexSheet.getRangeByIndexes(newRow, 0, 1, data.length).values = [data];
  log(`  → Ajouté à l'Index (ligne ${newRow + 1})`);
}

// ============================================================
// TABLEAU DE BORD
// ============================================================

async function addToDashboard(context, projet) {
  const dashSheet = context.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await context.sync();
  if (dashSheet.isNullObject) { logError("Tableau de bord introuvable"); return; }

  // Trouver la ligne TOTAL
  const totalRow = await findMarker(dashSheet, context, "TOTAL");
  if (!totalRow) { logError("Ligne TOTAL non trouvée dans le Tableau de bord"); return; }

  // Trouver le marqueur AVANCEMENT dans le projet pour les formules
  const projWs = context.workbook.worksheets.getItem(projet.code);
  const projStruct = await readProjectStructure(projWs, context);
  
  // Lignes de l'avancement dans le projet
  const safe = `'${projet.code}'`;
  
  // On insère avant TOTAL — on écrit sur la ligne totalRow (qui sera décalée)
  // Les formules pointent vers les marqueurs du projet
  const avancRow = projStruct.markers.AVANCEMENT;
  if (!avancRow) { logError("Marqueur AVANCEMENT non trouvé"); return; }

  // Budgété = avancRow+3, Consommé = avancRow+4, Restant = avancRow+5, % = avancRow+6
  const rBud = avancRow + 3;
  const rCons = avancRow + 4;
  const rRest = avancRow + 5;
  const rPct = avancRow + 6;

  const formulas = [
    [projet.code],
    [projet.nom],
    [`=${safe}!B${avancRow + 1 - avancRow + rBud - 2}`], // Catégorie
    [`=${safe}!B${rBud}`],
    [`=${safe}!B${rCons}`],
    [`=${safe}!B${rRest}`],
    [`=${safe}!B${rPct}`],
    [`=${safe}!C${rBud}`],
    [`=${safe}!C${rCons}`],
    [`=${safe}!C${rRest}`],
  ];

  for (let c = 0; c < formulas.length; c++) {
    dashSheet.getRangeByIndexes(totalRow - 1, c, 1, 1).formulas = [formulas[c]];
  }

  log(`  → Ajouté au Tableau de bord`);
}

// ============================================================
// MISE À JOUR LISTE RESSOURCES
// ============================================================

async function updateResourcesList(context, projet) {
  const resSheet = context.workbook.worksheets.getItemOrNullObject("Ressources");
  await context.sync();
  if (resSheet.isNullObject) return;

  const usedRange = resSheet.getUsedRange();
  usedRange.load("values");
  await context.sync();

  // Collecter les noms existants (colonne B, typiquement col index 1)
  const existingNames = new Set();
  for (let r = 0; r < usedRange.values.length; r++) {
    const name = String(usedRange.values[r][1] || "").trim();
    if (name && name !== "Nom") existingNames.add(name);
  }

  // Vérifier si des membres du projet sont nouveaux
  for (const membre of projet.membres) {
    if (!existingNames.has(membre.nom)) {
      log(`  → Nouvelle ressource détectée: ${membre.nom}`);
      // On ne l'ajoute pas automatiquement pour l'instant
      // TODO: insérer dans l'ordre correct (Louis-Michel, alpha, externes)
    }
  }
}

// ============================================================
// ORCHESTRATEUR
// ============================================================

async function processNutcacheImport(projets, reportPeriode) {
  log(`\n══════════════════════════════════════`);
  log(`Import Nutcache — ${projets.length} projet(s)`);
  log(`Période: ${reportPeriode.debut || "?"} → ${reportPeriode.fin || "?"}`);
  log(`══════════════════════════════════════`);

  try {
    await Excel.run(async (context) => {
      const existingCodes = await getExistingProjectCodes(context);
      const existingSheets = await getExistingSheetNames(context);

      log(`Projets existants: ${existingCodes.join(", ") || "(aucun)"}`);

      let nbUpdated = 0, nbCreated = 0;

      for (const projet of projets) {
        const codeExists = existingCodes.includes(projet.code);
        const sheetExists = existingSheets.includes(projet.code);

        if (codeExists && sheetExists) {
          await updateExistingProject(context, projet, reportPeriode);
          nbUpdated++;
        } else {
          await createNewProject(context, projet, reportPeriode);
          nbCreated++;
        }

        await updateResourcesList(context, projet);
      }

      await context.sync();

      log(`\n══════════════════════════════════════`);
      log(`TERMINÉ: ${nbUpdated} mis à jour, ${nbCreated} créé(s)`);
      log(`══════════════════════════════════════`);

      updateSummary(projets.length, nbUpdated, nbCreated);
    });
  } catch (error) {
    logError(`Erreur: ${error.message}`);
    throw error;
  }
}

function updateSummary(total, updated, created) {
  const el = document.getElementById("summary");
  if (el) {
    el.innerHTML = `
      <div class="summary-item"><span class="summary-num">${total}</span> projet(s) traité(s)</div>
      <div class="summary-item"><span class="summary-num summary-update">${updated}</span> mis à jour</div>
      <div class="summary-item"><span class="summary-num summary-create">${created}</span> créé(s)</div>
    `;
    el.style.display = "block";
  }
}
