/**
 * updater.js — v12 — Réécriture complète
 * Utilise scanRange(100 lignes) pour TOUS les lookups.
 * Pas de getUsedRange() qui est peu fiable sur templates copiés.
 */

const TEMPLATE_SHEET = "_TEMPLATE";
const INDEX_SHEET = "Index Projets";
const DASHBOARD_SHEET = "Tableau de bord";

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
 * Scanne les N premières lignes de la colonne A et retourne un tableau de valeurs.
 * Toujours fiable, contrairement à getUsedRange.
 */
async function scanColumnA(ws, context, numRows) {
  numRows = numRows || 100;
  const range = ws.getRangeByIndexes(0, 0, numRows, 1);
  range.load("values");
  await context.sync();
  return range.values.map(r => String(r[0] || "").trim());
}

/**
 * Trouve la première ligne contenant exactement 'text' ou commençant par 'text' dans colA.
 */
function findInScan(colA, text) {
  for (let i = 0; i < colA.length; i++) {
    if (colA[i] === text || colA[i].startsWith(text)) return i + 1; // 1-based
  }
  return null;
}

/**
 * Trouve la première ligne contenant 'partial' quelque part dans colA.
 */
function findPartialInScan(colA, partial) {
  for (let i = 0; i < colA.length; i++) {
    if (colA[i].includes(partial)) return i + 1;
  }
  return null;
}

/**
 * Trouve la première ligne "TOTAL" après une position donnée.
 */
function findTotalAfterInScan(colA, afterRow) {
  for (let i = afterRow; i < colA.length; i++) {
    if (colA[i] === "TOTAL") return i + 1;
  }
  return null;
}

/**
 * Lit la structure complète d'un onglet projet via scan colA.
 */
async function readStructure(ws, context) {
  const colA = await scanColumnA(ws, context, 100);
  
  const titleRow = findPartialInScan(colA, "CODE") || findPartialInScan(colA, "{{") || 4;
  const periodeRow = findPartialInScan(colA, "riode") || findPartialInScan(colA, "Periode");
  const avancementRow = findInScan(colA, "AVANCEMENT");
  const allocRow = findInScan(colA, "ALLOCATION DES RESSOURCES");
  const allocHeaderRow = allocRow ? allocRow + 1 : null;
  const allocDataStart = allocRow ? allocRow + 2 : null;
  const allocTotalRow = allocRow ? findTotalAfterInScan(colA, allocRow) : null;
  const planifRow = findPartialInScan(colA, "PLANIFICATION CAPACITAIRE");
  const planifHeaderRow = planifRow ? planifRow + 2 : null;
  const planifDataStart = planifRow ? planifRow + 3 : null;
  const planifTotalRow = planifRow ? findTotalAfterInScan(colA, planifRow) : null;
  const phasesRow = findPartialInScan(colA, "TAIL PAR PHASE");
  const phasesHeaderRow = phasesRow ? phasesRow + 1 : null;
  const phasesDataStart = phasesRow ? phasesRow + 2 : null;
  
  log(`  Structure: titre=${titleRow} periode=${periodeRow} alloc=${allocRow}(data=${allocDataStart},total=${allocTotalRow}) planif=${planifRow} phases=${phasesRow}`);
  
  return {
    colA, titleRow, periodeRow, avancementRow,
    allocRow, allocHeaderRow, allocDataStart, allocTotalRow,
    allocMaxRows: (allocTotalRow && allocDataStart) ? allocTotalRow - allocDataStart : 1,
    planifRow, planifHeaderRow, planifDataStart, planifTotalRow,
    phasesRow, phasesHeaderRow, phasesDataStart
  };
}

// ============================================================
// CODES EXISTANTS
// ============================================================

async function getExistingProjectCodes(context) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(INDEX_SHEET);
  await context.sync();
  if (sheet.isNullObject) return [];
  
  const colA = await scanColumnA(sheet, context, 50);
  return colA.filter(v => /^\d{5}$|^[A-Z]{2,4}-\d{3}$/.test(v));
}

async function getExistingSheetNames(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  return sheets.items.map(s => s.name);
}

// ============================================================
// MISE À JOUR PROJET EXISTANT
// ============================================================

async function updateExistingProject(context, projet, reportPeriode) {
  log(`↻ Mise à jour : ${projet.code}`);
  const ws = context.workbook.worksheets.getItem(projet.code);
  const s = await readStructure(ws, context);

  // Période
  if (s.periodeRow && reportPeriode.debut) {
    ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin || ""]];
  }

  // Lire les ressources existantes
  if (!s.allocDataStart || !s.allocTotalRow) {
    logError(`Structure allocation non trouvée pour ${projet.code}`);
    return;
  }
  
  const maxRows = s.allocMaxRows;
  const allocRange = ws.getRangeByIndexes(s.allocDataStart - 1, 0, maxRows, 5);
  allocRange.load("values");
  await context.sync();

  const existingNames = {};
  let nextEmpty = s.allocDataStart;
  for (let i = 0; i < allocRange.values.length; i++) {
    const name = String(allocRange.values[i][0] || "").trim();
    if (name && name !== "TOTAL") {
      existingNames[name] = {
        row: s.allocDataStart + i,
        heures: parseFloat(allocRange.values[i][2]) || 0,
        couts: parseFloat(allocRange.values[i][3]) || 0
      };
      nextEmpty = s.allocDataStart + i + 1;
    }
  }

  // Effacer données (écrasement)
  for (let i = 0; i < maxRows; i++) {
    const r = s.allocDataStart - 1 + i;
    ws.getRangeByIndexes(r, 2, 1, 3).values = [[0, 0, 0]];
  }

  // Injecter membres
  for (const m of projet.membres) {
    let row;
    if (existingNames[m.nom]) {
      row = existingNames[m.nom].row;
    } else {
      if (nextEmpty >= s.allocTotalRow) { logError(`Plus de place pour ${m.nom}`); continue; }
      row = nextEmpty;
      ws.getRangeByIndexes(row - 1, 0, 1, 1).values = [[m.nom]];
      nextEmpty++;
    }
    ws.getRangeByIndexes(row - 1, 2, 1, 3).values = [[m.heures, m.couts, m.tauxDeduit]];
  }

  // Phases
  if (s.phasesDataStart) {
    const phaseRows = [];
    for (const m of projet.membres)
      for (const svc of m.services)
        phaseRows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
    
    ws.getRangeByIndexes(s.phasesDataStart - 1, 0, 30, 5).clear("Contents");
    if (phaseRows.length > 0)
      ws.getRangeByIndexes(s.phasesDataStart - 1, 0, phaseRows.length, 5).values = phaseRows;
  }

  await context.sync();
  log(`✓ ${projet.code} mis à jour (${projet.membres.length} membres)`);
}

// ============================================================
// CRÉATION NOUVEAU PROJET
// ============================================================

async function createNewProject(context, projet, reportPeriode) {
  log(`✚ Nouveau : ${projet.code} — ${projet.nom}`);

  const tpl = context.workbook.worksheets.getItem(TEMPLATE_SHEET);
  const ws = tpl.copy("End");
  await context.sync();

  ws.name = projet.code;
  ws.visibility = "Visible";
  ws.tabColor = projet.tabColor || "#077C79";

  // Scanner la structure du template copié
  const s = await readStructure(ws, context);

  // 1. TITRE — écrire en R4 (position du {{CODE}})
  const titleText = `${projet.code} — ${projet.nom}`;
  ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 1).values = [[titleText]];
  
  // Appliquer couleur titre
  const titleRange = ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 7);
  titleRange.format.fill.color = projet.tabColor || "#077C79";
  const fontColor = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";
  titleRange.format.font.color = fontColor;
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  // Effacer R1 si titre dupliqué
  ws.getRangeByIndexes(0, 0, 1, 7).values = [["","","","","","",""]];

  // 2. EN-TÊTE — scanner et remplir par label
  const colA = s.colA;
  for (let i = 0; i < colA.length; i++) {
    const label = colA[i];
    if (label === "Client") ws.getRangeByIndexes(i, 1, 1, 1).values = [[projet.client]];
    if (label === "Catégorie" || label === "Categorie") ws.getRangeByIndexes(i, 1, 1, 1).values = [[projet.categorie || "Facturable"]];
  }

  // 3. PÉRIODE — écrire sur la bonne ligne (pas R11!)
  if (s.periodeRow) {
    log(`  Période en ligne ${s.periodeRow}`);
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin]];
  }

  // 4. MEMBRES ALLOCATION
  if (s.allocDataStart) {
    for (let i = 0; i < projet.membres.length; i++) {
      const m = projet.membres[i];
      const r = s.allocDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
    }
  }

  // 5. MEMBRES PLANIF CAPACITAIRE
  if (s.planifDataStart) {
    for (let i = 0; i < projet.membres.length; i++) {
      ws.getRangeByIndexes(s.planifDataStart - 1 + i, 0, 1, 1).values = [[projet.membres[i].nom]];
    }
  }

  // 6. PHASES
  if (s.phasesDataStart) {
    const phaseRows = [];
    for (const m of projet.membres)
      for (const svc of m.services)
        phaseRows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
    if (phaseRows.length > 0)
      ws.getRangeByIndexes(s.phasesDataStart - 1, 0, phaseRows.length, 5).values = phaseRows;
  }

  // 7. COULEURS DES EN-TÊTES DE TABLEAUX
  await applyProjectColors(ws, context, s, projet);

  // 8. INDEX + DASHBOARD
  await addToIndex(context, projet);
  await addToDashboard(context, projet, s);

  await context.sync();
  log(`✓ ${projet.code} créé`);
}

// ============================================================
// COULEURS PAR TYPE
// ============================================================

async function applyProjectColors(ws, context, s, projet) {
  const accent = projet.tabColor || "#077C79";
  const fc = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";

  function colorRange(rowIdx, cols) {
    if (!rowIdx) return;
    const range = ws.getRangeByIndexes(rowIdx - 1, 0, 1, cols);
    range.format.fill.color = accent;
    range.format.font.color = fc;
  }

  // Avancement header (Heures / $)
  if (s.avancementRow) colorRange(s.avancementRow + 1, 3);
  // Allocation header
  if (s.allocHeaderRow) colorRange(s.allocHeaderRow, 7);
  // Planif capacitaire header
  if (s.planifHeaderRow) colorRange(s.planifHeaderRow, 14);
  // Phases header
  if (s.phasesHeaderRow) colorRange(s.phasesHeaderRow, 5);
  // Notes header
  const notesRow = findPartialInScan(s.colA, "NOTES DE SUIVI");
  if (notesRow) colorRange(notesRow + 1, 3);

  await context.sync();
}

// ============================================================
// INDEX PROJETS — ajouter sur la prochaine ligne vide
// ============================================================

async function addToIndex(context, projet) {
  const sheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  
  // Scanner pour trouver la prochaine ligne vide après les headers
  const colA = await scanColumnA(sheet, context, 50);
  
  // Trouver la ligne des headers (contient "Code")
  let headerRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i] === "Code" || colA[i].startsWith("Code")) { headerRow = i + 1; break; }
  }
  if (!headerRow) { logError("Headers Index non trouvés"); return; }

  // Trouver la prochaine ligne vide après les headers
  let insertRow = headerRow + 1;
  for (let i = headerRow; i < colA.length; i++) {
    if (colA[i] && colA[i] !== "") {
      insertRow = i + 2; // après cette ligne
    }
  }

  const data = [
    projet.code, projet.nom, projet.client,
    projet.categorie || "Facturable", "En cours", "",
    "", "", 0, 0,
    "Importé depuis Nutcache"
  ];
  sheet.getRangeByIndexes(insertRow - 1, 0, 1, data.length).values = [data];
  log(`  → Index ligne ${insertRow}`);
}

// ============================================================
// TABLEAU DE BORD — ajouter avant TOTAL
// ============================================================

async function addToDashboard(context, projet, projStruct) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await context.sync();
  if (sheet.isNullObject) { logError("Tableau de bord introuvable"); return; }

  const colA = await scanColumnA(sheet, context, 50);
  
  // Trouver la ligne TOTAL
  let totalRow = null;
  for (let i = colA.length - 1; i >= 0; i--) {
    if (colA[i] === "TOTAL") { totalRow = i + 1; break; }
  }
  if (!totalRow) { logError("TOTAL non trouvé dans Dashboard"); return; }

  // Insérer une ligne avant TOTAL (décaler TOTAL vers le bas)
  sheet.getRangeByIndexes(totalRow - 1, 0, 1, 1).insert("Down");
  await context.sync();

  // La ligne insérée est maintenant à totalRow, TOTAL est à totalRow+1
  const safe = `'${projet.code}'`;
  
  // Utiliser les positions de la structure du projet
  const rBud = projStruct.avancementRow ? projStruct.avancementRow + 2 : 18;
  const rCons = rBud + 1;
  const rRest = rBud + 2;
  const rPct = rBud + 3;

  const data = [[
    projet.code,
    projet.nom,
    projet.categorie || "Facturable"
  ]];
  sheet.getRangeByIndexes(totalRow - 1, 0, 1, 3).values = data;

  // Formules
  const formulas = [
    [`=${safe}!B${rBud}`],   // Budget h
    [`=${safe}!B${rCons}`],  // Consommé h
    [`=${safe}!B${rRest}`],  // Restant h
    [`=${safe}!B${rPct}`],   // % Avanc
    [`=${safe}!C${rBud}`],   // Budget $
    [`=${safe}!C${rCons}`],  // Consommé $
    [`=${safe}!C${rRest}`],  // Écart $
  ];
  for (let c = 0; c < formulas.length; c++) {
    sheet.getRangeByIndexes(totalRow - 1, 3 + c, 1, 1).formulas = [formulas[c]];
  }

  log(`  → Dashboard ligne ${totalRow}`);
}

// ============================================================
// ORCHESTRATEUR
// ============================================================

async function processNutcacheImport(projets, reportPeriode) {
  log(`\n══════════════════════════════════════`);
  log(`Import — ${projets.length} projet(s)`);
  if (reportPeriode.debut) log(`Période: ${reportPeriode.debut} → ${reportPeriode.fin}`);
  log(`══════════════════════════════════════`);

  try {
    await Excel.run(async (context) => {
      const existingCodes = await getExistingProjectCodes(context);
      const existingSheets = await getExistingSheetNames(context);
      log(`Existants: ${existingCodes.join(", ") || "(aucun)"}`);

      let up = 0, cr = 0;
      for (const p of projets) {
        if (existingCodes.includes(p.code) && existingSheets.includes(p.code)) {
          await updateExistingProject(context, p, reportPeriode);
          up++;
        } else {
          await createNewProject(context, p, reportPeriode);
          cr++;
        }
      }
      await context.sync();
      log(`\nTERMINÉ: ${up} mis à jour, ${cr} créé(s)`);
      updateSummary(projets.length, up, cr);
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
      <div class="summary-item"><span class="summary-num">${total}</span> projet(s)</div>
      <div class="summary-item"><span class="summary-num summary-update">${updated}</span> mis à jour</div>
      <div class="summary-item"><span class="summary-num summary-create">${created}</span> créé(s)</div>
    `;
    el.style.display = "block";
  }
}
