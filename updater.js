/**
 * updater.js — v13 — Correction complète
 * 
 * Changements clés vs v12 :
 * - scanColumnA lit 7 colonnes (A:G) pour gérer les cellules fusionnées
 * - Titre : écrit directement en A4 (position fixe du template) 
 * - Période : écrit directement en B14:C14 (position fixe)
 * - Index : scanne pour trouver prochaine ligne vide CORRECTEMENT
 * - Dashboard : insère ligne avant TOTAL avec formules correctes
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
 * Scanne les N premières lignes, colonnes A à G.
 * Retourne un tableau de strings = première valeur non-vide sur chaque ligne.
 * Gère les cellules fusionnées qui peuvent mettre la valeur en colonne B-G.
 */
async function scanRows(ws, context, numRows) {
  numRows = numRows || 100;
  const range = ws.getRangeByIndexes(0, 0, numRows, 7);
  range.load("values");
  await context.sync();
  
  const result = [];
  for (let r = 0; r < range.values.length; r++) {
    let val = "";
    for (let c = 0; c < 7; c++) {
      const v = String(range.values[r][c] || "").trim();
      if (v) { val = v; break; }
    }
    result.push(val);
  }
  return result;
}

function findExact(colA, text) {
  for (let i = 0; i < colA.length; i++) {
    if (colA[i] === text) return i + 1;
  }
  return null;
}

function findPartial(colA, partial) {
  for (let i = 0; i < colA.length; i++) {
    if (colA[i].includes(partial)) return i + 1;
  }
  return null;
}

function findTotalAfter(colA, afterRowIdx) {
  for (let i = afterRowIdx; i < colA.length; i++) {
    if (colA[i] === "TOTAL") return i + 1;
  }
  return null;
}

/**
 * Lit la structure du template copié.
 */
async function readStructure(ws, context) {
  const colA = await scanRows(ws, context, 100);
  
  // Debug: logger les marqueurs trouvés
  const titleRow = findPartial(colA, "CODE") || findPartial(colA, "{{") || 4;
  const clientRow = findExact(colA, "Client");
  const catRow = findPartial(colA, "Catégorie") || findPartial(colA, "Categorie");
  const periodeRow = findPartial(colA, "riode") || findPartial(colA, "Periode");
  const budgetHRow = findPartial(colA, "Budget heures");
  const budgetDRow = findPartial(colA, "Budget $");
  const avancementRow = findExact(colA, "AVANCEMENT");
  const allocRow = findExact(colA, "ALLOCATION DES RESSOURCES");
  const allocHeaderRow = allocRow ? allocRow + 1 : null;
  const allocDataStart = allocRow ? allocRow + 2 : null;
  const allocTotalRow = allocRow ? findTotalAfter(colA, allocRow - 1) : null;
  const planifRow = findPartial(colA, "PLANIFICATION CAPACITAIRE");
  const planifHeaderRow = planifRow ? planifRow + 2 : null;
  const planifDataStart = planifRow ? planifRow + 3 : null;
  const planifTotalRow = planifRow ? findTotalAfter(colA, planifRow - 1) : null;
  const phasesRow = findPartial(colA, "TAIL PAR PHASE") || findPartial(colA, "PHASE / SERVICE");
  const phasesHeaderRow = phasesRow ? phasesRow + 1 : null;
  const phasesDataStart = phasesRow ? phasesRow + 2 : null;
  const notesRow = findPartial(colA, "NOTES DE SUIVI");
  
  log(`  Struct: titre=${titleRow} client=${clientRow} periode=${periodeRow} avanc=${avancementRow}`);
  log(`  Struct: alloc=${allocRow}(hdr=${allocHeaderRow} data=${allocDataStart} tot=${allocTotalRow})`);
  log(`  Struct: planif=${planifRow} phases=${phasesRow} notes=${notesRow}`);
  
  return {
    colA, titleRow, clientRow, catRow, periodeRow, budgetHRow, budgetDRow,
    avancementRow, allocRow, allocHeaderRow, allocDataStart, allocTotalRow,
    allocMaxRows: (allocTotalRow && allocDataStart) ? allocTotalRow - allocDataStart : 1,
    planifRow, planifHeaderRow, planifDataStart, planifTotalRow,
    phasesRow, phasesHeaderRow, phasesDataStart, notesRow
  };
}

// ============================================================
// CODES EXISTANTS
// ============================================================

async function getExistingProjectCodes(context) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(INDEX_SHEET);
  await context.sync();
  if (sheet.isNullObject) return [];
  const colA = await scanRows(sheet, context, 50);
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
  log(`↻ MAJ : ${projet.code}`);
  const ws = context.workbook.worksheets.getItem(projet.code);
  const s = await readStructure(ws, context);

  if (s.periodeRow && reportPeriode.debut) {
    ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin || ""]];
  }

  if (!s.allocDataStart || !s.allocTotalRow) {
    logError(`Struct alloc manquante pour ${projet.code}`);
    return;
  }

  const maxR = s.allocMaxRows;
  const allocRange = ws.getRangeByIndexes(s.allocDataStart - 1, 0, maxR, 5);
  allocRange.load("values");
  await context.sync();

  // Effacer données existantes (écrasement)
  for (let i = 0; i < maxR; i++) {
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 2, 1, 3).values = [[0, 0, 0]];
  }

  const existing = {};
  let nextEmpty = s.allocDataStart;
  for (let i = 0; i < allocRange.values.length; i++) {
    const nm = String(allocRange.values[i][0] || "").trim();
    if (nm && nm !== "TOTAL") {
      existing[nm] = s.allocDataStart + i;
      nextEmpty = s.allocDataStart + i + 1;
    }
  }

  for (const m of projet.membres) {
    let row = existing[m.nom];
    if (!row) {
      if (nextEmpty >= s.allocTotalRow) { logError(`Plus de place: ${m.nom}`); continue; }
      row = nextEmpty;
      ws.getRangeByIndexes(row - 1, 0, 1, 1).values = [[m.nom]];
      nextEmpty++;
    }
    ws.getRangeByIndexes(row - 1, 2, 1, 3).values = [[m.heures, m.couts, m.tauxDeduit]];
  }

  // Phases
  await updatePhases(ws, s, projet);
  await context.sync();
  log(`✓ ${projet.code} MAJ (${projet.membres.length} membres)`);
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
  ws.showGridlines = false;

  const s = await readStructure(ws, context);

  // 1. TITRE — sur la ligne trouvée (fallback R4)
  log(`  Titre → R${s.titleRow}`);
  ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];
  
  // Couleur titre
  const accent = projet.tabColor || "#077C79";
  const fontOnAccent = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";
  const titleRange = ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 7);
  titleRange.format.fill.color = accent;
  titleRange.format.font.color = fontOnAccent;
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  // Effacer R1 si titre dupliqué (bug de templateSheet.copy)
  ws.getRangeByIndexes(0, 0, 1, 7).clear("Contents");

  // 2. EN-TÊTE — écrire via les positions trouvées
  if (s.clientRow) {
    log(`  Client → R${s.clientRow}`);
    ws.getRangeByIndexes(s.clientRow - 1, 1, 1, 1).values = [[projet.client]];
  }
  if (s.catRow) {
    ws.getRangeByIndexes(s.catRow - 1, 1, 1, 1).values = [[projet.categorie || "Facturable"]];
  }

  // 3. PÉRIODE — sur la bonne ligne
  if (s.periodeRow) {
    log(`  Période → R${s.periodeRow}`);
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin]];
  } else {
    logError("  Marqueur Période non trouvé!");
  }

  // 4. MEMBRES ALLOCATION
  if (s.allocDataStart) {
    log(`  Alloc data → R${s.allocDataStart} (${projet.membres.length} membres)`);
    for (let i = 0; i < projet.membres.length; i++) {
      const m = projet.membres[i];
      const r = s.allocDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
    }
  }

  // 5. MEMBRES PLANIF CAPACITAIRE
  if (s.planifDataStart) {
    log(`  Planif data → R${s.planifDataStart}`);
    for (let i = 0; i < projet.membres.length; i++) {
      ws.getRangeByIndexes(s.planifDataStart - 1 + i, 0, 1, 1).values = [[projet.membres[i].nom]];
    }
  }

  // 6. PHASES
  await updatePhases(ws, s, projet);

  // 7. COULEURS EN-TÊTES
  applyColors(ws, s, accent, fontOnAccent);

  // 8. INDEX + DASHBOARD
  await addToIndex(context, projet);
  await addToDashboard(context, projet, s);

  await context.sync();
  log(`✓ ${projet.code} créé`);
}

// ============================================================
// PHASES
// ============================================================

async function updatePhases(ws, s, projet) {
  if (!s.phasesDataStart) return;
  
  const phaseRows = [];
  for (const m of projet.membres)
    for (const svc of m.services)
      phaseRows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
  
  // Effacer 30 lignes
  ws.getRangeByIndexes(s.phasesDataStart - 1, 0, 30, 5).clear("Contents");
  
  if (phaseRows.length > 0)
    ws.getRangeByIndexes(s.phasesDataStart - 1, 0, phaseRows.length, 5).values = phaseRows;
}

// ============================================================
// COULEURS EN-TÊTES PAR TYPE
// ============================================================

function applyColors(ws, s, accent, fontOnAccent) {
  function colorRow(row, cols) {
    if (!row) return;
    const r = ws.getRangeByIndexes(row - 1, 0, 1, cols);
    r.format.fill.color = accent;
    r.format.font.color = fontOnAccent;
  }
  
  // Avancement header (ligne après AVANCEMENT = Heures/$)
  if (s.avancementRow) colorRow(s.avancementRow + 1, 3);
  // Allocation header
  if (s.allocHeaderRow) colorRow(s.allocHeaderRow, 7);
  // Planif header  
  if (s.planifHeaderRow) colorRow(s.planifHeaderRow, 14);
  // Phases header
  if (s.phasesHeaderRow) colorRow(s.phasesHeaderRow, 5);
  // Notes header
  if (s.notesRow) colorRow(s.notesRow + 1, 3);
}

// ============================================================
// INDEX PROJETS
// ============================================================

async function addToIndex(context, projet) {
  const sheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  
  // Lire colonne A pour trouver la prochaine ligne vide APRÈS les headers
  const range = sheet.getRangeByIndexes(0, 0, 50, 1);
  range.load("values");
  await context.sync();
  
  // Trouver la ligne des headers (contient "Code")
  let headerRow = -1;
  for (let i = 0; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (v === "Code" || v === "Code projet") { headerRow = i; break; }
  }
  if (headerRow === -1) { logError("Headers Index non trouvés"); return; }
  
  // Trouver la prochaine ligne vide après les headers
  let insertIdx = headerRow + 1;
  for (let i = headerRow + 1; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (!v) { insertIdx = i; break; }
    insertIdx = i + 1; // après la dernière ligne non-vide
  }

  log(`  Index → R${insertIdx + 1} (après header R${headerRow + 1})`);
  
  const data = [
    projet.code, projet.nom, projet.client,
    projet.categorie || "Facturable", "En cours", "",
    "", "", 0, 0,
    "Importé depuis Nutcache"
  ];
  sheet.getRangeByIndexes(insertIdx, 0, 1, data.length).values = [data];
}

// ============================================================
// TABLEAU DE BORD
// ============================================================

async function addToDashboard(context, projet, projStruct) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await context.sync();
  if (sheet.isNullObject) { logError("Dashboard introuvable"); return; }

  const range = sheet.getRangeByIndexes(0, 0, 50, 1);
  range.load("values");
  await context.sync();

  // Trouver TOTAL (dernière occurrence)
  let totalIdx = -1;
  for (let i = range.values.length - 1; i >= 0; i--) {
    if (String(range.values[i][0] || "").trim() === "TOTAL") { totalIdx = i; break; }
  }
  if (totalIdx === -1) { logError("TOTAL non trouvé dans Dashboard"); return; }

  // Insérer une ligne AVANT le TOTAL
  const insertRange = sheet.getRangeByIndexes(totalIdx, 0, 1, 10);
  insertRange.insert("Down");
  await context.sync();

  log(`  Dashboard → insertion avant TOTAL (R${totalIdx + 1})`);

  // La ligne insérée est à totalIdx, le TOTAL est maintenant à totalIdx+1
  const safe = `'${projet.code}'`;
  
  // Positions dans le projet créé
  const rBud = projStruct.avancementRow ? projStruct.avancementRow + 2 : 18;
  const rCons = rBud + 1;
  const rRest = rBud + 2;
  const rPct = rBud + 3;

  // Écrire code + nom + catégorie en valeurs
  sheet.getRangeByIndexes(totalIdx, 0, 1, 3).values = [[
    projet.code,
    projet.nom,
    projet.categorie || "Facturable"
  ]];

  // Écrire formules colonnes D-J (index 3-9)
  const formulas = [
    `=${safe}!B${rBud}`,    // D: Budget h
    `=${safe}!B${rCons}`,   // E: Consommé h
    `=${safe}!B${rRest}`,   // F: Restant h
    `=${safe}!B${rPct}`,    // G: % Avanc
    `=${safe}!C${rBud}`,    // H: Budget $
    `=${safe}!C${rCons}`,   // I: Consommé $
    `=${safe}!C${rRest}`,   // J: Écart $
  ];
  for (let c = 0; c < formulas.length; c++) {
    sheet.getRangeByIndexes(totalIdx, 3 + c, 1, 1).formulas = [[formulas[c]]];
  }
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
    console.error(error);
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
