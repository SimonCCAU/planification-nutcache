/**
 * updater.js — v14
 * 
 * Corrections clés :
 * - Insère des lignes quand les membres > lignes disponibles
 * - Nettoie les styles BOLD hérités du template
 * - Sauts de ligne entre sections
 * - Logo : supprime le dupliqué en R1
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
 * Retourne un tableau de strings (première valeur non-vide par ligne).
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
  for (let i = 0; i < colA.length; i++) if (colA[i] === text) return i + 1;
  return null;
}

function findPartial(colA, partial) {
  for (let i = 0; i < colA.length; i++) if (colA[i].includes(partial)) return i + 1;
  return null;
}

function findTotalAfter(colA, afterRowIdx) {
  for (let i = afterRowIdx; i < colA.length; i++) if (colA[i] === "TOTAL") return i + 1;
  return null;
}

async function readStructure(ws, context) {
  const colA = await scanRows(ws, context, 100);
  const s = {};
  s.colA = colA;
  s.titleRow = findPartial(colA, "CODE") || findPartial(colA, "{{") || 4;
  s.clientRow = findExact(colA, "Client");
  s.catRow = findPartial(colA, "Catégorie") || findPartial(colA, "Categorie");
  s.periodeRow = findPartial(colA, "riode") || findPartial(colA, "Periode");
  s.avancementRow = findExact(colA, "AVANCEMENT");
  s.allocRow = findExact(colA, "ALLOCATION DES RESSOURCES");
  s.allocHeaderRow = s.allocRow ? s.allocRow + 1 : null;
  s.allocDataStart = s.allocRow ? s.allocRow + 2 : null;
  s.allocTotalRow = s.allocRow ? findTotalAfter(colA, s.allocRow - 1) : null;
  s.allocSlots = (s.allocTotalRow && s.allocDataStart) ? s.allocTotalRow - s.allocDataStart : 0;
  s.planifRow = findPartial(colA, "PLANIFICATION CAPACITAIRE");
  s.planifHeaderRow = s.planifRow ? s.planifRow + 2 : null;
  s.planifDataStart = s.planifRow ? s.planifRow + 3 : null;
  s.planifTotalRow = s.planifRow ? findTotalAfter(colA, s.planifRow - 1) : null;
  s.planifSlots = (s.planifTotalRow && s.planifDataStart) ? s.planifTotalRow - s.planifDataStart : 0;
  s.phasesRow = findPartial(colA, "TAIL PAR PHASE") || findPartial(colA, "PHASE / SERVICE");
  s.phasesHeaderRow = s.phasesRow ? s.phasesRow + 1 : null;
  s.phasesDataStart = s.phasesRow ? s.phasesRow + 2 : null;
  s.notesRow = findPartial(colA, "NOTES DE SUIVI");
  
  log(`  titre=${s.titleRow} periode=${s.periodeRow} alloc=${s.allocRow}(slots=${s.allocSlots}) planif=${s.planifRow}(slots=${s.planifSlots}) phases=${s.phasesRow}`);
  return s;
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
// INSERTION DE LIGNES SI NÉCESSAIRE
// ============================================================

/**
 * Si on a besoin de N lignes mais seulement S slots disponibles,
 * insère (N - S) lignes à la position insertRow pour agrandir la section.
 */
async function ensureRows(ws, context, needed, slots, insertRowIdx) {
  const extra = needed - slots;
  if (extra <= 0) return 0;
  // Insérer des lignes avant le TOTAL
  const insertRange = ws.getRangeByIndexes(insertRowIdx, 0, extra, 14);
  insertRange.insert("Down");
  await context.sync();
  log(`  → Inséré ${extra} ligne(s) à R${insertRowIdx + 1}`);
  return extra;
}

// ============================================================
// MISE À JOUR PROJET EXISTANT
// ============================================================

async function updateExistingProject(context, projet, reportPeriode) {
  log(`↻ MAJ : ${projet.code}`);
  const ws = context.workbook.worksheets.getItem(projet.code);
  let s = await readStructure(ws, context);

  if (s.periodeRow && reportPeriode.debut) {
    ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin || ""]];
  }

  if (!s.allocDataStart || !s.allocTotalRow) {
    logError(`Struct alloc manquante pour ${projet.code}`);
    return;
  }

  // Insérer des lignes si nécessaire dans allocation
  const extraAlloc = await ensureRows(ws, context, projet.membres.length, s.allocSlots, s.allocTotalRow - 1);
  if (extraAlloc > 0) s = await readStructure(ws, context); // re-scanner

  // Effacer données existantes
  for (let i = 0; i < s.allocSlots; i++) {
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 5).values = [["", 0, 0, 0, 0]];
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 7).format.font.bold = false;
  }

  for (let i = 0; i < projet.membres.length; i++) {
    const m = projet.membres[i];
    const r = s.allocDataStart - 1 + i;
    ws.getRangeByIndexes(r, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
  }

  // Phases
  await updatePhases(ws, context, s, projet);
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

  // Supprimer les images dupliquées (logo) — effacer R1
  ws.getRangeByIndexes(0, 0, 3, 7).clear("Contents");

  let s = await readStructure(ws, context);
  const nMembres = projet.membres.length;

  // --- INSÉRER LIGNES POUR ALLOCATION ---
  const extraAlloc = await ensureRows(ws, context, nMembres, s.allocSlots, s.allocTotalRow - 1);
  if (extraAlloc > 0) s = await readStructure(ws, context);

  // --- INSÉRER LIGNES POUR PLANIF CAP ---
  const extraPlanif = await ensureRows(ws, context, nMembres, s.planifSlots, s.planifTotalRow - 1);
  if (extraPlanif > 0) s = await readStructure(ws, context);

  const accent = projet.tabColor || "#077C79";
  const fontOnAccent = (projet.categorie === "Consultant municipal") ? "#000000" : "#FFFFFF";

  // 1. TITRE
  log(`  Titre → R${s.titleRow}`);
  ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];
  const titleRange = ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 7);
  titleRange.format.fill.color = accent;
  titleRange.format.font.color = fontOnAccent;
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  // 2. EN-TÊTE
  if (s.clientRow) ws.getRangeByIndexes(s.clientRow - 1, 1, 1, 1).values = [[projet.client]];
  if (s.catRow) ws.getRangeByIndexes(s.catRow - 1, 1, 1, 1).values = [[projet.categorie || "Facturable"]];

  // 3. PÉRIODE
  if (s.periodeRow) {
    log(`  Période → R${s.periodeRow}`);
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin]];
  }

  // 4. ALLOCATION — écrire membres + nettoyer styles
  if (s.allocDataStart) {
    log(`  Alloc → R${s.allocDataStart} (${nMembres} membres, ${s.allocSlots} slots)`);
    for (let i = 0; i < nMembres; i++) {
      const m = projet.membres[i];
      const r = s.allocDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
      // Nettoyer le bold hérité du template TOTAL
      ws.getRangeByIndexes(r, 0, 1, 7).format.font.bold = false;
      ws.getRangeByIndexes(r, 0, 1, 7).format.fill.clear();
    }
  }

  // Mettre à jour les formules Consommé pour pointer vers le bon TOTAL
  if (s.avancementRow && s.allocTotalRow) {
    const rCons = s.avancementRow + 3; // Consommé = 3 lignes après AVANCEMENT
    ws.getRangeByIndexes(rCons - 1, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(rCons - 1, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  // 5. PLANIFICATION CAPACITAIRE — écrire noms + nettoyer
  if (s.planifDataStart) {
    log(`  Planif → R${s.planifDataStart}`);
    for (let i = 0; i < nMembres; i++) {
      const r = s.planifDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 1).values = [[projet.membres[i].nom]];
      // Nettoyer styles hérités
      ws.getRangeByIndexes(r, 0, 1, 14).format.font.bold = false;
      ws.getRangeByIndexes(r, 0, 1, 14).format.fill.clear();
    }
  }

  // 6. PHASES
  await updatePhases(ws, context, s, projet);

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

async function updatePhases(ws, context, s, projet) {
  if (!s.phasesDataStart) return;
  
  const phaseRows = [];
  for (const m of projet.membres)
    for (const svc of m.services)
      phaseRows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
  
  // Effacer 30 lignes et nettoyer styles
  const clearRange = ws.getRangeByIndexes(s.phasesDataStart - 1, 0, 30, 5);
  clearRange.clear("Contents");
  clearRange.format.font.bold = false;
  clearRange.format.fill.clear();
  
  if (phaseRows.length > 0)
    ws.getRangeByIndexes(s.phasesDataStart - 1, 0, phaseRows.length, 5).values = phaseRows;
}

// ============================================================
// COULEURS
// ============================================================

function applyColors(ws, s, accent, fontOnAccent) {
  function colorRow(row, cols) {
    if (!row) return;
    const r = ws.getRangeByIndexes(row - 1, 0, 1, cols);
    r.format.fill.color = accent;
    r.format.font.color = fontOnAccent;
    r.format.font.bold = true;
  }
  if (s.avancementRow) colorRow(s.avancementRow + 1, 3);
  if (s.allocHeaderRow) colorRow(s.allocHeaderRow, 7);
  if (s.planifHeaderRow) colorRow(s.planifHeaderRow, 14);
  if (s.phasesHeaderRow) colorRow(s.phasesHeaderRow, 5);
  if (s.notesRow) colorRow(s.notesRow + 1, 3);
}

// ============================================================
// INDEX PROJETS
// ============================================================

async function addToIndex(context, projet) {
  const sheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  const range = sheet.getRangeByIndexes(0, 0, 50, 1);
  range.load("values");
  await context.sync();

  let headerIdx = -1;
  for (let i = 0; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (v === "Code" || v.startsWith("Code")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) { logError("Headers Index non trouvés"); return; }

  // Trouver prochaine ligne vide
  let insertIdx = headerIdx + 1;
  for (let i = headerIdx + 1; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (!v) { insertIdx = i; break; }
    insertIdx = i + 1;
  }

  log(`  Index → R${insertIdx + 1}`);
  sheet.getRangeByIndexes(insertIdx, 0, 1, 11).values = [[
    projet.code, projet.nom, projet.client,
    projet.categorie || "Facturable", "En cours", "",
    "", "", 0, 0, "Import Nutcache"
  ]];
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

  let totalIdx = -1;
  for (let i = range.values.length - 1; i >= 0; i--) {
    if (String(range.values[i][0] || "").trim() === "TOTAL") { totalIdx = i; break; }
  }
  if (totalIdx === -1) { logError("TOTAL non trouvé"); return; }

  // Insérer ligne avant TOTAL
  sheet.getRangeByIndexes(totalIdx, 0, 1, 10).insert("Down");
  await context.sync();

  const safe = `'${projet.code}'`;
  const rBud = projStruct.avancementRow ? projStruct.avancementRow + 2 : 18;

  sheet.getRangeByIndexes(totalIdx, 0, 1, 3).values = [[
    projet.code, projet.nom, projet.categorie || "Facturable"
  ]];

  const formulas = [
    `=${safe}!B${rBud}`, `=${safe}!B${rBud+1}`, `=${safe}!B${rBud+2}`, `=${safe}!B${rBud+3}`,
    `=${safe}!C${rBud}`, `=${safe}!C${rBud+1}`, `=${safe}!C${rBud+2}`
  ];
  for (let c = 0; c < formulas.length; c++) {
    sheet.getRangeByIndexes(totalIdx, 3 + c, 1, 1).formulas = [[formulas[c]]];
  }
  log(`  Dashboard → R${totalIdx + 1}`);
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
