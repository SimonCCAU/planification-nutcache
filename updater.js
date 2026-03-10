/**
 * updater.js — v15
 * 
 * 3 templates par type de projet (styles pré-intégrés)
 * Dashboard : pas de doublon, mise en forme conditionnelle colonne A
 * Insertion dynamique de lignes
 * Nettoyage des styles hérités
 */

const TPL_MAP = {
  "Facturable": "_TPL_FACTURABLE",
  "Non-facturable": "_TPL_INTERNE",
  "Consultant municipal": "_TPL_CONSULTANT"
};
const INDEX_SHEET = "Index Projets";
const DASHBOARD_SHEET = "Tableau de bord";

const COLOR_MAP = {
  "Facturable": { bg: "#077C79", font: "#FFFFFF" },
  "Non-facturable": { bg: "#EB5E63", font: "#FFFFFF" },
  "Consultant municipal": { bg: "#EEDB4C", font: "#000000" }
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

async function scanRows(ws, context, numRows) {
  numRows = numRows || 100;
  const range = ws.getRangeByIndexes(0, 0, numRows, 7);
  range.load("values");
  await context.sync();
  return range.values.map(row => {
    for (let c = 0; c < 7; c++) {
      const v = String(row[c] || "").trim();
      if (v) return v;
    }
    return "";
  });
}

function findExact(colA, text) {
  for (let i = 0; i < colA.length; i++) if (colA[i] === text) return i + 1;
  return null;
}

function findPartial(colA, partial) {
  for (let i = 0; i < colA.length; i++) if (colA[i].includes(partial)) return i + 1;
  return null;
}

function findTotalAfter(colA, afterRow) {
  for (let i = afterRow; i < colA.length; i++) if (colA[i] === "TOTAL") return i + 1;
  return null;
}

async function readStructure(ws, context) {
  const colA = await scanRows(ws, context, 100);
  const s = { colA };
  s.titleRow = findPartial(colA, "CODE") || findPartial(colA, "{{") || 4;
  s.clientRow = findExact(colA, "Client");
  s.catRow = findPartial(colA, "gorie") || findPartial(colA, "Categorie");
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
  s.phasesRow = findPartial(colA, "TAIL PAR PHASE") || findPartial(colA, "PHASE");
  s.phasesHeaderRow = s.phasesRow ? s.phasesRow + 1 : null;
  s.phasesDataStart = s.phasesRow ? s.phasesRow + 2 : null;
  s.notesRow = findPartial(colA, "NOTES DE SUIVI");
  log(`  alloc=${s.allocRow}(${s.allocSlots}sl) planif=${s.planifRow}(${s.planifSlots}sl) phases=${s.phasesRow}`);
  return s;
}

async function ensureRows(ws, context, needed, slots, totalRowIdx) {
  const extra = needed - slots;
  if (extra <= 0) return 0;
  ws.getRangeByIndexes(totalRowIdx - 1, 0, extra, 14).insert("Down");
  await context.sync();
  log(`  +${extra} lignes avant R${totalRowIdx}`);
  return extra;
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
// MISE À JOUR EXISTANT
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
    logError(`Alloc manquante: ${projet.code}`); return;
  }

  const extra = await ensureRows(ws, context, projet.membres.length, s.allocSlots, s.allocTotalRow);
  if (extra > 0) s = await readStructure(ws, context);

  // Effacer + écrire
  for (let i = 0; i < s.allocSlots; i++) {
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 7).clear("Contents");
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 7).format.font.bold = false;
  }
  for (let i = 0; i < projet.membres.length; i++) {
    const m = projet.membres[i];
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
  }

  // Mettre à jour formule Consommé
  if (s.avancementRow && s.allocTotalRow) {
    const rCons = s.avancementRow + 3;
    ws.getRangeByIndexes(rCons - 1, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(rCons - 1, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  await updatePhases(ws, context, s, projet);
  await context.sync();
  log(`✓ ${projet.code} MAJ`);
}

// ============================================================
// NOUVEAU PROJET
// ============================================================

async function createNewProject(context, projet, reportPeriode) {
  log(`✚ ${projet.code} — ${projet.nom}`);

  // Choisir le bon template selon le type
  const tplName = TPL_MAP[projet.categorie] || "_TPL_FACTURABLE";
  let tpl = context.workbook.worksheets.getItemOrNullObject(tplName);
  await context.sync();
  if (tpl.isNullObject) {
    tpl = context.workbook.worksheets.getItem("_TEMPLATE");
    log(`  Template ${tplName} absent, fallback _TEMPLATE`);
  }
  
  const ws = tpl.copy("End");
  await context.sync();
  ws.name = projet.code;
  ws.visibility = "Visible";
  ws.tabColor = projet.tabColor || "#077C79";

  // Nettoyer R1 (logo dupliqué)
  ws.getRangeByIndexes(0, 0, 3, 7).clear("Contents");

  let s = await readStructure(ws, context);
  const n = projet.membres.length;

  // Insérer lignes allocation
  const ea = await ensureRows(ws, context, n, s.allocSlots, s.allocTotalRow);
  if (ea > 0) s = await readStructure(ws, context);

  // Insérer lignes planif cap
  const ep = await ensureRows(ws, context, n, s.planifSlots, s.planifTotalRow);
  if (ep > 0) s = await readStructure(ws, context);

  // TITRE
  ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];

  // EN-TÊTE
  if (s.clientRow) ws.getRangeByIndexes(s.clientRow - 1, 1, 1, 1).values = [[projet.client]];
  if (s.catRow) ws.getRangeByIndexes(s.catRow - 1, 1, 1, 1).values = [[projet.categorie || "Facturable"]];

  // PÉRIODE
  if (s.periodeRow) {
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin]];
  }

  // ALLOCATION
  if (s.allocDataStart) {
    for (let i = 0; i < n; i++) {
      const m = projet.membres[i]; const r = s.allocDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
      ws.getRangeByIndexes(r, 0, 1, 7).format.font.bold = false;
    }
  }

  // Mettre à jour formule Consommé
  if (s.avancementRow && s.allocTotalRow) {
    const rCons = s.avancementRow + 3;
    ws.getRangeByIndexes(rCons - 1, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(rCons - 1, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  // PLANIF CAP
  if (s.planifDataStart) {
    for (let i = 0; i < n; i++) {
      const r = s.planifDataStart - 1 + i;
      ws.getRangeByIndexes(r, 0, 1, 1).values = [[projet.membres[i].nom]];
      ws.getRangeByIndexes(r, 0, 1, 14).format.font.bold = false;
    }
  }

  // PHASES
  await updatePhases(ws, context, s, projet);

  // INDEX + DASHBOARD
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
  const rows = [];
  for (const m of projet.membres)
    for (const svc of m.services)
      rows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
  
  // Nettoyer 40 lignes
  const cr = ws.getRangeByIndexes(s.phasesDataStart - 1, 0, 40, 5);
  cr.clear("Contents");
  cr.format.font.bold = false;
  cr.format.fill.clear();
  
  if (rows.length > 0)
    ws.getRangeByIndexes(s.phasesDataStart - 1, 0, rows.length, 5).values = rows;
}

// ============================================================
// INDEX — pas de doublon
// ============================================================

async function addToIndex(context, projet) {
  const sheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  const range = sheet.getRangeByIndexes(0, 0, 50, 1);
  range.load("values");
  await context.sync();

  // Vérifier si déjà présent
  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim() === projet.code) {
      log(`  Index: ${projet.code} déjà présent`);
      return;
    }
  }

  let headerIdx = -1;
  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim().startsWith("Code")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) { logError("Headers Index non trouvés"); return; }

  let insertIdx = headerIdx + 1;
  for (let i = headerIdx + 1; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (!v) { insertIdx = i; break; }
    insertIdx = i + 1;
  }

  sheet.getRangeByIndexes(insertIdx, 0, 1, 11).values = [[
    projet.code, projet.nom, projet.client,
    projet.categorie || "Facturable", "En cours", "", "", "", 0, 0, "Import Nutcache"
  ]];
  log(`  Index → R${insertIdx + 1}`);
}

// ============================================================
// DASHBOARD — pas de doublon + couleur colonne A
// ============================================================

async function addToDashboard(context, projet, projStruct) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await context.sync();
  if (sheet.isNullObject) { logError("Dashboard introuvable"); return; }

  const range = sheet.getRangeByIndexes(0, 0, 50, 1);
  range.load("values");
  await context.sync();

  // Vérifier si déjà présent
  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim() === projet.code) {
      log(`  Dashboard: ${projet.code} déjà présent`);
      return;
    }
  }

  let totalIdx = -1;
  for (let i = range.values.length - 1; i >= 0; i--) {
    if (String(range.values[i][0] || "").trim() === "TOTAL") { totalIdx = i; break; }
  }
  if (totalIdx === -1) { logError("TOTAL non trouvé"); return; }

  sheet.getRangeByIndexes(totalIdx, 0, 1, 10).insert("Down");
  await context.sync();

  const safe = `'${projet.code}'`;
  const rBud = projStruct.avancementRow ? projStruct.avancementRow + 2 : 18;

  sheet.getRangeByIndexes(totalIdx, 0, 1, 3).values = [[
    projet.code, projet.nom, projet.categorie || "Facturable"
  ]];

  // Formules
  const fmls = [
    `=${safe}!B${rBud}`, `=${safe}!B${rBud+1}`, `=${safe}!B${rBud+2}`, `=${safe}!B${rBud+3}`,
    `=${safe}!C${rBud}`, `=${safe}!C${rBud+1}`, `=${safe}!C${rBud+2}`
  ];
  for (let c = 0; c < fmls.length; c++) {
    sheet.getRangeByIndexes(totalIdx, 3 + c, 1, 1).formulas = [[fmls[c]]];
  }

  // Mise en forme conditionnelle — couleur de la cellule Code (colonne A)
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];
  const codeCell = sheet.getRangeByIndexes(totalIdx, 0, 1, 1);
  codeCell.format.fill.color = colors.bg;
  codeCell.format.font.color = colors.font;
  codeCell.format.font.bold = true;

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
      log(`\nTERMINÉ: ${up} MAJ, ${cr} créé(s)`);
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
      <div class="summary-item"><span class="summary-num summary-update">${updated}</span> MAJ</div>
      <div class="summary-item"><span class="summary-num summary-create">${created}</span> créé(s)</div>
    `;
    el.style.display = "block";
  }
}
