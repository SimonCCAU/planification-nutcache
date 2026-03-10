/**
 * updater.js — v19
 * 
 * Corrections :
 * - Ressources peuplées dynamiquement (pas de consultant externe en dur)
 * - Dashboard : alternance gris + couleur code A par type
 * - Index : alternance gris
 * - Planif cap : ref heures/sem depuis Ressources
 * - Tri onglets corrigé
 * - Bordures gris pâle sur toutes les lignes
 */

const TPL_MAP = {
  "Facturable": "_TPL_FACTURABLE",
  "Non-facturable": "_TPL_INTERNE",
  "Consultant municipal": "_TPL_CONSULTANT"
};
const INDEX_SHEET = "Index Projets";
const DASHBOARD_SHEET = "Tableau de bord";
const RESSOURCES_SHEET = "Ressources";
const PLANIF_CAP_SHEET = "Planification capacitaire";

const COLOR_MAP = {
  "Facturable": { bg: "#077C79", font: "#FFFFFF", pale: "#E6F2F1" },
  "Non-facturable": { bg: "#EB5E63", font: "#FFFFFF", pale: "#FDEEEF" },
  "Consultant municipal": { bg: "#EEDB4C", font: "#000000", pale: "#FDF9E5" }
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

function findExact(a, t) { for (let i=0;i<a.length;i++) if(a[i]===t) return i+1; return null; }
function findPartial(a, t) { for (let i=0;i<a.length;i++) if(a[i].includes(t)) return i+1; return null; }
function findTotalAfter(a, r) { for (let i=r;i<a.length;i++) if(a[i]==="TOTAL") return i+1; return null; }

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
  return s;
}

async function ensureRows(ws, context, needed, slots, totalRowIdx) {
  const extra = needed - slots;
  if (extra <= 0) return 0;
  ws.getRangeByIndexes(totalRowIdx - 1, 0, extra, 14).insert("Down");
  await context.sync();
  return extra;
}

function styleDataRows(ws, startRow, count, cols, paleColor) {
  for (let i = 0; i < count; i++) {
    const r = startRow - 1 + i;
    const bg = (i % 2 === 0) ? paleColor : "#FFFFFF";
    const range = ws.getRangeByIndexes(r, 0, 1, cols);
    range.format.fill.color = bg;
    range.format.font.bold = false;
    range.format.font.color = "#4A4A4A";
    range.format.font.size = 10;
    try {
      const border = range.format.borders.getItem(Excel.BorderIndex.edgeBottom);
      border.style = Excel.BorderLineStyle.thin;
      border.color = "#E0E0E0";
    } catch(e) {}
  }
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
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];

  if (s.periodeRow && reportPeriode.debut) {
    ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin || ""]];
  }

  if (!s.allocDataStart || !s.allocTotalRow) { logError(`Alloc manquante: ${projet.code}`); return; }

  const ea = await ensureRows(ws, context, projet.membres.length, s.allocSlots, s.allocTotalRow);
  if (ea > 0) s = await readStructure(ws, context);

  for (let i = 0; i < s.allocSlots; i++) {
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 7).clear("Contents");
  }
  for (let i = 0; i < projet.membres.length; i++) {
    const m = projet.membres[i];
    ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
  }
  styleDataRows(ws, s.allocDataStart, projet.membres.length, 7, colors.pale);

  if (s.avancementRow && s.allocTotalRow) {
    const rc = s.avancementRow + 3;
    ws.getRangeByIndexes(rc - 1, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(rc - 1, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  await updatePhases(ws, context, s, projet, colors);
  await context.sync();
  log(`✓ ${projet.code} MAJ`);
}

// ============================================================
// NOUVEAU PROJET
// ============================================================

async function createNewProject(context, projet, reportPeriode) {
  log(`✚ ${projet.code} — ${projet.nom}`);

  const tplName = TPL_MAP[projet.categorie] || "_TPL_FACTURABLE";
  let tpl = context.workbook.worksheets.getItemOrNullObject(tplName);
  await context.sync();
  if (tpl.isNullObject) tpl = context.workbook.worksheets.getItem("_TEMPLATE");
  
  const ws = tpl.copy("End");
  await context.sync();
  ws.name = projet.code;
  ws.visibility = "Visible";
  ws.tabColor = projet.tabColor || "#077C79";

  ws.getRangeByIndexes(0, 0, 3, 7).clear("Contents");

  let s = await readStructure(ws, context);
  const n = projet.membres.length;
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];

  const ea = await ensureRows(ws, context, n, s.allocSlots, s.allocTotalRow);
  if (ea > 0) s = await readStructure(ws, context);
  const ep = await ensureRows(ws, context, n, s.planifSlots, s.planifTotalRow);
  if (ep > 0) s = await readStructure(ws, context);

  // TITRE
  ws.getRangeByIndexes(s.titleRow - 1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];

  // EN-TÊTE
  const clientName = (projet.client.toLowerCase() === "interne") ? "Atelier Urbain" : projet.client;
  if (s.clientRow) ws.getRangeByIndexes(s.clientRow - 1, 1, 1, 1).values = [[clientName]];
  if (s.catRow) ws.getRangeByIndexes(s.catRow - 1, 1, 1, 1).values = [[projet.categorie || "Facturable"]];

  // PÉRIODE
  if (s.periodeRow) {
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow - 1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow - 1, 2, 1, 1).values = [[reportPeriode.fin]];
  }

  // ALLOCATION
  if (s.allocDataStart) {
    for (let i = 0; i < n; i++) {
      const m = projet.membres[i];
      ws.getRangeByIndexes(s.allocDataStart - 1 + i, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
    }
    styleDataRows(ws, s.allocDataStart, n, 7, colors.pale);
  }

  // Formule Consommé
  if (s.avancementRow && s.allocTotalRow) {
    const rc = s.avancementRow + 3;
    ws.getRangeByIndexes(rc - 1, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(rc - 1, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  // PLANIF CAP
  if (s.planifDataStart) {
    for (let i = 0; i < n; i++) {
      ws.getRangeByIndexes(s.planifDataStart - 1 + i, 0, 1, 1).values = [[projet.membres[i].nom]];
    }
    styleDataRows(ws, s.planifDataStart, n, 14, colors.pale);
  }

  // PHASES
  await updatePhases(ws, context, s, projet, colors);

  // INDEX + DASHBOARD
  await addToIndex(context, projet, clientName);
  await addToDashboard(context, projet, s);

  await context.sync();
  log(`✓ ${projet.code} créé`);
}

// ============================================================
// PHASES
// ============================================================

async function updatePhases(ws, context, s, projet, colors) {
  if (!s.phasesDataStart) return;
  const rows = [];
  for (const m of projet.membres)
    for (const svc of m.services)
      rows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
  
  const cr = ws.getRangeByIndexes(s.phasesDataStart - 1, 0, 40, 5);
  cr.clear("Contents");
  cr.format.font.bold = false;
  cr.format.fill.clear();
  
  if (rows.length > 0) {
    ws.getRangeByIndexes(s.phasesDataStart - 1, 0, rows.length, 5).values = rows;
    styleDataRows(ws, s.phasesDataStart, rows.length, 5, colors.pale);
  }
}

// ============================================================
// INDEX — pas de doublon + alternance gris
// ============================================================

async function addToIndex(context, projet, clientName) {
  const sheet = context.workbook.worksheets.getItem(INDEX_SHEET);
  const range = sheet.getRangeByIndexes(0, 0, 100, 1);
  range.load("values");
  await context.sync();

  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim() === projet.code) return;
  }

  let headerIdx = -1;
  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim().startsWith("Code")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return;

  let insertIdx = headerIdx + 1;
  for (let i = headerIdx + 1; i < range.values.length; i++) {
    const v = String(range.values[i][0] || "").trim();
    if (!v) { insertIdx = i; break; }
    insertIdx = i + 1;
  }

  sheet.getRangeByIndexes(insertIdx, 0, 1, 11).values = [[
    projet.code, projet.nom, clientName || projet.client,
    projet.categorie || "Facturable", "En cours", "", "", "", 0, 0, "Import Nutcache"
  ]];

  // Alternance gris pâle
  const rowNum = insertIdx - headerIdx - 1; // 0-based index des données
  const bg = (rowNum % 2 === 0) ? "#F0F0F0" : "#FFFFFF";
  sheet.getRangeByIndexes(insertIdx, 0, 1, 11).format.fill.color = bg;
  sheet.getRangeByIndexes(insertIdx, 0, 1, 11).format.font.color = "#4A4A4A";
}

// ============================================================
// DASHBOARD — pas de doublon + couleur A + alternance
// ============================================================

async function addToDashboard(context, projet, projStruct) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await context.sync();
  if (sheet.isNullObject) return;

  const range = sheet.getRangeByIndexes(0, 0, 100, 1);
  range.load("values");
  await context.sync();

  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim() === projet.code) return;
  }

  let totalIdx = -1;
  for (let i = range.values.length - 1; i >= 0; i--) {
    if (String(range.values[i][0] || "").trim() === "TOTAL") { totalIdx = i; break; }
  }
  if (totalIdx === -1) return;

  sheet.getRangeByIndexes(totalIdx, 0, 1, 10).insert("Down");
  await context.sync();

  const safe = `'${projet.code}'`;
  const rBud = projStruct.avancementRow ? projStruct.avancementRow + 2 : 18;

  sheet.getRangeByIndexes(totalIdx, 0, 1, 3).values = [[
    projet.code, projet.nom, projet.categorie || "Facturable"
  ]];

  const fmls = [
    `=${safe}!B${rBud}`, `=${safe}!B${rBud+1}`, `=${safe}!B${rBud+2}`, `=${safe}!B${rBud+3}`,
    `=${safe}!C${rBud}`, `=${safe}!C${rBud+1}`, `=${safe}!C${rBud+2}`
  ];
  for (let c = 0; c < fmls.length; c++) {
    sheet.getRangeByIndexes(totalIdx, 3 + c, 1, 1).formulas = [[fmls[c]]];
  }

  // Couleur code A
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];
  const codeCell = sheet.getRangeByIndexes(totalIdx, 0, 1, 1);
  codeCell.format.fill.color = colors.bg;
  codeCell.format.font.color = colors.font;
  codeCell.format.font.bold = true;

  // Alternance gris pour colonnes B-J (sauf A qui a la couleur du type)
  // Compter combien de projets sont déjà dans le dashboard
  const rangeAfter = sheet.getRangeByIndexes(0, 0, 100, 1);
  rangeAfter.load("values");
  await context.sync();
  let dataRowCount = 0;
  let headerFound = false;
  for (let i = 0; i < rangeAfter.values.length; i++) {
    const v = String(rangeAfter.values[i][0] || "").trim();
    if (v === "Code") headerFound = true;
    else if (headerFound && v === "TOTAL") break;
    else if (headerFound && v) dataRowCount++;
  }
  const bg = ((dataRowCount - 1) % 2 === 0) ? "#F0F0F0" : "#FFFFFF";
  sheet.getRangeByIndexes(totalIdx, 1, 1, 9).format.fill.color = bg;
}

// ============================================================
// RESSOURCES — peupler dynamiquement
// ============================================================

async function updateResourcesList(context, projets) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(RESSOURCES_SHEET);
  await context.sync();
  if (sheet.isNullObject) return;

  // Lire noms existants (B7+)
  const range = sheet.getRangeByIndexes(6, 1, 50, 1);
  range.load("values");
  await context.sync();

  const existingNames = new Set();
  for (let i = 0; i < range.values.length; i++) {
    const nm = String(range.values[i][0] || "").trim();
    if (nm) existingNames.add(nm);
  }

  // Collecter tous les noms uniques
  const allNames = new Set();
  for (const p of projets) {
    for (const m of p.membres) {
      allNames.add(m.nom);
    }
  }

  // Nouveaux noms seulement
  const newNames = [];
  for (const nm of allNames) {
    if (!existingNames.has(nm)) newNames.push(nm);
  }
  if (newNames.length === 0) { log(`  Ressources: aucune nouvelle`); return; }

  // Trier : permanents alpha, puis externes
  const permanents = newNames.filter(n => !n.toLowerCase().includes("consultant") && !n.toLowerCase().includes("externe"));
  const externes = newNames.filter(n => n.toLowerCase().includes("consultant") || n.toLowerCase().includes("externe"));
  permanents.sort();
  externes.sort();
  const sorted = [...permanents, ...externes];

  // Trouver dernière ligne occupée
  let lastRow = 6;
  for (let i = 0; i < range.values.length; i++) {
    if (String(range.values[i][0] || "").trim()) lastRow = 7 + i;
  }

  log(`  Ressources: +${sorted.length} (${sorted.join(", ")})`);

  for (let i = 0; i < sorted.length; i++) {
    const r = lastRow + i; // 0-indexed
    const num = r - 5; // numéro séquentiel (R7=1, R8=2...)
    sheet.getRangeByIndexes(r, 0, 1, 3).values = [[num, sorted[i], 37.5]];
    const bg = ((r - 6) % 2 === 0) ? "#F0F0F0" : "#FFFFFF";
    sheet.getRangeByIndexes(r, 0, 1, 3).format.fill.color = bg;
    sheet.getRangeByIndexes(r, 0, 1, 3).format.font.color = "#4A4A4A";
    if (sorted[i].toLowerCase().includes("consultant") || sorted[i].toLowerCase().includes("externe")) {
      sheet.getRangeByIndexes(r, 0, 1, 3).format.font.italic = true;
    }
  }
}

// ============================================================
// TRI DES ONGLETS — corrigé
// ============================================================

async function sortProjectSheets(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name,items/position");
  await context.sync();

  const special = new Set(["Ressources", "Index Projets", "_TPL_FACTURABLE", "_TPL_INTERNE", 
    "_TPL_CONSULTANT", "_TEMPLATE", "Tableau de bord", "Planification capacitaire"]);

  const projectNames = sheets.items
    .filter(s => !special.has(s.name))
    .map(s => s.name);

  // Trier numériquement
  projectNames.sort((a, b) => {
    const na = parseInt(a.replace(/\D/g, "")) || 999999;
    const nb = parseInt(b.replace(/\D/g, "")) || 999999;
    if (na !== nb) return na - nb;
    return a.localeCompare(b);
  });

  // Position cible : après les templates, avant Tableau de bord
  // Trouver la position de Tableau de bord
  const tbIdx = sheets.items.findIndex(s => s.name === "Tableau de bord");
  
  if (tbIdx >= 0) {
    // Positionner chaque projet juste avant Tableau de bord
    for (let i = 0; i < projectNames.length; i++) {
      try {
        const ws = sheets.getItem(projectNames[i]);
        // On recharge les positions car elles changent après chaque move
        sheets.load("items/name,items/position");
        await context.sync();
        
        const currentTbIdx = sheets.items.findIndex(s => s.name === "Tableau de bord");
        ws.position = currentTbIdx;
        await context.sync();
      } catch(e) { /* skip */ }
    }
  }

  log(`  Onglets triés: ${projectNames.join(", ")}`);
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

      // Ressources
      await updateResourcesList(context, projets);

      // Tri des onglets
      try { await sortProjectSheets(context); } catch(e) { log(`  Tri: ${e.message}`); }

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
