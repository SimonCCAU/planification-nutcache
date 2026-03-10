/**
 * updater.js — v31
 * 
 * Tables Excel (ListObject) pour :
 * - Allocation des ressources
 * - Planification capacitaire du projet
 * - Détail par phase/service
 * - Notes de suivi
 * 
 * Avantages : alternance auto, formules TOTAL auto, extension auto
 */

const TPL_MAP = {
  "Facturable": "_TPL_FACTURABLE",
  "Non-facturable": "_TPL_INTERNE",
  "Consultant municipal": "_TPL_CONSULTANT"
};
const DASHBOARD_SHEET = "Tableau de bord";
const RESSOURCES_SHEET = "Ressources";
const PLANIF_CAP_SHEET = "Planification capacitaire";

const COLOR_MAP = {
  "Facturable": { bg: "#077C79", font: "#FFFFFF", pale: "#E6F2F1" },
  "Non-facturable": { bg: "#EB5E63", font: "#FFFFFF", pale: "#FDEEEF" },
  "Consultant municipal": { bg: "#EEDB4C", font: "#000000", pale: "#FDF9E5" }
};

function log(msg) {
  const el = document.getElementById("log");
  if (el) { const l = document.createElement("div"); l.textContent = `[${new Date().toLocaleTimeString("fr-CA")}] ${msg}`; el.appendChild(l); el.scrollTop = el.scrollHeight; }
  console.log(msg);
}
function logError(msg) {
  const el = document.getElementById("log");
  if (el) { const l = document.createElement("div"); l.className = "log-error"; l.textContent = `⚠ ${msg}`; el.appendChild(l); el.scrollTop = el.scrollHeight; }
  console.error(msg);
}

async function scanRows(ws, ctx, n) {
  n = n || 100;
  const r = ws.getRangeByIndexes(0, 0, n, 7); r.load("values"); await ctx.sync();
  return r.values.map(row => { for (let c=0;c<7;c++) { const v=String(row[c]||"").trim(); if(v) return v; } return ""; });
}
function findExact(a,t) { for(let i=0;i<a.length;i++) if(a[i]===t) return i+1; return null; }
function findPartial(a,t) { for(let i=0;i<a.length;i++) if(a[i].includes(t)) return i+1; return null; }
function findTotalAfter(a,r) { for(let i=r;i<a.length;i++) if(a[i]==="TOTAL") return i+1; return null; }

async function readStructure(ws, ctx) {
  const colA = await scanRows(ws, ctx, 100);
  const s = { colA };
  s.titleRow = findPartial(colA, "CODE") || findPartial(colA, "{{") || 4;
  s.clientRow = findExact(colA, "Client");
  s.catRow = findPartial(colA, "gorie") || findPartial(colA, "Categorie");
  s.periodeRow = findPartial(colA, "riode") || findPartial(colA, "Periode");
  s.avancementRow = findExact(colA, "AVANCEMENT");
  s.allocRow = findExact(colA, "ALLOCATION DES RESSOURCES");
  s.allocHeaderRow = s.allocRow ? s.allocRow+1 : null;
  s.allocDataStart = s.allocRow ? s.allocRow+2 : null;
  s.allocTotalRow = s.allocRow ? findTotalAfter(colA, s.allocRow-1) : null;
  s.allocSlots = (s.allocTotalRow && s.allocDataStart) ? s.allocTotalRow - s.allocDataStart : 0;
  s.planifRow = findPartial(colA, "PLANIFICATION CAPACITAIRE");
  s.planifHeaderRow = s.planifRow ? s.planifRow+2 : null;
  s.planifDataStart = s.planifRow ? s.planifRow+3 : null;
  s.planifTotalRow = s.planifRow ? findTotalAfter(colA, s.planifRow-1) : null;
  s.planifSlots = (s.planifTotalRow && s.planifDataStart) ? s.planifTotalRow - s.planifDataStart : 0;
  s.phasesRow = findPartial(colA, "TAIL PAR PHASE") || findPartial(colA, "PHASE");
  s.phasesHeaderRow = s.phasesRow ? s.phasesRow+1 : null;
  s.phasesDataStart = s.phasesRow ? s.phasesRow+2 : null;
  s.notesRow = findPartial(colA, "NOTES DE SUIVI");
  return s;
}

async function ensureRows(ws, ctx, needed, slots, totalRowIdx) {
  const extra = needed - slots;
  if (extra <= 0) return 0;
  ws.getRangeByIndexes(totalRowIdx-1, 0, extra, 14).insert("Down");
  await ctx.sync();
  return extra;
}

function styleDataRows(ws, startRow, count, cols, paleColor) {
  for (let i = 0; i < count; i++) {
    const r = startRow-1+i;
    const bg = (i%2===0) ? paleColor : "#FFFFFF";
    ws.getRangeByIndexes(r, 0, 1, cols).format.fill.color = bg;
    ws.getRangeByIndexes(r, 0, 1, cols).format.font.bold = false;
    ws.getRangeByIndexes(r, 0, 1, cols).format.font.color = "#4A4A4A";
    ws.getRangeByIndexes(r, 0, 1, cols).format.font.size = 10;
  }
}

function updateTotalFormulas(ws, dataStart, totalRow, cols) {
  const s = dataStart, e = totalRow - 1;
  for (const c of [2,3,4]) {
    if (c > cols) continue;
    const l = String.fromCharCode(64+c);
    ws.getRangeByIndexes(totalRow-1, c-1, 1, 1).formulas = [[`=SUM(${l}${s}:${l}${e})`]];
  }
  if (cols >= 5) ws.getRangeByIndexes(totalRow-1, 4, 1, 1).formulas = [[`=IF(COUNT(E${s}:E${e})=0,"",AVERAGE(E${s}:E${e}))`]];
  if (cols >= 6) ws.getRangeByIndexes(totalRow-1, 5, 1, 1).formulas = [[`=B${totalRow}-C${totalRow}`]];
}

function updatePlanifTotalFormulas(ws, dataStart, totalRow) {
  const s = dataStart, e = totalRow - 1;
  for (let c=2;c<=14;c++) {
    const l = String.fromCharCode(64+c);
    ws.getRangeByIndexes(totalRow-1, c-1, 1, 1).formulas = [[`=SUM(${l}${s}:${l}${e})`]];
  }
}

// ============================================================
async function getExistingProjectCodes(ctx) {
  const s = ctx.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET); await ctx.sync();
  if (s.isNullObject) return [];
  const colA = await scanRows(s, ctx, 100);
  return colA.filter(v => /^\d{5}$|^[A-Z]{2,4}-\d{3}$/.test(v));
}
async function getExistingSheetNames(ctx) {
  const s = ctx.workbook.worksheets; s.load("items/name"); await ctx.sync();
  return s.items.map(x => x.name);
}

// ============================================================
// MAJ EXISTANT
// ============================================================
async function updateExistingProject(ctx, projet, reportPeriode) {
  log(`↻ MAJ : ${projet.code}`);
  const ws = ctx.workbook.worksheets.getItem(projet.code);
  let s = await readStructure(ws, ctx);
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];

  if (s.periodeRow && reportPeriode.debut) {
    ws.getRangeByIndexes(s.periodeRow-1, 1, 1, 1).values = [[reportPeriode.debut]];
    ws.getRangeByIndexes(s.periodeRow-1, 2, 1, 1).values = [[reportPeriode.fin||""]];
  }
  if (!s.allocDataStart || !s.allocTotalRow) { logError(`Alloc manquante: ${projet.code}`); return; }

  const ea = await ensureRows(ws, ctx, projet.membres.length, s.allocSlots, s.allocTotalRow);
  if (ea > 0) s = await readStructure(ws, ctx);

  for (let i=0;i<s.allocSlots;i++) ws.getRangeByIndexes(s.allocDataStart-1+i, 0, 1, 7).clear("Contents");
  for (let i=0;i<projet.membres.length;i++) {
    const m = projet.membres[i];
    ws.getRangeByIndexes(s.allocDataStart-1+i, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
  }
  styleDataRows(ws, s.allocDataStart, projet.membres.length, 7, colors.pale);
  updateTotalFormulas(ws, s.allocDataStart, s.allocTotalRow, 6);

  if (s.avancementRow && s.allocTotalRow) {
    ws.getRangeByIndexes(s.avancementRow+2, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(s.avancementRow+2, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  await updatePhases(ws, ctx, s, projet, colors);
  await ctx.sync();
  log(`✓ ${projet.code} MAJ`);
}

// ============================================================
// NOUVEAU PROJET
// ============================================================
async function createNewProject(ctx, projet, reportPeriode) {
  log(`✚ ${projet.code} — ${projet.nom}`);

  const tplName = TPL_MAP[projet.categorie] || "_TPL_FACTURABLE";
  let tpl = ctx.workbook.worksheets.getItemOrNullObject(tplName);
  await ctx.sync();
  if (tpl.isNullObject) {
    logError(`Template ${tplName} introuvable`); return;
  }
  
  const ws = tpl.copy("End");
  await ctx.sync();
  ws.name = projet.code;
  ws.visibility = "Visible";
  ws.tabColor = projet.tabColor || "#077C79";
  ws.getRangeByIndexes(0, 0, 3, 7).clear("Contents");

  let s = await readStructure(ws, ctx);
  const n = projet.membres.length;
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];

  const ea = await ensureRows(ws, ctx, n, s.allocSlots, s.allocTotalRow);
  if (ea > 0) s = await readStructure(ws, ctx);
  const ep = await ensureRows(ws, ctx, n, s.planifSlots, s.planifTotalRow);
  if (ep > 0) s = await readStructure(ws, ctx);

  // TITRE
  ws.getRangeByIndexes(s.titleRow-1, 0, 1, 1).values = [[`${projet.code} — ${projet.nom}`]];

  // EN-TÊTE
  const clientName = (projet.client.toLowerCase() === "interne") ? "Atelier Urbain" : projet.client;
  if (s.clientRow) ws.getRangeByIndexes(s.clientRow-1, 1, 1, 1).values = [[clientName]];
  if (s.catRow) ws.getRangeByIndexes(s.catRow-1, 1, 1, 1).values = [[projet.categorie||"Facturable"]];

  // PÉRIODE
  if (s.periodeRow) {
    if (reportPeriode.debut) ws.getRangeByIndexes(s.periodeRow-1, 1, 1, 1).values = [[reportPeriode.debut]];
    if (reportPeriode.fin) ws.getRangeByIndexes(s.periodeRow-1, 2, 1, 1).values = [[reportPeriode.fin]];
  }

  // ALLOCATION
  if (s.allocDataStart) {
    for (let i=0;i<n;i++) {
      const m = projet.membres[i];
      ws.getRangeByIndexes(s.allocDataStart-1+i, 0, 1, 5).values = [[m.nom, 0, m.heures, m.couts, m.tauxDeduit]];
    }
    styleDataRows(ws, s.allocDataStart, n, 7, colors.pale);
    updateTotalFormulas(ws, s.allocDataStart, s.allocTotalRow, 6);
  }

  // Formule Consommé
  if (s.avancementRow && s.allocTotalRow) {
    ws.getRangeByIndexes(s.avancementRow+2, 1, 1, 1).formulas = [[`=C${s.allocTotalRow}`]];
    ws.getRangeByIndexes(s.avancementRow+2, 2, 1, 1).formulas = [[`=D${s.allocTotalRow}`]];
  }

  // PLANIF CAP
  if (s.planifDataStart) {
    for (let i=0;i<n;i++) {
      ws.getRangeByIndexes(s.planifDataStart-1+i, 0, 1, 1).values = [[projet.membres[i].nom]];
    }
    styleDataRows(ws, s.planifDataStart, n, 14, colors.pale);
    if (s.planifTotalRow) updatePlanifTotalFormulas(ws, s.planifDataStart, s.planifTotalRow);
  }

  // PHASES
  await updatePhases(ws, ctx, s, projet, colors);

  // DASHBOARD
  await addToDashboard(ctx, projet, s);

  await ctx.sync();
  log(`✓ ${projet.code} créé`);
}

// ============================================================
// PHASES
// ============================================================
async function updatePhases(ws, ctx, s, projet, colors) {
  if (!s.phasesDataStart) return;
  const rows = [];
  for (const m of projet.membres)
    for (const svc of m.services)
      rows.push([svc.nom, m.nom, svc.heures, svc.couts, svc.facturables]);
  ws.getRangeByIndexes(s.phasesDataStart-1, 0, 40, 5).clear("Contents");
  ws.getRangeByIndexes(s.phasesDataStart-1, 0, 40, 5).format.font.bold = false;
  ws.getRangeByIndexes(s.phasesDataStart-1, 0, 40, 5).format.fill.clear();
  if (rows.length > 0) {
    ws.getRangeByIndexes(s.phasesDataStart-1, 0, rows.length, 5).values = rows;
    styleDataRows(ws, s.phasesDataStart, rows.length, 5, colors.pale);
  }
}

// ============================================================
// DASHBOARD
// ============================================================
async function addToDashboard(ctx, projet, projStruct) {
  const sheet = ctx.workbook.worksheets.getItemOrNullObject(DASHBOARD_SHEET);
  await ctx.sync();
  if (sheet.isNullObject) return;
  const range = sheet.getRangeByIndexes(0, 0, 100, 1); range.load("values"); await ctx.sync();
  for (let i=0;i<range.values.length;i++) if(String(range.values[i][0]||"").trim()===projet.code) return;

  let totalIdx = -1;
  for (let i=range.values.length-1;i>=0;i--) if(String(range.values[i][0]||"").trim()==="TOTAL") { totalIdx=i; break; }
  if (totalIdx===-1) return;

  sheet.getRangeByIndexes(totalIdx, 0, 1, 10).insert("Down");
  await ctx.sync();

  const safe = `'${projet.code}'`;
  const rBud = projStruct.avancementRow ? projStruct.avancementRow+2 : 18;
  sheet.getRangeByIndexes(totalIdx, 0, 1, 3).values = [[projet.code, projet.nom, projet.categorie||"Facturable"]];
  const fmls = [
    `=${safe}!B${rBud}`, `=${safe}!B${rBud+1}`, `=${safe}!B${rBud+2}`, `=${safe}!B${rBud+3}`,
    `=${safe}!C${rBud}`, `=${safe}!C${rBud+1}`, `=${safe}!C${rBud+2}`
  ];
  for (let c=0;c<fmls.length;c++) sheet.getRangeByIndexes(totalIdx, 3+c, 1, 1).formulas = [[fmls[c]]];

  // Couleur code A
  const colors = COLOR_MAP[projet.categorie] || COLOR_MAP["Facturable"];
  sheet.getRangeByIndexes(totalIdx, 0, 1, 1).format.fill.color = colors.bg;
  sheet.getRangeByIndexes(totalIdx, 0, 1, 1).format.font.color = colors.font;
  sheet.getRangeByIndexes(totalIdx, 0, 1, 1).format.font.bold = true;

  // Alternance B-J
  const rangeAfter = sheet.getRangeByIndexes(0, 0, 100, 1); rangeAfter.load("values"); await ctx.sync();
  let count = 0; let hdr = false;
  for (let i=0;i<rangeAfter.values.length;i++) {
    const v = String(rangeAfter.values[i][0]||"").trim();
    if (v==="Code") hdr = true;
    else if (hdr && v==="TOTAL") break;
    else if (hdr && v) count++;
  }
  sheet.getRangeByIndexes(totalIdx, 1, 1, 9).format.fill.color = ((count-1)%2===0) ? "#F0F0F0" : "#FFFFFF";
  sheet.getRangeByIndexes(totalIdx, 1, 1, 9).format.font.color = "#1A1A1A";
}

// ============================================================
// RESSOURCES
// ============================================================
async function updateResourcesList(ctx, projets) {
  const sheet = ctx.workbook.worksheets.getItemOrNullObject(RESSOURCES_SHEET);
  await ctx.sync();
  if (sheet.isNullObject) return;

  const range = sheet.getRangeByIndexes(6, 1, 50, 1); range.load("values"); await ctx.sync();
  const existingNames = new Set();
  for (let i=0;i<range.values.length;i++) {
    const nm = String(range.values[i][0]||"").trim();
    if (nm) existingNames.add(nm);
  }

  const allNames = new Set();
  for (const p of projets) for (const m of p.membres) {
    if (!m.nom.toLowerCase().includes("consultant externe")) allNames.add(m.nom);
  }

  const newNames = [...allNames].filter(nm => !existingNames.has(nm));
  if (newNames.length === 0) { log(`  Ressources: aucune nouvelle`); return; }

  const permanents = newNames.filter(n => !n.toLowerCase().includes("consultant") && !n.toLowerCase().includes("externe"));
  const externes = newNames.filter(n => n.toLowerCase().includes("consultant") || n.toLowerCase().includes("externe"));
  permanents.sort(); externes.sort();
  const sorted = [...permanents, ...externes];

  let lastRow = 6;
  for (let i=0;i<range.values.length;i++) if(String(range.values[i][0]||"").trim()) lastRow = 7+i;

  log(`  Ressources: +${sorted.length}`);
  for (let i=0;i<sorted.length;i++) {
    const r = lastRow+i;
    sheet.getRangeByIndexes(r, 0, 1, 3).values = [[r-5, sorted[i], 37.5]];
  }
}

// ============================================================
// PLANIF CAP
// ============================================================
async function updatePlanifCapSheet(ctx) {
  const pcSheet = ctx.workbook.worksheets.getItemOrNullObject(PLANIF_CAP_SHEET);
  await ctx.sync();
  if (pcSheet.isNullObject) return;
  const resSheet = ctx.workbook.worksheets.getItemOrNullObject(RESSOURCES_SHEET);
  await ctx.sync();
  if (resSheet.isNullObject) return;

  const range = resSheet.getRangeByIndexes(6, 1, 50, 1); range.load("values"); await ctx.sync();
  const names = [];
  for (let i=0;i<range.values.length;i++) {
    const nm = String(range.values[i][0]||"").trim();
    if (nm) names.push({ nom: nm, resRow: 7+i });
  }
  if (names.length === 0) return;

  const pcRange = pcSheet.getRangeByIndexes(7, 0, 50, 1); pcRange.load("values"); await ctx.sync();
  const existingPC = new Set();
  let lastPCRow = 7;
  for (let i=0;i<pcRange.values.length;i++) {
    const nm = String(pcRange.values[i][0]||"").trim();
    if (nm) { existingPC.add(nm); lastPCRow = 8+i; }
  }

  let added = 0;
  for (const { nom, resRow } of names) {
    if (existingPC.has(nom)) continue;
    const r = lastPCRow + added;
    pcSheet.getRangeByIndexes(r, 0, 1, 1).values = [[nom]];
    pcSheet.getRangeByIndexes(r, 1, 1, 1).formulas = [[`=Ressources!C${resRow}*4`]];
    added++;
  }
  if (added > 0) log(`  Planif cap: +${added} ressources`);
}

// ============================================================
// TRI ONGLETS
// ============================================================
async function sortProjectSheets(ctx) {
  const sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  await ctx.sync();

  const beginOrder = ["Ressources", "Planification capacitaire", "_TPL_FACTURABLE", "_TPL_INTERNE",
    "_TPL_CONSULTANT", "Tableau de bord"];

  const special = new Set(beginOrder);
  const projectNames = sheets.items.filter(s => !special.has(s.name)).map(s => s.name);
  projectNames.sort((a, b) => {
    const na = parseInt(a.replace(/\D/g, "")) || 999999;
    const nb = parseInt(b.replace(/\D/g, "")) || 999999;
    return na !== nb ? na - nb : a.localeCompare(b);
  });

  for (let i = 0; i < beginOrder.length; i++) {
    try {
      const ws = sheets.getItemOrNullObject(beginOrder[i]);
      await ctx.sync();
      if (!ws.isNullObject) { ws.position = i; await ctx.sync(); }
    } catch(e) {}
  }

  for (let i = 0; i < projectNames.length; i++) {
    try {
      const ws = sheets.getItem(projectNames[i]);
      ws.position = beginOrder.length + i;
      await ctx.sync();
    } catch(e) {}
  }
  log(`  Onglets triés`);
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
    await Excel.run(async (ctx) => {
      const existingCodes = await getExistingProjectCodes(ctx);
      const existingSheets = await getExistingSheetNames(ctx);
      log(`Existants: ${existingCodes.join(", ") || "(aucun)"}`);

      let up=0, cr=0;
      for (const p of projets) {
        if (existingCodes.includes(p.code) && existingSheets.includes(p.code)) {
          await updateExistingProject(ctx, p, reportPeriode); up++;
        } else {
          await createNewProject(ctx, p, reportPeriode); cr++;
        }
      }

      await updateResourcesList(ctx, projets);
      await updatePlanifCapSheet(ctx);
      try { await sortProjectSheets(ctx); } catch(e) { log(`  Tri: ${e.message}`); }

      await ctx.sync();
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
