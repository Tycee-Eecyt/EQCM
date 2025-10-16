// Code.gs — Google Apps Script backend for EQ Character Manager
const CONFIG = {
  SECRET: '',              // Optional: set a shared secret string; leave blank to disable check
  ZONES_SHEET: 'Zone Tracker',
  FACTION_SHEET: 'CoV Faction',
  // Inventory summary tab name
  INV_SUMMARY_SHEET: 'Raid Kit',
  INV_ITEMS_SHEET: 'Inventory Items' // optional catch-all
};

function doPost(e){
  try{
    const body = JSON.parse(e.postData && e.postData.contents || '{}');
    if (CONFIG.SECRET && String(body.secret||'') !== CONFIG.SECRET) {
      return ContentService.createTextOutput('Unauthorized').setMimeType(ContentService.MimeType.TEXT).setResponseCode(401);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === 'pushInventorySheet'){
      return respond_(pushInventorySheet_(ss, body));
    }
    if (body.action === 'replaceAll'){
      replaceAll_(ss, body);
      return respond_({ ok: true, mode: 'replaceAll' });
    }
    if (body.action === 'listCharacters'){
      const res = listCharacters_(ss);
      return respond_(res);
    }
    // Replace CoV Faction tab with EXACT CSV contents
    if (body.action === 'replaceFactionsCsv'){
      const csv = String(body.csv || '');
      if (!csv) return respond_({ ok:false, error: 'Missing csv' }, 400);
      const res = replaceFactionsCsv_(ss, csv);
      return respond_({ ok: true, mode: 'replaceFactionsCsv', rows: res.rows, cols: res.cols });
    }
    // Block JSON writes to factions: accept only CSV action above
    if (body.action === 'replaceFactions'){
      return respond_({ ok:false, error: 'Factions JSON disabled; use action=replaceFactionsCsv' }, 403);
    }
    // Optional bulk JSON import for other tabs (faster path). CoV Faction is guarded and skipped here.
    if (body.action === 'jsonImport'){
      const res = jsonImport_(ss, body);
      return respond_(Object.assign({ ok:true, mode:'jsonImport' }, res));
    }

    const upserts = body.upserts || {};
    if (upserts.zones)    upsertZones_(ss, upserts.zones);
    // Factions JSON is ignored by design; only CSV is accepted to modify the CoV Faction tab
    if (upserts.inventory) upsertInventorySummary_(ss, upserts.inventory, body.meta || {});
    if (upserts.inventoryDetails) upsertInventoryDetails_(ss, upserts.inventoryDetails);
    return respond_({ ok: true });
  }catch(err){
    return respond_({ ok:false, error: String(err && err.stack || err) }, 500);
  }
}

function respond_(obj, code){
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  if (code) out.setResponseCode(code);
  return out;
}

function getOrMakeSheet_(ss, name){
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sh, header){
  const existing = sh.getRange(1,1,1,sh.getMaxColumns()).getValues()[0].filter(String);
  if (existing.join('\t') !== header.join('\t')){
    sh.clearContents();
    sh.getRange(1,1,1,header.length).setValues([header]);
  }
}

function upsertRowsByKey_(sh, header, keyCol, rows){
  ensureHeader_(sh, header);
  const idx = header.indexOf(keyCol) + 1;
  const data = sh.getDataRange().getValues();
  const map = new Map();
  for (let r=1; r<data.length; r++){
    const k = String(data[r][idx-1]||'');
    if (k) map.set(k, r+1);
  }
  rows.forEach(row => {
    const key = String(row[header.indexOf(keyCol)]||'');
    if (!key) return;
    let r = map.get(key);
    if (r){
      sh.getRange(r,1,1,header.length).setValues([row]);
    } else {
      sh.appendRow(row);
      map.set(key, sh.getLastRow());
    }
  });
}

function writeAllRows_(sh, header, rows){
  sh.clearContents();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if (rows && rows.length){
    sh.getRange(2,1,rows.length,header.length).setValues(rows);
  }
}

function normalizeZoneRow_(row, cols){
  const out = Array.isArray(row) ? row.slice(0, cols) : [];
  while (out.length < cols) out.push('');
  return out;
}

function buildZoneRowMap_(values, keyIdx, utcIdx, cols){
  const map = new Map();
  if (!Array.isArray(values)) return map;
  values.forEach(row => {
    if (!row || row.length <= keyIdx) return;
    const key = String(row[keyIdx] || '').trim();
    if (!key) return;
    map.set(key, { utc: row[utcIdx], values: normalizeZoneRow_(row, cols) });
  });
  return map;
}

function buildZoneRowIndexMap_(values, keyIdx, utcIdx, startRow, cols){
  const map = new Map();
  if (!Array.isArray(values)) return map;
  for (let i = 0; i < values.length; i++){
    const row = values[i];
    if (!row || row.length <= keyIdx) continue;
    const key = String(row[keyIdx] || '').trim();
    if (!key) continue;
    map.set(key, { rowIndex: startRow + i, utc: row[utcIdx], values: normalizeZoneRow_(row, cols) });
  }
  return map;
}

function parseDateValue_(value){
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && !isNaN(value)) return value;
  const str = String(value || '').trim();
  if (!str) return NaN;
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) return parsed;
  const cleaned = str.replace(/^\[|\]$/g, '');
  const parsedCleaned = Date.parse(cleaned);
  return isNaN(parsedCleaned) ? NaN : parsedCleaned;
}

function compareDateValues_(nextUtc, prevUtc){
  const nextMs = parseDateValue_(nextUtc);
  const prevMs = parseDateValue_(prevUtc);
  if (!isNaN(nextMs) && !isNaN(prevMs)){
    if (nextMs > prevMs) return 1;
    if (nextMs < prevMs) return -1;
    return 0;
  }
  if (!isNaN(nextMs)) return 1;
  if (!isNaN(prevMs)) return -1;
  const nextStr = String(nextUtc || '').trim();
  const prevStr = String(prevUtc || '').trim();
  if (nextStr && !prevStr) return 1;
  if (!nextStr && prevStr) return -1;
  if (nextStr > prevStr) return 1;
  if (nextStr < prevStr) return -1;
  return 0;
}

function upsertZones_(ss, rows){
  const header = ['Character','Last Zone','Zone Time (UTC)','Zone Time (Local)','Device TZ','Source Log File'];
  const sh = getOrMakeSheet_(ss, CONFIG.ZONES_SHEET);
  ensureHeader_(sh, header);
  const keyIdx = header.indexOf('Source Log File');
  const utcIdx = header.indexOf('Zone Time (UTC)');
  if (keyIdx < 0 || utcIdx < 0) return;

  const existingValues = sh.getDataRange().getValues();
  const existingMap = buildZoneRowIndexMap_(existingValues.slice(1), keyIdx, utcIdx, 2, header.length);

  rows.forEach(o => {
    const row = [o.character,o.zone,o.utc,o.local,o.tz,o.source];
    const key = String(row[keyIdx] || '').trim();
    if (!key) return;
    const entry = existingMap.get(key);
    if (entry){
      if (compareDateValues_(row[utcIdx], entry.utc) > 0){
        sh.getRange(entry.rowIndex, 1, 1, header.length).setValues([row]);
        entry.utc = row[utcIdx];
        entry.values = normalizeZoneRow_(row, header.length);
      }
    } else {
      sh.appendRow(row);
      const rowIndex = sh.getLastRow();
      existingMap.set(key, { rowIndex, utc: row[utcIdx], values: normalizeZoneRow_(row, header.length) });
    }
  });
}

function replaceZones_(ss, rows){
  const header = ['Character','Last Zone','Zone Time (UTC)','Zone Time (Local)','Device TZ','Source Log File'];
  const sh = getOrMakeSheet_(ss, CONFIG.ZONES_SHEET);
  const keyIdx = header.indexOf('Source Log File');
  const utcIdx = header.indexOf('Zone Time (UTC)');
  if (keyIdx < 0 || utcIdx < 0) {
    writeAllRows_(sh, header, (rows||[]).map(o => [o.character,o.zone,o.utc,o.local,o.tz,o.source]));
    return;
  }
  const existingValues = sh.getDataRange().getValues();
  const existingRows = existingValues.length > 1 ? existingValues.slice(1) : [];
  const existingMap = buildZoneRowMap_(existingRows, keyIdx, utcIdx, header.length);

  const incoming = (rows||[]).map(o => [o.character,o.zone,o.utc,o.local,o.tz,o.source]);
  const finalRows = [];
  const seen = new Set();

  incoming.forEach(row => {
    const key = String(row[keyIdx] || '').trim();
    if (!key) {
      finalRows.push(row);
      return;
    }
    if (seen.has(key)) return;
    seen.add(key);
    const existing = existingMap.get(key);
    if (!existing){
      finalRows.push(row);
      return;
    }
    const cmp = compareDateValues_(row[utcIdx], existing.utc);
    // Preserve the sheet's value whenever it already holds the newer timestamp.
    if (cmp > 0){
      finalRows.push(row);
      return;
    }
    finalRows.push(existing.values.slice());
  });

  writeAllRows_(sh, header, finalRows);
}

function upsertFactions_(ss, rows){
  const header = ['Character','Standing','Score','Mob','Consider Time (UTC)','Consider Time (Local)','Notes'];
  const sh = getOrMakeSheet_(ss, CONFIG.FACTION_SHEET);
  const data = rows.map(o => [o.character,o.standing,o.score,o.mob,o.utc,o.local,(o.standingDisplay||'')]);
  upsertRowsByKey_(sh, header, 'Character', data);
}

function replaceFactions_(ss, rows){
  const header = ['Character','Standing','Score','Mob','Consider Time (UTC)','Consider Time (Local)','Notes'];
  const sh = getOrMakeSheet_(ss, CONFIG.FACTION_SHEET);
  const data = (rows||[]).map(o => [o.character,o.standing,o.score,o.mob,o.utc,o.local,(o.standingDisplay||'')]);
  writeAllRows_(sh, header, data);
}

// Replace CoV Faction with the EXACT contents of a CSV string
function replaceFactionsCsv_(ss, csvText){
  const sh = getOrMakeSheet_(ss, CONFIG.FACTION_SHEET);
  const rows = Utilities.parseCsv(csvText || '');
  // Clear and write exactly what the CSV contains (including header row as-is)
  sh.clearContents();
  if (rows && rows.length){
    sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
    return { rows: rows.length, cols: rows[0].length };
  }
  return { rows: 0, cols: 0 };
}

function upsertInventorySummary_(ss, rows, meta){
  const fixedHeaders = (meta && meta.invFixedHeaders && meta.invFixedHeaders.length) ? meta.invFixedHeaders : [
    'Vial of Velium Vapors','Leatherfoot Raider Skullcap','Shiny Brass Idol','Ring of Shadows Count',
    'Reaper of the Dead','Pearl Count','Peridot Count','10 Dose Potion of Stinging Wort Count','Pegasus Feather Cloak',
    'MB Class Five','MB Class Four','MB Class Three','MB Class Two','MB Class One','Larrikan\'s Mask'
  ];
  const fixedProps = (meta && meta.invFixedProps && meta.invFixedProps.length) ? meta.invFixedProps : [
    'vialVeliumVapors','leatherfootSkullcap','shinyBrassIdol','ringOfShadowsCount',
    'reaperOfTheDead','pearlCount','peridotCount','tenDosePotionOfStingingWortCount','pegasusFeatherCloak',
    'mbClassFive','mbClassFour','mbClassThree','mbClassTwo','mbClassOne','larrikansMask'
  ];
  const base = ['Character','Inventory File','Source Log File','Created (UTC)','Modified (UTC)'].concat(fixedHeaders);
  // Union any extra kit columns provided as o.kitExtras { HeaderLabel: value }
  const extrasSet = new Set();
  (rows||[]).forEach(o => { const ex = o.kitExtras||{}; Object.keys(ex).forEach(k => extrasSet.add(String(k))); });
  const extras = Array.from(extrasSet);
  const header = base.concat(extras);
  const sh = getOrMakeSheet_(ss, CONFIG.INV_SUMMARY_SHEET);
  const data = (rows||[]).map(o => {
    const fixedVals = fixedProps.map(p => o.raidKit && (o.raidKit[p] ?? '') || '');
    const baseVals = [o.character,o.file,o.logFile,o.created,o.modified].concat(fixedVals);
    const ex = o.kitExtras||{};
    const extraVals = extras.map(h => ex[h] ?? '');
    return baseVals.concat(extraVals);
  });
  upsertRowsByKey_(sh, header, 'Character', data);
}

function replaceInventorySummary_(ss, rows, meta){
  const fixedHeaders = (meta && meta.invFixedHeaders && meta.invFixedHeaders.length) ? meta.invFixedHeaders : [
    'Vial of Velium Vapors','Leatherfoot Raider Skullcap','Shiny Brass Idol','Ring of Shadows Count',
    'Reaper of the Dead','Pearl Count','Peridot Count','10 Dose Potion of Stinging Wort Count','Pegasus Feather Cloak',
    'MB Class Five','MB Class Four','MB Class Three','MB Class Two','MB Class One','Larrikan\'s Mask'
  ];
  const fixedProps = (meta && meta.invFixedProps && meta.invFixedProps.length) ? meta.invFixedProps : [
    'vialVeliumVapors','leatherfootSkullcap','shinyBrassIdol','ringOfShadowsCount',
    'reaperOfTheDead','pearlCount','peridotCount','tenDosePotionOfStingingWortCount','pegasusFeatherCloak',
    'mbClassFive','mbClassFour','mbClassThree','mbClassTwo','mbClassOne','larrikansMask'
  ];
  const base = ['Character','Inventory File','Source Log File','Created (UTC)','Modified (UTC)'].concat(fixedHeaders);
  const extrasSet = new Set();
  (rows||[]).forEach(o => { const ex = o.kitExtras||{}; Object.keys(ex).forEach(k => extrasSet.add(String(k))); });
  const extras = Array.from(extrasSet);
  const header = base.concat(extras);
  const sh = getOrMakeSheet_(ss, CONFIG.INV_SUMMARY_SHEET);
  const data = (rows||[]).map(o => {
    const fixedVals = fixedProps.map(p => o.raidKit && (o.raidKit[p] ?? '') || '');
    const baseVals = [o.character,o.file,o.logFile,o.created,o.modified].concat(fixedVals);
    const ex = o.kitExtras||{};
    const extraVals = extras.map(h => ex[h] ?? '');
    return baseVals.concat(extraVals);
  });
  writeAllRows_(sh, header, data);
}

function upsertInventoryDetails_(ss, rows){
  const header = ['Character','Inventory File','Created (UTC)','Modified (UTC)','Location','Name','ID','Count','Slots'];
  const sh = getOrMakeSheet_(ss, CONFIG.INV_ITEMS_SHEET);
  ensureHeader_(sh, header);
  const data = [];
  rows.forEach(o => {
    (o.items||[]).forEach(it => {
      data.push([o.character, o.file, o.created, o.modified, it.Location, it.Name, it.ID, it.Count, it.Slots]);
    });
  });
  if (data.length) sh.getRange(sh.getLastRow()+1,1,data.length,header.length).setValues(data);
}

function replaceInventoryDetails_(ss, rows){
  const header = ['Character','Inventory File','Created (UTC)','Modified (UTC)','Location','Name','ID','Count','Slots'];
  const sh = getOrMakeSheet_(ss, CONFIG.INV_ITEMS_SHEET);
  const data = [];
  (rows||[]).forEach(o => {
    (o.items||[]).forEach(it => {
      data.push([o.character, o.file, o.created, o.modified, it.Location, it.Name, it.ID, it.Count, it.Slots]);
    });
  });
  writeAllRows_(sh, header, data);
}

function replaceAll_(ss, body){
  const up = body.upserts || body || {};
  if (up.zones)    replaceZones_(ss, up.zones);
  // Factions JSON is ignored; use replaceFactionsCsv action instead
  if (up.inventory) replaceInventorySummary_(ss, up.inventory, body.meta || {});
  if (up.inventoryDetails) replaceInventoryDetails_(ss, up.inventoryDetails);
}

// Return unique character names currently present on any of the primary tabs.
function listCharacters_(ss){
  try{
    const tabs = [CONFIG.ZONES_SHEET, CONFIG.FACTION_SHEET, CONFIG.INV_SUMMARY_SHEET];
    const set = new Set();
    tabs.forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const data = sh.getDataRange().getValues();
      if (!data || data.length < 2) return; // header + at least one row
      const header = data[0];
      const idx = header.indexOf('Character');
      if (idx < 0) return;
      for (let r=1; r<data.length; r++){
        const n = String(data[r][idx]||'').trim();
        if (n) set.add(n);
      }
    });
    return { ok:true, characters: Array.from(set).sort() };
  } catch(err){ return { ok:false, error: String(err && err.message || err) }; }
}

// Create a NEW sheet tab with a character's inventory (similar to screenshot)
function pushInventorySheet_(ss, body){
  const character = String(body.character||'').trim();
  const baseName = String(body.sheetName||'').trim() || ('Inventory - ' + character);
  const info = body.info || {};
  const items = body.items || [];
  if (!character) return { ok:false, error: 'Missing character' };

  // Unique sheet name
  let name = baseName, n=1;
  while (ss.getSheetByName(name)) { name = baseName + ' (' + (++n) + ')'; }
  const sh = ss.insertSheet(name);

  // Header rows (A1:D1, A2:D2)
  sh.getRange(1,1,1,4).setValues([['Inventory for','File','Created On','Modified On']]);
  sh.getRange(2,1,1,4).setValues([[character, info.file || '', info.created || '', info.modified || '']]);

  // Blank spacer row
  sh.getRange(3,1,1,1).setValue('');

  // Table header & data
  const header = ['Location','Name','ID','Count','Slots'];
  sh.getRange(5,1,1,header.length).setValues([header]);
  const data = items.map(it => [it.Location, it.Name, it.ID, it.Count, it.Slots]);
  if (data.length) sh.getRange(6,1,data.length,header.length).setValues(data);

  // Formatting
  sh.setFrozenRows(5);
  sh.autoResizeColumns(1, header.length);

  return { ok:true, sheet: name };
}

// Bulk JSON import for non-faction tabs. Each tab: { sheet(name), header:[], rows:[[]], mode:'replace'|'append' }
// CoV Faction sheet is intentionally skipped here to prevent JSON-based overwrites; use replaceFactionsCsv.
function jsonImport_(ss, body){
  const tabs = Array.isArray(body && body.tabs) ? body.tabs : [];
  let imported = 0;
  tabs.forEach(tab => {
    try{
      const name = String(tab.sheet || tab.name || '').trim();
      if (!name) return;
      if (name === CONFIG.FACTION_SHEET) return; // guard
      const header = Array.isArray(tab.header) ? tab.header : [];
      const rows = Array.isArray(tab.rows) ? tab.rows : [];
      const sh = getOrMakeSheet_(ss, name);
      const mode = String(tab.mode||'replace').toLowerCase();
      if (mode === 'append'){
        ensureHeader_(sh, header);
        if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,header.length).setValues(rows);
      } else {
        writeAllRows_(sh, header, rows);
      }
      imported++;
    } catch(err){}
  });
  return { imported };
}

