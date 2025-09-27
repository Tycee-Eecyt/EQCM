// Code.gs — Google Apps Script backend for EQ Character Manager
const CONFIG = {
  SECRET: '',              // Optional: set a shared secret string; leave blank to disable check
  ZONES_SHEET: 'Zone Tracker',
  FACTION_SHEET: 'CoV Faction',
  INV_SUMMARY_SHEET: 'Inventory Summary',
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

    const upserts = body.upserts || {};
    if (upserts.zones)    upsertZones_(ss, upserts.zones);
    if (upserts.factions) upsertFactions_(ss, upserts.factions);
    if (upserts.inventory) upsertInventorySummary_(ss, upserts.inventory);
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

function upsertZones_(ss, rows){
  const header = ['Character','Last Zone','Zone Time (UTC)','Zone Time (Local)','Device TZ','Source Log File'];
  const sh = getOrMakeSheet_(ss, CONFIG.ZONES_SHEET);
  const data = rows.map(o => [o.character,o.zone,o.utc,o.local,o.tz,o.source]);
  upsertRowsByKey_(sh, header, 'Character', data);
}

function upsertFactions_(ss, rows){
  const header = ['Character','Standing','Score','Mob','Consider Time (UTC)','Consider Time (Local)','Notes'];
  const sh = getOrMakeSheet_(ss, CONFIG.FACTION_SHEET);
  const data = rows.map(o => [o.character,o.standing,o.score,o.mob,o.utc,o.local,(o.standingDisplay||'')]);
  upsertRowsByKey_(sh, header, 'Character', data);
}

function upsertInventorySummary_(ss, rows){
  const header = ['Character','Inventory File','Source Log File','Created (UTC)','Modified (UTC)',
                  'Vial of Velium Vapors','Velium Vial Count','Leatherfoot Raider Skullcap','Shiny Brass Idol',
                  'Ring of Shadows Count','Reaper of the Dead','Pearl Count','Peridot Count',
                  'MB Class Five','MB Class Four','MB Class Three','MB Class Two','MB Class One','Larrikan\'s Mask'];
  const sh = getOrMakeSheet_(ss, CONFIG.INV_SUMMARY_SHEET);
  const data = rows.map(o => [o.character,o.file,o.logFile,o.created,o.modified,
                              o.raidKit?.vialVeliumVapors||'', o.raidKit?.veliumVialCount||0, o.raidKit?.leatherfootSkullcap||'',
                              o.raidKit?.shinyBrassIdol||'', o.raidKit?.ringOfShadowsCount||0, o.raidKit?.reaperOfTheDead||'',
                              o.raidKit?.pearlCount||0, o.raidKit?.peridotCount||0,
                              o.raidKit?.mbClassFive||0, o.raidKit?.mbClassFour||0, o.raidKit?.mbClassThree||0,
                              o.raidKit?.mbClassTwo||0, o.raidKit?.mbClassOne||0, o.raidKit?.larrikansMask||'' ]);
  upsertRowsByKey_(sh, header, 'Character', data);
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

