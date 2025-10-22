const express = require('express');
const { google } = require('googleapis');
const { parse } = require('csv-parse/sync');
const dotenv = require('dotenv');
const { storeWebhookPayload, recordRawWebhook } = require('./db/payload-store');
const { markCharacterSynced } = require('./db/summaries');
const { startScheduler, scheduleCharacterSync, enqueueFullBackfill } = require('./jobs/scheduler');
const { startWorker } = require('./jobs/worker');
dotenv.config();

const CONFIG = {
  SECRET: process.env.WEBHOOK_SECRET || '',
  ZONES_SHEET: process.env.ZONES_SHEET_NAME || 'Zone Tracker',
  FACTION_SHEET: process.env.FACTION_SHEET_NAME || 'CoV Faction',
  INV_SUMMARY_SHEET: process.env.INVENTORY_SUMMARY_SHEET_NAME || 'Raid Kit',
  INV_ITEMS_SHEET: process.env.INVENTORY_ITEMS_SHEET_NAME || 'Inventory Items'
};

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) {
  console.log('INFO: GOOGLE_SHEET_ID/SPREADSHEET_ID not set; expecting clients to supply sheetId or sheetUrl in webhook payloads.');
}

const GOOGLE_SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

function resolveServiceAccountCredentials() {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    try {
      return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    } catch (err) {
      throw new Error(`Failed to parse GOOGLE_SERVICE_ACCOUNT_JSON: ${err.message}`);
    }
  }
  return undefined;
}

const authOptions = {
  scopes: GOOGLE_SCOPES
};

if (process.env.GOOGLE_APPLICATION_CREDENTIALS && !process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  authOptions.keyFilename = process.env.GOOGLE_APPLICATION_CREDENTIALS;
}

const serviceAccountCreds = resolveServiceAccountCredentials();
if (serviceAccountCreds) {
  authOptions.credentials = serviceAccountCreds;
}

const googleAuth = new google.auth.GoogleAuth(authOptions);

async function getSheetsClient() {
  const authClient = await googleAuth.getClient();
  return google.sheets({ version: 'v4', auth: authClient });
}

const sheetExistsCache = new Set();

async function ensureSheetExists(sheets, spreadsheetId, title) {
  const cacheKey = `${spreadsheetId}::${title}`;
  if (sheetExistsCache.has(cacheKey)) return;
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId
  });
  const exists = Array.isArray(spreadsheet.data.sheets)
    ? spreadsheet.data.sheets.some((sheet) => sheet.properties && sheet.properties.title === title)
    : false;
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: { properties: { title } }
          }
        ]
      }
    });
  }
  sheetExistsCache.add(cacheKey);
}

function sheetRange(title) {
  return `'${title.replace(/'/g, "''")}'`;
}

function sheetRangeAll(title) {
  return `${sheetRange(title)}!A1:ZZZ`;
}

function extractSheetId(value) {
  const str = String(value || '').trim();
  if (!str) return '';
  if (/^[A-Za-z0-9_-]{20,}$/.test(str)) return str;
  try {
    const url = new URL(/^https?:\/\//i.test(str) ? str : `https://${str}`);
    const parts = url.pathname.split('/').filter(Boolean);
    const idx = parts.indexOf('d');
    if (idx >= 0 && parts[idx + 1]) return parts[idx + 1];
    return '';
  } catch {
    return '';
  }
}

async function getSheetWithHeader(sheets, spreadsheetId, title) {
  await ensureSheetExists(sheets, spreadsheetId, title);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: sheetRange(title)
  });
  const values = res.data.values || [];
  if (!values.length) {
    return { header: [], rows: [] };
  }
  const [header, ...rows] = values;
  return { header, rows };
}

async function writeSheet(sheets, spreadsheetId, title, header, rows) {
  const body = [];
  if (header && header.length) body.push(header);
  if (rows && rows.length) {
    rows.forEach((row) => body.push(row));
  }
  const sanitized = body.map((row) =>
    row.map((cell) => sanitizeCellValue(cell))
  );
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: sheetRangeAll(title)
  });
  if (sanitized.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetRange(title)}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: sanitized }
    });
  }
}

function normalizeRow(row, cols) {
  const out = Array.isArray(row) ? row.slice(0, cols) : [];
  while (out.length < cols) out.push('');
  return out;
}

function sanitizeCellValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') {
    return value.length <= MAX_SHEET_CELL_LENGTH
      ? value
      : value.slice(0, MAX_SHEET_CELL_LENGTH);
  }
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  const str = String(value);
  return str.length <= MAX_SHEET_CELL_LENGTH
    ? str
    : str.slice(0, MAX_SHEET_CELL_LENGTH);
}

function sanitizeRows(rows) {
  return (rows || []).map((row) => row.map((cell) => sanitizeCellValue(cell)));
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && !Number.isNaN(value)) return value;
  const str = String(value || '').trim();
  if (!str) return Number.NaN;
  const parsed = Date.parse(str);
  if (!Number.isNaN(parsed)) return parsed;
  const cleaned = str.replace(/^\[|\]$/g, '');
  const parsedCleaned = Date.parse(cleaned);
  return Number.isNaN(parsedCleaned) ? Number.NaN : parsedCleaned;
}

function compareDateValues(nextUtc, prevUtc) {
  const nextMs = parseDateValue(nextUtc);
  const prevMs = parseDateValue(prevUtc);
  if (!Number.isNaN(nextMs) && !Number.isNaN(prevMs)) {
    if (nextMs > prevMs) return 1;
    if (nextMs < prevMs) return -1;
    return 0;
  }
  if (!Number.isNaN(nextMs)) return 1;
  if (!Number.isNaN(prevMs)) return -1;
  const nextStr = String(nextUtc || '').trim();
  const prevStr = String(prevUtc || '').trim();
  if (nextStr && !prevStr) return 1;
  if (!nextStr && prevStr) return -1;
  if (nextStr > prevStr) return 1;
  if (nextStr < prevStr) return -1;
  return 0;
}

function getBestTimestamp(values) {
  let bestStr = '';
  let bestMs = Number.NaN;
  (Array.isArray(values) ? values : []).forEach((val) => {
    const str = String(val || '').trim();
    if (!str) return;
    const ms = parseDateValue(str);
    if (!Number.isNaN(ms)) {
      if (Number.isNaN(bestMs) || ms > bestMs) {
        bestMs = ms;
        bestStr = str;
      }
    } else if (!bestStr) {
      bestStr = str;
    }
  });
  return { value: bestStr, ms: bestMs };
}

function keyFromRow(row, keyIdx) {
  if (keyIdx < 0) return '';
  return String(row[keyIdx] || '').trim();
}

function remapRowsToHeader(rows, originalHeader, targetHeader) {
  if (!rows.length) return [];
  const indexMap = targetHeader.map((label) => originalHeader.indexOf(label));
  return rows.map((row) =>
    targetHeader.map((_, idx) => {
      const sourceIdx = indexMap[idx];
      if (sourceIdx === -1) return '';
      return row[sourceIdx] !== undefined ? row[sourceIdx] : '';
    })
  );
}

function mergeForUpsert(existingRows, incomingRows, keyIdx, dateIdxs, headerLength) {
  const order = [];
  const map = new Map();
  const finalRows = [];

  existingRows.forEach((raw) => {
    const row = normalizeRow(raw, headerLength);
    const key = keyFromRow(row, keyIdx);
    if (!key) {
      finalRows.push(row);
      return;
    }
    const ts = getBestTimestamp(dateIdxs.map((idx) => row[idx]));
    map.set(key, { row, ts });
    order.push(key);
  });

  incomingRows.forEach((raw) => {
    const row = normalizeRow(raw, headerLength);
    const key = keyFromRow(row, keyIdx);
    if (!key) {
      finalRows.push(row);
      return;
    }
    const ts = getBestTimestamp(dateIdxs.map((idx) => row[idx]));
    const existing = map.get(key);
    if (!existing) {
      map.set(key, { row, ts });
      order.push(key);
      return;
    }
    if (compareDateValues(ts.value, existing.ts.value) > 0) {
      map.set(key, { row, ts });
    }
  });

  order.forEach((key) => {
    const entry = map.get(key);
    if (entry) {
      finalRows.push(entry.row.slice());
    }
  });

  return finalRows;
}

function mergeForReplace(existingRows, incomingRows, keyIdx, dateIdxs, headerLength) {
  const existingMap = new Map();
  existingRows.forEach((raw) => {
    const row = normalizeRow(raw, headerLength);
    const key = keyFromRow(row, keyIdx);
    if (!key) return;
    const ts = getBestTimestamp(dateIdxs.map((idx) => row[idx]));
    existingMap.set(key, { row, ts });
  });

  const finalRows = [];
  incomingRows.forEach((raw) => {
    const row = normalizeRow(raw, headerLength);
    const key = keyFromRow(row, keyIdx);
    if (!key) {
      finalRows.push(row);
      return;
    }
    const ts = getBestTimestamp(dateIdxs.map((idx) => row[idx]));
    const existing = existingMap.get(key);
    if (existing && compareDateValues(ts.value, existing.ts.value) < 0) {
      finalRows.push(existing.row.slice());
    } else {
      finalRows.push(row);
    }
  });

  return finalRows;
}

async function upsertZones(sheets, spreadsheetId, rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const header = ['Character', 'Last Zone', 'Zone Time (UTC)', 'Zone Time (Local)', 'Device TZ', 'Source Log File'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.ZONES_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = rows.map((o) => [
    o.character,
    o.zone,
    o.utc,
    o.local,
    o.tz,
    o.source
  ]);
  const merged = mergeForUpsert(existingRows, data, header.indexOf('Source Log File'), [header.indexOf('Zone Time (UTC)')], header.length);
  await writeSheet(sheets, spreadsheetId, CONFIG.ZONES_SHEET, header, merged);
}

async function replaceZones(sheets, spreadsheetId, rows) {
  const header = ['Character', 'Last Zone', 'Zone Time (UTC)', 'Zone Time (Local)', 'Device TZ', 'Source Log File'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.ZONES_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = (Array.isArray(rows) ? rows : []).map((o) => [
    o.character,
    o.zone,
    o.utc,
    o.local,
    o.tz,
    o.source
  ]);
  const merged = mergeForReplace(existingRows, data, header.indexOf('Source Log File'), [header.indexOf('Zone Time (UTC)')], header.length);
  await writeSheet(sheets, spreadsheetId, CONFIG.ZONES_SHEET, header, merged);
}

async function upsertFactions(sheets, spreadsheetId, rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const header = ['Character', 'Standing', 'Score', 'Mob', 'Consider Time (UTC)', 'Consider Time (Local)', 'Notes'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.FACTION_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = rows.map((o) => [
    o.character,
    o.standing,
    o.score,
    o.mob,
    o.utc,
    o.local,
    o.standingDisplay || ''
  ]);
  const merged = mergeForUpsert(
    existingRows,
    data,
    header.indexOf('Character'),
    [header.indexOf('Consider Time (UTC)'), header.indexOf('Consider Time (Local)')],
    header.length
  );
  await writeSheet(sheets, spreadsheetId, CONFIG.FACTION_SHEET, header, merged);
}

async function replaceFactionsJson(sheets, spreadsheetId, rows) {
  const header = ['Character', 'Standing', 'Score', 'Mob', 'Consider Time (UTC)', 'Consider Time (Local)', 'Notes'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.FACTION_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = (Array.isArray(rows) ? rows : []).map((o) => [
    o.character,
    o.standing,
    o.score,
    o.mob,
    o.utc,
    o.local,
    o.standingDisplay || ''
  ]);
  const merged = mergeForReplace(
    existingRows,
    data,
    header.indexOf('Character'),
    [header.indexOf('Consider Time (UTC)'), header.indexOf('Consider Time (Local)')],
    header.length
  );
  await writeSheet(sheets, spreadsheetId, CONFIG.FACTION_SHEET, header, merged);
}

function computeInventoryHeaders(rows, meta, existingHeader) {
  const fixedHeaders =
    meta && Array.isArray(meta.invFixedHeaders) && meta.invFixedHeaders.length
      ? meta.invFixedHeaders
      : [
          'Vial of Velium Vapors',
          'Leatherfoot Raider Skullcap',
          'Shiny Brass Idol',
          'Ring of Shadows Count',
          'Reaper of the Dead',
          'Pearl Count',
          'Peridot Count',
          '10 Dose Potion of Stinging Wort Count',
          'Pegasus Feather Cloak',
          'MB Class Five',
          'MB Class Four',
          'MB Class Three',
          'MB Class Two',
          'MB Class One',
          "Larrikan's Mask"
        ];
  const fixedProps =
    meta && Array.isArray(meta.invFixedProps) && meta.invFixedProps.length
      ? meta.invFixedProps
      : [
          'vialVeliumVapors',
          'leatherfootSkullcap',
          'shinyBrassIdol',
          'ringOfShadowsCount',
          'reaperOfTheDead',
          'pearlCount',
          'peridotCount',
          'tenDosePotionOfStingingWortCount',
          'pegasusFeatherCloak',
          'mbClassFive',
          'mbClassFour',
          'mbClassThree',
          'mbClassTwo',
          'mbClassOne',
          'larrikansMask'
        ];
  const base = ['Character', 'Inventory File', 'Source Log File', 'Created (UTC)', 'Modified (UTC)'].concat(fixedHeaders);

  const existingExtras = Array.isArray(existingHeader) && existingHeader.length > base.length
    ? existingHeader.slice(base.length)
    : [];
  const extrasSet = new Set(existingExtras);
  (Array.isArray(rows) ? rows : []).forEach((o) => {
    const extra = o.kitExtras || {};
    Object.keys(extra).forEach((label) => extrasSet.add(String(label)));
  });
  const extras = [...existingExtras];
  extrasSet.forEach((label) => {
    if (!extras.includes(label)) extras.push(label);
  });

  return { header: base.concat(extras), fixedProps, extras };
}

async function upsertInventorySummary(sheets, spreadsheetId, rows, meta) {
  if (!Array.isArray(rows) || !rows.length) return;
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.INV_SUMMARY_SHEET);
  const { header, fixedProps, extras } = computeInventoryHeaders(rows, meta, existingHeader);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = rows.map((o) => {
    const fixedVals = fixedProps.map((prop) => (o.raidKit && Object.prototype.hasOwnProperty.call(o.raidKit, prop) ? o.raidKit[prop] : ''));
    const baseVals = [o.character, o.file, o.logFile, o.created, o.modified].concat(fixedVals);
    const ex = o.kitExtras || {};
    const extraVals = extras.map((label) => ex[label] ?? '');
    return baseVals.concat(extraVals);
  });
  const merged = mergeForUpsert(
    existingRows,
    data,
    header.indexOf('Character'),
    [header.indexOf('Created (UTC)'), header.indexOf('Modified (UTC)')],
    header.length
  );
  await writeSheet(sheets, spreadsheetId, CONFIG.INV_SUMMARY_SHEET, header, merged);
}

async function replaceInventorySummary(sheets, spreadsheetId, rows, meta) {
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.INV_SUMMARY_SHEET);
  const { header, fixedProps, extras } = computeInventoryHeaders(rows, meta, existingHeader);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const data = (Array.isArray(rows) ? rows : []).map((o) => {
    const fixedVals = fixedProps.map((prop) => (o.raidKit && Object.prototype.hasOwnProperty.call(o.raidKit, prop) ? o.raidKit[prop] : ''));
    const baseVals = [o.character, o.file, o.logFile, o.created, o.modified].concat(fixedVals);
    const ex = o.kitExtras || {};
    const extraVals = extras.map((label) => ex[label] ?? '');
    return baseVals.concat(extraVals);
  });
  const merged = mergeForReplace(
    existingRows,
    data,
    header.indexOf('Character'),
    [header.indexOf('Created (UTC)'), header.indexOf('Modified (UTC)')],
    header.length
  );
  await writeSheet(sheets, spreadsheetId, CONFIG.INV_SUMMARY_SHEET, header, merged);
}

function makeInventoryDetailsKey(character, file) {
  const c = String(character || '').trim();
  const f = String(file || '').trim();
  if (!c && !f) return '';
  return `${c}\u0001${f}`;
}

function groupInventoryRows(rows, headerLength, characterIdx, fileIdx, createdIdx, modifiedIdx) {
  const map = new Map();
  rows.forEach((raw) => {
    const row = normalizeRow(raw, headerLength);
    const key = makeInventoryDetailsKey(row[characterIdx], row[fileIdx]);
    if (!key) return;
    const ts = getBestTimestamp([row[createdIdx], row[modifiedIdx]]);
    const entry = map.get(key);
    if (!entry || compareDateValues(ts.value, entry.ts.value) > 0) {
      map.set(key, { rows: [row], ts });
    }
  });
  return map;
}

async function upsertInventoryDetails(sheets, spreadsheetId, rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const header = ['Character', 'Inventory File', 'Created (UTC)', 'Modified (UTC)', 'Location', 'Name', 'ID', 'Count', 'Slots'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.INV_ITEMS_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const characterIdx = header.indexOf('Character');
  const fileIdx = header.indexOf('Inventory File');
  const createdIdx = header.indexOf('Created (UTC)');
  const modifiedIdx = header.indexOf('Modified (UTC)');

  const groups = groupInventoryRows(existingRows, header.length, characterIdx, fileIdx, createdIdx, modifiedIdx);

  rows.forEach((o) => {
    const key = makeInventoryDetailsKey(o.character, o.file);
    if (!key) return;
    const items = Array.isArray(o.items) ? o.items : [];
    const newRows = items.map((it) =>
      normalizeRow(
        [o.character, o.file, o.created, o.modified, it.Location, it.Name, it.ID, it.Count, it.Slots],
        header.length
      )
    );
    const ts = getBestTimestamp([o.created, o.modified]);
    const existing = groups.get(key);
    if (!existing || compareDateValues(ts.value, existing.ts.value) > 0) {
      groups.set(key, { rows: newRows, ts });
    }
  });

  const finalRows = [];
  groups.forEach((entry) => {
    (entry.rows || []).forEach((row) => finalRows.push(row));
  });
  await writeSheet(sheets, spreadsheetId, CONFIG.INV_ITEMS_SHEET, header, finalRows);
}

async function replaceInventoryDetails(sheets, spreadsheetId, rows) {
  const header = ['Character', 'Inventory File', 'Created (UTC)', 'Modified (UTC)', 'Location', 'Name', 'ID', 'Count', 'Slots'];
  const { header: existingHeader, rows: existingRowsRaw } = await getSheetWithHeader(sheets, spreadsheetId, CONFIG.INV_ITEMS_SHEET);
  const existingRows = existingHeader.length
    ? remapRowsToHeader(existingRowsRaw, existingHeader, header)
    : [];
  const characterIdx = header.indexOf('Character');
  const fileIdx = header.indexOf('Inventory File');
  const createdIdx = header.indexOf('Created (UTC)');
  const modifiedIdx = header.indexOf('Modified (UTC)');

  const existingGroups = groupInventoryRows(existingRows, header.length, characterIdx, fileIdx, createdIdx, modifiedIdx);
  const result = new Map();

  (Array.isArray(rows) ? rows : []).forEach((o) => {
    const key = makeInventoryDetailsKey(o.character, o.file);
    if (!key) return;
    const items = Array.isArray(o.items) ? o.items : [];
    const newRows = items.map((it) =>
      normalizeRow(
        [o.character, o.file, o.created, o.modified, it.Location, it.Name, it.ID, it.Count, it.Slots],
        header.length
      )
    );
    const ts = getBestTimestamp([o.created, o.modified]);
    const existing = existingGroups.get(key);
    let candidate = { rows: newRows, ts };
    if (existing && compareDateValues(ts.value, existing.ts.value) <= 0) {
      candidate = existing;
    }
    const current = result.get(key);
    if (!current || compareDateValues(candidate.ts.value, current.ts.value) > 0) {
      result.set(key, candidate);
    }
  });

  const finalRows = [];
  result.forEach((entry) => {
    (entry.rows || []).forEach((row) => finalRows.push(row));
  });
  await writeSheet(sheets, spreadsheetId, CONFIG.INV_ITEMS_SHEET, header, finalRows);
}

async function replaceFactionsCsv(sheets, spreadsheetId, csvText) {
  const sheetName = CONFIG.FACTION_SHEET;
  await ensureSheetExists(sheets, spreadsheetId, sheetName);
  const rows = sanitizeRows(parse(csvText || '', { relaxColumnCount: true }));
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: sheetRangeAll(sheetName)
  });
  if (rows.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetRange(sheetName)}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: rows }
    });
  }
  return { rows: rows.length, cols: rows[0] ? rows[0].length : 0 };
}

async function listCharacters(sheets, spreadsheetId) {
  const tabs = [CONFIG.ZONES_SHEET, CONFIG.FACTION_SHEET, CONFIG.INV_SUMMARY_SHEET];
  const set = new Set();
  for (const name of tabs) {
    await ensureSheetExists(sheets, spreadsheetId, name);
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: sheetRange(name)
    });
    const values = res.data.values || [];
    if (values.length < 2) continue;
    const header = values[0];
    const idx = header.indexOf('Character');
    if (idx < 0) continue;
    for (let i = 1; i < values.length; i += 1) {
      const n = String(values[i][idx] || '').trim();
      if (n) set.add(n);
    }
  }
  return Array.from(set).sort();
}

async function pushInventorySheet(sheets, spreadsheetId, body) {
  const character = String(body.character || '').trim();
  if (!character) {
    return { ok: false, error: 'Missing character' };
  }
  const baseName = String(body.sheetName || '').trim() || `Inventory - ${character}`;
  const info = body.info || {};
  const items = Array.isArray(body.items) ? body.items : [];

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const titles = new Set(
    (spreadsheet.data.sheets || []).map((sheet) => (sheet.properties && sheet.properties.title) || '')
  );
  let name = baseName;
  let n = 1;
  while (titles.has(name)) {
    n += 1;
    name = `${baseName} (${n})`;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: { title: name }
          }
        }
      ]
    }
  });
  sheetExistsCache.add(`${spreadsheetId}::${name}`);

  const tableHeader = ['Location', 'Name', 'ID', 'Count', 'Slots'];
  const data = [
    ['Inventory for', 'File', 'Created On', 'Modified On'],
    [character, info.file || '', info.created || '', info.modified || ''],
    [''],
    [],
    tableHeader
  ];
  items.forEach((it) => {
    data.push([it.Location || '', it.Name || '', it.ID || '', it.Count || '', it.Slots || '']);
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetRange(name)}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: sanitizeRows(data) }
  });

  return { ok: true, sheet: name };
}

async function jsonImport(sheets, spreadsheetId, body) {
  const tabs = Array.isArray(body && body.tabs) ? body.tabs : [];
  let imported = 0;
  for (const tab of tabs) {
    try {
      const name = String(tab.sheet || tab.name || '').trim();
      if (!name) continue;
      if (name === CONFIG.FACTION_SHEET) continue;
      const header = Array.isArray(tab.header) ? tab.header : [];
      const rows = Array.isArray(tab.rows) ? tab.rows : [];
        await ensureSheetExists(sheets, spreadsheetId, name);
      const mode = String(tab.mode || 'replace').toLowerCase();
        const payloadRows = header.length ? [header, ...rows] : rows;
        if (mode === 'append') {
          await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: `${sheetRange(name)}!A1`,
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            requestBody: { values: sanitizeRows(payloadRows) }
          });
        } else {
          const values = sanitizeRows(payloadRows);
          await sheets.spreadsheets.values.clear({
            spreadsheetId,
            range: sheetRangeAll(name)
          });
          if (values.length) {
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${sheetRange(name)}!A1`,
            valueInputOption: 'RAW',
            requestBody: { values }
          });
        }
      }
      imported += 1;
    } catch (err) {
      console.error('jsonImport tab error', err);
    }
  }
  return { imported };
}

const sheetHelpers = {
  getSheetsClient,
  upsertZones,
  upsertFactions,
  upsertInventorySummary,
  upsertInventoryDetails
};
async function replaceAll(sheets, spreadsheetId, body) {
  const up = body.upserts || body || {};
  if (up.zones) await replaceZones(sheets, spreadsheetId, up.zones);
  if (up.inventory) await replaceInventorySummary(sheets, spreadsheetId, up.inventory, body.meta || {});
  if (up.inventoryDetails) await replaceInventoryDetails(sheets, spreadsheetId, up.inventoryDetails);
}

const app = express();
app.use(express.json({ limit: '5mb' }));

app.post('/webhook', async (req, res) => {
  try {
    const body = req.body || {};
    const overrideId = extractSheetId(body.sheetId || body.sheetID || body.spreadsheetId || body.sheetUrl);
    const spreadsheetId = overrideId || SPREADSHEET_ID;
    if (!spreadsheetId) {
      res.status(500).json({ ok: false, error: 'GOOGLE_SHEET_ID is not configured' });
      return;
    }
    if (CONFIG.SECRET && String(body.secret || '') !== CONFIG.SECRET) {
      res.status(401).json({ ok: false, error: 'Unauthorized' });
      return;
    }
    try {
      await recordRawWebhook(spreadsheetId, body);
    } catch (err) {
      console.error('Record raw webhook error', err);
    }

    const sheets = await getSheetsClient();

    if (body.action === 'pushInventorySheet') {
      const result = await pushInventorySheet(sheets, spreadsheetId, body);
      res.status(result.ok === false ? 400 : 200).json(result);
      return;
    }
    if (body.action === 'replaceAll') {
      await replaceAll(sheets, spreadsheetId, body);
      res.json({ ok: true, mode: 'replaceAll' });
      return;
    }
    if (body.action === 'listCharacters') {
      const characters = await listCharacters(sheets, spreadsheetId);
      res.json({ ok: true, characters });
      return;
    }
    if (body.action === 'replaceFactionsCsv') {
      const csv = String(body.csv || '');
      if (!csv) {
        res.status(400).json({ ok: false, error: 'Missing csv' });
        return;
      }
      const result = await replaceFactionsCsv(sheets, spreadsheetId, csv);
      res.json({ ok: true, mode: 'replaceFactionsCsv', rows: result.rows, cols: result.cols });
      return;
    }
    if (body.action === 'replaceFactions') {
      res.status(403).json({ ok: false, error: 'Factions JSON disabled; use action=replaceFactionsCsv' });
      return;
    }
    if (body.action === 'jsonImport') {
      const result = await jsonImport(sheets, spreadsheetId, body);
      res.json({ ok: true, mode: 'jsonImport', imported: result.imported });
      return;
    }

    const immediate = !!(body.immediate || body.force || body.forceSync);
    let affectedCharacters = [];
    const upserts = body.upserts || {};
    if (Object.keys(upserts).length) {
      try {
        affectedCharacters = await storeWebhookPayload(spreadsheetId, upserts, body.meta || {}, { enqueue: !immediate });
      } catch (err) {
        console.error('Store webhook payload error', err);
      }
    }
    if (!immediate) {
      res.json({ ok: true, queued: true, characters: affectedCharacters });
      return;
    }
    if (upserts.zones) await upsertZones(sheets, spreadsheetId, upserts.zones);
    
    if (upserts.factions) await upsertFactions(sheets, spreadsheetId, upserts.factions);
    if (upserts.inventory) await upsertInventorySummary(sheets, spreadsheetId, upserts.inventory, body.meta || {});
    if (upserts.inventoryDetails) await upsertInventoryDetails(sheets, spreadsheetId, upserts.inventoryDetails);
    res.json({ ok: true });
  } catch (err) {
    console.error('Webhook error', err);
    res.status(500).json({ ok: false, error: err.message || String(err) });
  }
});

startScheduler().catch((err) => {
  console.error('Scheduler failed to start', err);
});

let workerController;
try {
  workerController = startWorker(sheetHelpers);
} catch (err) {
  console.error('Worker failed to start', err);
}
if (workerController && typeof workerController.catch === 'function') {
  workerController.catch((err) => {
    console.error('Worker failed to start', err);
  });
}

const port = Number(process.env.PORT || process.env.API_PORT || 3000);
app.listen(port, () => {
  console.log(`EQCM webhook server listening on port ${port}`);
});

module.exports = app;
const MAX_SHEET_CELL_LENGTH = 50000;
