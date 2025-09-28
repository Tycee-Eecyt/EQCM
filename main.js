
function buildCovSetFromSettings(){
  try{
    const base = (RAW_COV||"").split(/\r?\n/)
      .map(s => s.replace(/\s*\([^)]*\)\s*$/, ''))
      .filter(Boolean);
    const add = (state.settings?.covAdditions || []).map(String);
    const rem = new Set((state.settings?.covRemovals || []).map(s => normalizeMobName(String(s))));
    const merged = base.concat(add).filter(name => !rem.has(normalizeMobName(name)));
    return new Set(merged.map(normalizeMobName));
  }catch(e){ log('COV set build error', e); return new Set(); }
}


if (typeof getLogId !== 'function') {
  global.getLogId = function(filePath){
    try { return require('path').basename(String(filePath||'').trim()); }
    catch { return String(filePath||''); }
  };
}

// EQ Character Manager — v1.6.0 — Author: Tyler A
const { app, Tray, Menu, BrowserWindow, dialog, shell, nativeImage, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');

// ---- Log ID helpers (always available) ----
function getLogId(filePath){
  try { return require('path').basename(String(filePath||'').trim()); }
  catch { return String(filePath||''); }
}
function parseLogName(filePath){
  // eqlog_<Name>_<Server>.txt
  const b = getLogId(filePath);
  const m = /^eqlog_(.+)_(.+)\.txt$/i.exec(b);
  if (!m) return { name: null, server: null };
  return { name: m[1], server: m[2] };
}


// ---- Favorites CSV helper ----
function filterRowsByFavorites(rows){
  try {
    const favs = (state.settings && state.settings.favorites) || [];
    if (!favs.length) return rows || [];
    const set = new Set(favs.map(s => String(s).toLowerCase()));
    return (rows || []).filter(r => set.has(String(r && r[0] || '').toLowerCase()));
  } catch { return rows || []; }
}

const os = require('os');
const http = require('http');
const https = require('https');

let quitting = false;
app.on('before-quit', () => { quitting = true; });
app.on('window-all-closed', () => { /* keep alive in tray */ });

// ---------- data paths ----------
function getDataDir(){
  try { return path.join(app.getPath('userData'), 'data'); } catch { return path.join(os.homedir(), '.eq-character-manager', 'data'); }
}
const DATA_DIR = getDataDir();
fs.mkdirSync(DATA_DIR, { recursive: true });
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');
const STATE_FILE    = path.join(DATA_DIR, 'state.json');
const LOG_FILE      = path.join(DATA_DIR, 'eqwatcher.log');
const USER_COV_FILE = path.join(DATA_DIR, 'cov_list.txt');
const SHEETS_DIR    = path.join(DATA_DIR, 'sheets');

function log(...args){
  const line = `[${new Date().toISOString()}] ` + args.map(a => typeof a === 'string' ? a : JSON.stringify(a)).join(' ');
  try { fs.appendFileSync(LOG_FILE, line + '\n', 'utf8'); } catch {}
  console.log(line);
}

// ---------- defaults & state ----------
const DEFAULT_SETTINGS = {
  favorites: [],
  favoritesOnly: false,
  perCharSync: {},

  appsScriptUrl: "",
  appsScriptSecret: "",
  sheetUrl: "",  // spreadsheet link (not a specific tab)

  covList: [],
  acceptAllConsiders: false,
  strictUnstable: false,

  logsDir: "",
  baseDir: "",
  scanIntervalSec: 60,

  localSheetsEnabled: true,
  localSheetsDir: "",
  remoteSheetsEnabled: true
};

let state = {
  tz: Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC',
  offsets: {},             // file -> byte offset
  latestZones: {},         // char -> { zone, detectedUtcISO, detectedLocalISO, sourceFile }
  // New: track per-log (character+server) so same name on multiple servers doesn't collide
  latestZonesByFile: {},   // filePath -> { character, zone, detectedUtcISO, detectedLocalISO, sourceFile }
  covFaction: {},          // char -> { standing, standingDisplay, score, mob, detectedUtcISO, detectedLocalISO }
  inventory: {},           // char -> { filePath, fileCreated, fileModified, items[] }
  settings: {}
};

// Helper: latest zone source file for a character across all servers
function getLatestZoneSourceForChar(character){
  try{
    let best = null;
    const values = Object.values(state.latestZonesByFile || {});
    for (const v of values){
      if ((v.character||'') !== character) continue;
      if (!best || String(v.detectedUtcISO||'') > String(best.detectedUtcISO||'')) best = v;
    }
    if (best) return best.sourceFile || '';
  }catch(e){}
  return (state.latestZones[character]?.sourceFile||'') || '';
}

function ensureSettings(){
  if (!state.settings || typeof state.settings !== 'object') state.settings = {};
  state.settings = Object.assign({}, DEFAULT_SETTINGS, state.settings);
  return state.settings;
}

// ---------- state/settings I/O ----------
function loadState(){
  try { if (fs.existsSync(STATE_FILE)) state = Object.assign({}, state, JSON.parse(fs.readFileSync(STATE_FILE,'utf8'))); }
  catch(e){ log('loadState error', e.message); }
}
function saveState(){
  try { fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2), 'utf8'); }
  catch(e){ log('saveState error', e.message); }
}
function loadSettings(){
  try {
    if (fs.existsSync(SETTINGS_FILE)) state.settings = Object.assign({}, DEFAULT_SETTINGS, JSON.parse(fs.readFileSync(SETTINGS_FILE,'utf8')));
    else { state.settings = Object.assign({}, DEFAULT_SETTINGS); fs.writeFileSync(SETTINGS_FILE, JSON.stringify(state.settings, null, 2), 'utf8'); }
  } catch(e){ log('loadSettings error', e.message); }
}
function saveSettings(){
  try { ensureSettings(); fs.writeFileSync(SETTINGS_FILE, JSON.stringify(state.settings, null, 2), 'utf8'); }
  catch(e){ log('saveSettings error', e.message); }
}
loadState(); loadSettings(); ensureSettings();

// ---------- CoV list (names only; bundled) ----------
const RAW_COV = `A cerulean sky gazer
a cragwyrm
A Dalgortha
a fiery temple guardian
a fiery watcher
A Gargoyle Guardian
a glimmer drake
a gravid drake
A Hungry Cube
A Large Velium Statue
a lava dancer
A Shambling Cube
a shimmering green drake
a Velious Drake
Aaryonar
Abudan Fe\`Dhar
Adwetram Fe\`Dhar
Ahcaz
an ancient ice wurm defender
An Ancient Sky Drake
An Elder Onyx Drake
an emerald sky defender
an onyx sky drake
Arreken Skyward
Asteinnon Fe\`Dhar
Ayillish
Azureake
Belijor the Emerald Eye
Bezeb
Bouncer Boulder
Bratavar
Bufa
Cargalia
Cekenar
Chymot
Commander Leuz
Crendatha Fe\`Dhar
Dagarn the Destroyer
Dalshim Fe\`Dhar
Del Sapara
Deoryn Fe\`Dhar
Derasinal
Draazak
Dygwyn Fe\`Dhar
Dyr Fe\`Dhar
Eashen of the Sky
Elaend Fe\`Dhar
Elder Hajnix
Elder Kajind
Elder Kalur
Eldriaks Fe\`Dhar
Elyshum Fe\`Dhar
Entariz
Fardonad Fe\`Dhar
Gafala
Gangel
Glati
Glydoc Fe\`Dhar
Gozzrem
Grudash the Baker
Harla Dar
Honvar
Hytloc
Ionat
Jaelk
Jaled Dar\`s shade
Jaylorx
Jen Sapara
Jendavudd Fe\`Dhar
Jorlleag
Jualicn
Kalacs Fe\`Dhar
Kardakor
Karkona
Kelorek\`Dar
Klandicar
Komawin Fe\`Dhar
Lady Mirenilla
Lady Nevederia
Laegdric Fe\`Dhar
Lararith
Lawula
Lendiniara the Keeper
Lignark
Linbrak
Lord Feshlak
Lord Koi\`Doken
Lord Kreizenn
Lord Yelinak
Lothieder Fe\`Dhar
Makala
Mazi
Medry Fe\`Dhar
Morachii Fe\`Dhar
Mraaka
Myga
Nalelin Fe\`Dhar
Nalginor Fe\`Dhar
Neordla
Norsirx
Ocoenydd Fe\`Dhar
Oct Velic
Oglard
Onava
Onerind Fe\`Dhar
Orthor Velic
Pantrilla
Placlis
Poalgin Fe\`Dhar
Qalcnic Fe\`Dhar
Quadrix Velic
Quoza
Qynydd Fe\`Dhar
Ralgyn
Riran Fe\`Dhar
Rolandal
Salginor
Scout Charisa
Sentry Kale
Sevalak
Sontalak
Suez
Taegria Fe\`Dhar
Talgixn Fe\`Dhar
Talnifs
Talon Velic
Telkorenar
Telnaq
Tetragon Velic
The Seer
Theldek the Stinger
Tonvan Fe\`Dhar
Tranala
Tri Velic
Tsiraka
Tyddyn Fe\`Dhar
Ualkic
Uiliak
Umykith Fe\`Dhar
Vellyn Fe\`Dhar
Vitaela
Vobryn Fe\`Dhar
Von
Vulak\`Aerr
Wuoshi
Yaced
Yal
Yeldema
Yendilor the Cerulean Wing
Yvolcarn
Zaldin Fe\`Dhar
Zalerez
Zemm
Ziglark Whisperwing
Zil Sapara
Zildainez
Zlexak
Zynil
a cobalt drake`;

// ---------- normalization & matching ----------
function simpleStem(word){
  const w = String(word || '').toLowerCase();
  if (w.length <= 3) return w;
  if (w.endsWith('ies') && w.length > 4) return w.slice(0, -3) + 'y';
  if (w.endsWith('es') && w.length > 4) return w.slice(0, -2);
  if (w.endsWith('s') && !w.endsWith('ss') && !w.endsWith('us')) return w.slice(0, -1);
  return w;
}
function stripLeadingArticle(str){ return String(str || '').replace(/^(?:the|an|a)\s+/i, ''); }
function normalizeMobName(s){
  if (s == null) return '';
  let t = String(s)
    .replace(/\s*\(.*?\)\s*/g, ' ')     // drop parentheticals (maps)
    .replace(/[\`']/g, '')
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .toLowerCase();
  t = stripLeadingArticle(t);
  return t.split(/\s+/).map(simpleStem).join(' ');
}
function buildCovSet(){
  try {
    const base = String(RAW_COV||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
    let user = [];
    try { if (fs.existsSync(USER_COV_FILE)) user = String(fs.readFileSync(USER_COV_FILE,'utf8')).split(/\r?\n/).map(s=>s.trim()).filter(Boolean); } catch(e){}
    const merged = Array.from(new Set(base.concat(user)));
    const fromSettings = (state.settings.covList||[]).map(s=>String(s).trim()).filter(Boolean);
    for (const m of fromSettings) if (!merged.includes(m)) merged.push(m);
    return new Set(merged.map(normalizeMobName));
  } catch(e){ log('buildCovSet error', e.message); return new Set(); }
}
function getCovSet(){ return buildCovSet(); }

// ---------- regexes ----------
const RE_ZONE = /^\[(?<ts>[^\]]+)\]\s+You have entered (?<zone>.+?)\./i;
const RE_CON  = new RegExp(String.raw`^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+(?:regards you as an ally|looks upon you warmly|kindly considers you|judges you amiably|regards you indifferently|looks your way apprehensively|glowers at you dubiously|glares at you threateningly|scowls at you).*?$`, 'i');
const RE_INVIS_ON  = /(You vanish\.|Someone fades away\.|You gather shadows about you\.|Someone steps into the shadows and disappears\.)/i;
const RE_INVIS_OFF = /(You appear\.|Your shadows fade\.)/i;
const RE_SNEAK     = /(You are as quiet as a cat stalking it's prey|You are as quiet as a herd of stampeding elephants)/i;
const RE_ATTACK    = /^.*\]\s+You\s+(?:slash|pierce|bash|crush|kick|hit|smash|backstab|strike)\b/i;

// ---------- faction rules ----------
const STANDINGS = [
  { key: 'Ally',          test: /regards you as an ally/i,          score: 1450 },
  { key: 'Warmly',        test: /looks upon you warmly/i,           score: 875  },
  { key: 'Kindly',        test: /kindly considers you/i,            score: 575  },
  { key: 'Amiable',       test: /judges you amiably/i,              score: 250  },
  { key: 'Indifferent',   test: /regards you indifferently/i,       score: 0    },
  { key: 'Apprehensive',  test: /looks your way apprehensively/i,   score: -250 },
  { key: 'Dubious',       test: /glowers at you dubiously/i,        score: -575 },
  { key: 'Threatening',   test: /glares at you threateningly/i,     score: -875 },
  { key: 'Scowls',        test: /scowls at you/i,                   score: -1450}
];

function parseEqTimestamp(tsStr){
  // Example: Sat Mar 23 20:03:36 2024
  const m = tsStr.match(/^\w+\s+(\w+)\s+(\d+)\s+(\d{2}):(\d{2}):(\d{2})\s+(\d{4})$/);
  if (!m) {
    const now = new Date();
    return { utcISO: now.toISOString(), localISO: now.toLocaleString(), when: now };
  }
  const MONTHS = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const [_, mon, d, hh, mm, ss, yyyy] = m;
  const dt = new Date(Number(yyyy), MONTHS[mon] ?? 0, Number(d), Number(hh), Number(mm), Number(ss));
  return { utcISO: new Date(dt.getTime() - dt.getTimezoneOffset()*60000).toISOString(), localISO: dt.toLocaleString(), when: dt };
}

// ---------- CSV ----------
function writeCsv(filePath, header, rows){
  try{
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    const lines = [header.join(','), ...rows.map(r => r.map(v => {
      const s = String(v ?? '');
      return /[",\n]/.test(s) ? ('"' + s.replace(/"/g,'""') + '"') : s;
    }).join(','))];
    fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
  }catch(e){ log('writeCsv error', filePath, e.message); }
}

// ---------- Webhook helpers ----------

function postJson(url, data, maxRedirects = 3) {
  return new Promise((resolve, reject) => {
    const initialBody = JSON.stringify(data);
    const doRequest = (u, redirectsLeft, method = 'POST', body = initialBody) => {
      try {
        const h = u.startsWith('https') ? https : http;
        const headers = {};
        if (method !== 'GET') {
          headers['Content-Type'] = 'application/json';
          headers['Content-Length'] = Buffer.byteLength(body);
        }
        const req = h.request(
          u,
          { method, headers },
          (res) => {
            let out = '';
            res.on('data', (d) => (out += d.toString()));
            res.on('end', () => {
              if (
                res.statusCode >= 300 &&
                res.statusCode < 400 &&
                res.headers &&
                res.headers.location &&
                redirectsLeft > 0
              ) {
                const next = new URL(res.headers.location, u).toString();
                // For 301/302/303, switch to GET without body; for 307/308, keep method and body
                if (res.statusCode === 307 || res.statusCode === 308) {
                  return doRequest(next, redirectsLeft - 1, method, body);
                } else {
                  return doRequest(next, redirectsLeft - 1, 'GET', null);
                }
              }
              resolve({ status: res.statusCode, body: out });
            });
          }
        );
        req.on('error', reject);
        if (method !== 'GET' && body) req.write(body);
        req.end();
      } catch (e) {
        reject(e);
      }
    };
    doRequest(url, maxRedirects);
  });
}

function isLikelyAppsScriptExec(url){
  try{
    const u = new URL(String(url || '').trim());
    if (u.protocol !== 'https:') return false;

    if (u.hostname === 'script.google.com') {
      const parts = u.pathname.split('/').filter(Boolean);
      return (
        parts.length >= 4 &&
        parts[0] === 'macros' &&
        parts[1] === 's' &&
        /^[A-Za-z0-9_-]+$/.test(parts[2]) &&
        parts[3] === 'exec'
      );
    }
    if (u.hostname.endsWith('googleusercontent.com')) {
      return /\/macros\//.test(u.pathname);
    }
    return false;
  } catch { return false; }
}

function getRaidKitSummary(items){
  const list = Array.isArray(items) ? items : [];
  const count = (namePattern) => {
    const re = new RegExp(namePattern, 'i');
    let n = 0;
    for (const it of list) {
      if (re.test(it.Name || '')) n += Number(it.Count || 0) || 0;
    }
    return n;
  };
  const has = (namePattern) => {
    const re = new RegExp(namePattern, 'i');
    return list.some(it => re.test(it.Name || ''));
  };
  return {
    vialVeliumVapors: has('^Vial of Velium Vapors$') ? 'Y' : 'N',
    veliumVialCount: count('^Velium Vial$'),
    leatherfootSkullcap: has("^Leatherfoot Raider Skullcap$") ? 'Y' : 'N',
    shinyBrassIdol: has("^Shiny Brass Idol$") ? 'Y' : 'N',
    ringOfShadowsCount: count("^Ring of Shadows$"),
    reaperOfTheDead: has("^Reaper of the Dead$") ? 'Y' : 'N',
    pearlCount: count("^Pearl$"),
    peridotCount: count("^Peridot$"),
    mbClassFive: count("^Mana Battery - Class Five$"),
    mbClassFour: count("^Mana Battery - Class Four$"),
    mbClassThree: count("^Mana Battery - Class Three$"),
    mbClassTwo: count("^Mana Battery - Class Two$"),
    mbClassOne: count("^Mana Battery - Class One$"),
    larrikansMask: has("^Larrikan'?s Mask$") ? 'Y' : 'N'
  };
}
async function maybePostWebhook(){
  ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) return;
  const url = (state.settings.appsScriptUrl||'').trim(); if (!url) return;
  if (!isLikelyAppsScriptExec(url)) { log('Webhook not attempted: URL does not look like a /exec endpoint'); return; }
  const secret = (state.settings.appsScriptSecret||'').trim();

  // Prefer per-file rows so characters on multiple servers don't collide
  const zoneRows = (state.latestZonesByFile && Object.keys(state.latestZonesByFile).length)
    ? Object.values(state.latestZonesByFile).map(v => ({ character: v.character||'', zone: v.zone||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'', tz: state.tz||'', source: v.sourceFile||'' }))
    : Object.entries(state.latestZones || {}).map(([character, v]) => ({ character, zone: v.zone||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'', tz: state.tz||'', source: v.sourceFile||'' }));
  const covRows  = Object.entries(state.covFaction || {}).map(([character, v]) => ({ character, standing: v.standing||'', standingDisplay: v.standingDisplay||'', score: v.score ?? '', mob: v.mob||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'' }));
  const invRows  = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', logFile: getLatestZoneSourceForChar(character), created: v.fileCreated||'', modified: v.fileModified||'', raidKit: getRaidKitSummary(v.items||[]) }));
  const invDetails = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', created: v.fileCreated||'', modified: v.fileModified||'', items: v.items||[] }));

  const payload = { secret, upserts: { zones: zoneRows, factions: covRows, inventory: invRows, inventoryDetails: invDetails } };
  try {
    const res = await postJson(url, payload);
    const debugWebhook = String(process.env.DEBUG_WEBHOOK || '').toLowerCase();
    const debugOn = debugWebhook === '1' || debugWebhook === 'true' || debugWebhook === 'yes';
    const is2xx = (res.status >= 200 && res.status < 300);
    if (!is2xx || debugOn) {
      const bodyPreview = (res.body || '').slice(0, 180);
      log('Webhook response', res.status, bodyPreview);
    }
  }
  catch(e){ log('Webhook error', e.message); }
}

// One-shot replace-all import to Apps Script
async function sendReplaceAllWebhook(){
  ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) { log('ReplaceAll skipped: remoteSheetsEnabled=false'); return; }
  const url = (state.settings.appsScriptUrl||'').trim();
  if (!url || !isLikelyAppsScriptExec(url)) { log('ReplaceAll aborted: invalid Apps Script URL'); return; }
  const secret = (state.settings.appsScriptSecret||'').trim();

  const zoneRows = (state.latestZonesByFile && Object.keys(state.latestZonesByFile).length)
    ? Object.values(state.latestZonesByFile).map(v => ({ character: v.character||'', zone: v.zone||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'', tz: state.tz||'', source: v.sourceFile||'' }))
    : Object.entries(state.latestZones || {}).map(([character, v]) => ({ character, zone: v.zone||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'', tz: state.tz||'', source: v.sourceFile||'' }));
  const covRows  = Object.entries(state.covFaction || {}).map(([character, v]) => ({ character, standing: v.standing||'', standingDisplay: v.standingDisplay||'', score: v.score ?? '', mob: v.mob||'', utc: v.detectedUtcISO||'', local: v.detectedLocalISO||'' }));
  const invRows  = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', logFile: getLatestZoneSourceForChar(character), created: v.fileCreated||'', modified: v.fileModified||'', raidKit: getRaidKitSummary(v.items||[]) }));
  const invDetails = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', created: v.fileCreated||'', modified: v.fileModified||'', items: v.items||[] }));

  const payload = { secret, action: 'replaceAll', upserts: { zones: zoneRows, factions: covRows, inventory: invRows, inventoryDetails: invDetails } };
  try{
    const res = await postJson(url, payload);
    log('ReplaceAll response', res.status, (res.body||'').slice(0, 180));
  }catch(e){
    log('ReplaceAll error', e.message);
  }
}

// New: push a single character inventory to a new sheet tab (like the screenshot)
async function pushInventoryToNewSheet(character){
  ensureSettings();
  if (!character) return;
  const url = (state.settings.appsScriptUrl||'').trim();
  if (!url || !isLikelyAppsScriptExec(url)) { log('Push inventory aborted: invalid Apps Script URL'); return; }
  const secret = (state.settings.appsScriptSecret||'').trim();
  const inv = state.inventory[character];
  if (!inv || !Array.isArray(inv.items)) { log('Push inventory aborted: no inventory for', character); return; }
  const sheetName = `Inventory - ${character}`;
  const payload = {
    secret,
    action: 'pushInventorySheet',
    character,
    sheetName,
    info: {
      file: inv.filePath||'',
      created: inv.fileCreated||'',
      modified: inv.fileModified||'',
      logFile: getLatestZoneSourceForChar(character)
    },
    items: inv.items
  };
  try{
    const res = await postJson(url, payload);
    log('Push inventory response', res.status, (res.body||'').slice(0,200));
    const sheetUrl = (state.settings.sheetUrl||'').trim();
    if (sheetUrl) shell.openExternal(sheetUrl);
  }catch(e){
    log('Push inventory error', e.message);
  }
}

function buildPushInventorySubmenu(){
  const chars = Object.keys(state.inventory || {}).sort();
  if (!chars.length) return { label: 'Push inventory to sheet', enabled: false };
  return {
    label: 'Push inventory to sheet',
    submenu: chars.map(ch => ({ label: ch, click: () => pushInventoryToNewSheet(ch) }))
  };
}

// ---------- EQ helpers ----------
async function ensureLogsDir(){
  ensureSettings();
  const dir = state.settings.logsDir;
  if (dir && fs.existsSync(dir)) return dir;
  const pick = await dialog.showOpenDialog({ title: 'Select EverQuest Logs folder', properties: ['openDirectory'] });
  if (!pick.canceled && pick.filePaths[0]) { state.settings.logsDir = pick.filePaths[0]; saveSettings(); return state.settings.logsDir; }
  return null;
}
async function ensureBaseDir(){
  ensureSettings();
  const dir = state.settings.baseDir;
  if (dir && fs.existsSync(dir)) return dir;
  const pick = await dialog.showOpenDialog({ title: 'Select EverQuest base folder', properties: ['openDirectory'] });
  if (!pick.canceled && pick.filePaths[0]) { state.settings.baseDir = pick.filePaths[0]; saveSettings(); return state.settings.baseDir; }
  return null;
}

// ---------- inventory parsing ----------
function parseInventoryFile(filePath){
  try{
    const txt = fs.readFileSync(filePath, 'utf8').replace(/\r\n/g, '\n');
    const lines = txt.split('\n').filter(l => l.length>0);
    lines.shift(); // header
    const items = lines.map(l => {
      const parts = l.split(/\t/);
      return { Location: parts[0], Name: parts[1], ID: parts[2], Count: Number(parts[3]||0), Slots: Number(parts[4]||0) };
    });
    const st = fs.statSync(filePath);
    return {
      filePath,
      fileCreated: new Date(st.birthtimeMs || st.ctimeMs).toISOString(),
      fileModified: new Date(st.mtimeMs).toISOString(),
      items
    };
  }catch(e){ log('parseInventory error', filePath, e.message); return null; }
}

// ---------- scanning ----------
let scanTimer = null;

async function scanLogs(){
  const dir = state.settings.logsDir;
  if (!dir || !fs.existsSync(dir)) return;
  if (!state.latestZonesByFile || typeof state.latestZonesByFile !== 'object') state.latestZonesByFile = {};
  const files = fs.readdirSync(dir).filter(f => /^eqlog_.+?\.txt$/i.test(f));
  for (const f of files){
    const full = path.join(dir, f);
    const parsed = parseLogName(f);
    const char = (parsed && parsed.name) ? parsed.name : f.replace(/^eqlog_([^_]+).*$/i,'$1');
    try {
      const last = state.offsets[full] || 0;
      const stat = fs.statSync(full);
      let start = last && last < stat.size ? last : Math.max(0, stat.size - 256*1024);
      const fd = fs.openSync(full, 'r');
      const len = Math.max(0, stat.size - start);
      const buf = Buffer.alloc(len);
      fs.readSync(fd, buf, 0, len, start);
      fs.closeSync(fd);
      const text = buf.toString('utf8');
      state.offsets[full] = stat.size;

      const lines = text.replace(/\r\n/g,'\n').split('\n').filter(Boolean);
      const covLast = state._covLastStable || (state._covLastStable = {});
      if (!covLast[char]) covLast[char] = {};

      let sawZone = false;
      for (let i=0;i<lines.length;i++){
        const line = lines[i];

        // Zone
        const mZ = line.match(RE_ZONE);
        if (mZ){
          const { ts, zone } = mZ.groups;
          const t = parseEqTimestamp(ts);
          // Maintain legacy per-character mapping (last one wins)
          state.latestZones[char] = { zone, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO, sourceFile: full };
          // New: also record per-file (character+server) to disambiguate same-name chars on multiple servers
          state.latestZonesByFile[full] = { character: char, zone, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO, sourceFile: full };
          sawZone = true;
          continue;
        }

        // Consider
        const mC = line.match(RE_CON);
        if (mC){
          const { ts, mob } = mC.groups;
          const t = parseEqTimestamp(ts);

          // standing
          const hit = STANDINGS.find(s => s.test.test(line));
          const standing = hit ? hit.key : 'Indifferent';
          const score = hit ? hit.score : 0;

          // unstable window (look-behind)
          const lookBehind = state.settings.strictUnstable ? 10 : 3;
          const startIdx = Math.max(0, i - lookBehind);
          const window = lines.slice(startIdx, i+1);
          const unstableInvis  = window.some(l => RE_INVIS_ON.test(l) || RE_INVIS_OFF.test(l) || RE_SNEAK.test(l));
          const unstableCombat = window.some(l => RE_ATTACK.test(l));

          // CoV membership
          const mobNorm = normalizeMobName(mob);
          const COV_SET = getCovSet();
          let isCov = false;
          if (state.settings.acceptAllConsiders) isCov = true;
          else if (COV_SET.has(mobNorm)) isCov = true;
          else { for (const name of COV_SET){ if (mobNorm.startsWith(name)) { isCov = true; break; } } }

          if (isCov){
            if (unstableInvis || unstableCombat){
              const lastMob = covLast[char][mobNorm];
              const lastChar = state.covFaction[char];
              if (lastMob){
                state.covFaction[char] = { standing: lastMob.standing, standingDisplay: lastMob.standing + ' (fallback)', score: lastMob.score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Unstable /con, applied mob fallback', char, mob);
              } else if (lastChar){
                state.covFaction[char] = Object.assign({}, lastChar, { standingDisplay: (lastChar.standingDisplay||lastChar.standing||'') + ' (fallback)', mob: lastChar.mob || mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO });
                log('Unstable /con, applied char fallback', char, mob);
              } else {
                state.covFaction[char] = { standing, standingDisplay: standing + ' (uncertain)', score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Unstable /con, accepted as baseline', char, mob);
              }
            } else {
              state.covFaction[char] = { standing, standingDisplay: standing, score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
              covLast[char][mobNorm] = state.covFaction[char];
            }
          }
        }
      }
      // If we didn't see any zone line for this file and have no prior record, add a placeholder entry
      if (!sawZone && !state.latestZonesByFile[full]){
        state.latestZonesByFile[full] = { character: char, zone: '', detectedUtcISO: '', detectedLocalISO: '', sourceFile: full };
      }
    } catch(e){
      log('Scan file error', full, e.message);
    }
  }
}

async function scanInventory(){
  const baseDir = state.settings.baseDir;
  if (!baseDir || !fs.existsSync(baseDir)) return;
  const files = fs.readdirSync(baseDir).filter(f => /-Inventory\.txt$/i.test(f));
  for (const f of files){
    const full = path.join(baseDir, f);
    const char = f.replace(/-Inventory\.txt$/i,'').trim();
    const parsed = parseInventoryFile(full);
    if (parsed) state.inventory[char] = parsed;
  }
}

async function doScanCycle(){
  try{
    await scanLogs();
    await scanInventory();
    saveState();
    await maybeWriteLocalSheets();
    await maybePostWebhook();
  } catch(e){
    log('[LOG] Scan cycle error:', e.message);
  }
}

function startScanning(){
  ensureSettings();
  if (scanTimer) return;
  doScanCycle();
  const ms = Math.max(5000, 1000 * (state.settings.scanIntervalSec || 60));
  scanTimer = setInterval(doScanCycle, ms);
}
function stopScanning(){ if (scanTimer){ clearInterval(scanTimer); scanTimer=null; } }
function restartScanning(){ stopScanning(); startScanning(); }

// ---------- local "sheets" (CSV) ----------
async function maybeWriteLocalSheets(){
  ensureSettings();
  if (state.settings.localSheetsEnabled === false) return;
  const dir = state.settings.localSheetsDir && state.settings.localSheetsDir.trim() ? state.settings.localSheetsDir.trim() : SHEETS_DIR;

  const zHead = ['Character','Last Zone','Zone Time (UTC)','Zone Time (Local)','Device TZ','Source Log File'];
  const zRows = (state.latestZonesByFile && Object.keys(state.latestZonesByFile).length)
    ? Object.values(state.latestZonesByFile).map(v => [v.character||'', v.zone||'', v.detectedUtcISO||'', v.detectedLocalISO||'', state.tz||'', v.sourceFile||''])
    : Object.entries(state.latestZones || {}).map(([char, v]) => [char, v.zone||'', v.detectedUtcISO||'', v.detectedLocalISO||'', state.tz||'', v.sourceFile||'']);
const zRowsOut = filterRowsByFavorites(zRows);
  writeCsv(path.join(dir, 'Zone Tracker.csv'), zHead, zRowsOut);

  const fHead = ['Character','Standing','Score','Mob','Consider Time (UTC)','Consider Time (Local)','Notes'];
  const fRows = Object.entries(state.covFaction || {}).map(([char, v]) => [char, v.standing||'', v.score ?? '', v.mob||'', v.detectedUtcISO||'', v.detectedLocalISO||'', (v.standingDisplay||'').includes('fallback')? 'fallback' : (v.standingDisplay||'').includes('uncertain')? 'uncertain' : '' ]);
const fRowsOut = filterRowsByFavorites(fRows);
  writeCsv(path.join(dir, 'CoV Faction.csv'), fHead, fRowsOut);

  const iHead = ['Character','Log ID','Inventory File','Source Log File','Created (UTC)','Modified (UTC)',
                 'Vial of Velium Vapors','Velium Vial Count','Leatherfoot Raider Skullcap','Shiny Brass Idol',
                 'Ring of Shadows Count','Reaper of the Dead','Pearl Count','Peridot Count','Larrikan\'s Mask',
                 'MB Class Five','MB Class Four','MB Class Three','MB Class Two','MB Class One',
                 'Spreadsheet URL','Suggested Sheet Name'];
const iRows = Object.entries(state.inventory || {}).map(([char, v]) => {
  const kit = getRaidKitSummary(v.items||[]);
  const suggested = `Inventory - ${char}`;
  return [char, getLogId(v.filePath||''), v.filePath||'', getLatestZoneSourceForChar(char), v.fileCreated||'', v.fileModified||'',
          kit.vialVeliumVapors, kit.veliumVialCount, kit.leatherfootSkullcap, kit.shinyBrassIdol,
          kit.ringOfShadowsCount, kit.reaperOfTheDead, kit.pearlCount, kit.peridotCount, kit.larrikansMask,
          kit.mbClassFive, kit.mbClassFour, kit.mbClassThree, kit.mbClassTwo, kit.mbClassOne,
          (state.settings.sheetUrl||''), suggested];
});
const iRowsOut = filterRowsByFavorites(iRows);
writeCsv(path.join(dir, 'Inventory Summary.csv'), iHead, iRowsOut);

  // per-character CSV
  for (const [char, inv] of Object.entries(state.inventory || {})){
    const rows = (inv.items||[]).map(it => [char, inv.filePath||'', inv.fileCreated||'', inv.fileModified||'', it.Location||'', it.Name||'', it.ID||'', it.Count||0, it.Slots||0]);
    const header = ['Character','Inventory File','Created (UTC)','Modified (UTC)','Location','Name','ID','Count','Slots'];
    writeCsv(path.join(dir, `Inventory Items - ${char}.csv`), header, rows);
  }
}

// ---------- UI ----------
let tray = null, settingsWin = null;
function buildTrayTooltip(){
  ensureSettings();
  const tz = state.tz || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  const interval = state.settings.scanIntervalSec || 60;
  const logs = state.settings.logsDir || '(pick logs folder)';
  const base = state.settings.baseDir || '(pick base folder)';
  const status = scanTimer ? 'running' : 'paused';
  return `EQ Character Manager\nStatus: ${status}\nTZ: ${tz} — Scan: ${interval}s\nLogs: ${logs}\nBase: ${base}`;
}
function buildMenu(){
  ensureSettings();
  const scanChoices = [30,60,120,300];
  const scanSub = { label: 'Scan interval', submenu: scanChoices.map(sec => ({ label: `${sec}s${(state.settings.scanIntervalSec===sec) ? ' ✓':''}`, click: () => { state.settings.scanIntervalSec = sec; saveSettings(); restartScanning(); rebuildTray(); } })) };
  return Menu.buildFromTemplate([
    buildPushInventorySubmenu(),
    { label: 'Settings…', click: openSettingsWindow },
    { type: 'separator' },
    { label: 'Docs (Sheets deploy)', click: openDocsWindow },
    { label: 'Full refresh to sheet (replace all)', click: async () => {
        try{
          const url = (state.settings.appsScriptUrl||'').trim();
          if (!url || !isLikelyAppsScriptExec(url)) { dialog.showMessageBox({ type: 'warning', message: 'Apps Script URL is not set or invalid.' }); return; }
          const res = await dialog.showMessageBox({
            type: 'question', buttons: ['Cancel','Proceed'], defaultId: 1, cancelId: 0,
            message: 'Replace all rows in Google Sheet?',
            detail: 'This will clear and rewrite Zone Tracker, CoV Faction, Inventory Summary, and Inventory Items using current data.'
          });
          if (res.response === 1) await sendReplaceAllWebhook();
        }catch(e){ log('ReplaceAll menu error', e.message); }
      }
    },
    { label: 'Rescan now', click: () => { doScanCycle(); } },
    { label: 'Open data folder', click: () => { shell.openPath(DATA_DIR); } },
    { type: 'separator' },
    { label: scanTimer ? 'Pause scanning' : 'Start scanning', click: () => { scanTimer ? stopScanning() : startScanning(); rebuildTray(); } },
    scanSub,
    { type: 'separator' },
    { label: 'Quit', click: () => { quitting = true; app.quit(); } }
  ]);
}
function rebuildTray(){ if (!tray) return; tray.setContextMenu(buildMenu()); tray.setToolTip(buildTrayTooltip()); }
function makeHidable(win){
  win.on('close', (e) => { if (!quitting) { e.preventDefault(); win.hide(); } });
}
function openSettingsWindow(){
  if (settingsWin && !settingsWin.isDestroyed()) { settingsWin.show(); settingsWin.focus(); return; }
  settingsWin = new BrowserWindow({ width: 900, height: 700, resizable: true, webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  settingsWin.setMenu(null);
  settingsWin.loadFile(path.join(__dirname, 'settings.html'));
  makeHidable(settingsWin);
  settingsWin.on('closed', () => settingsWin = null);
}
function openDocsWindow(){
  const win = new BrowserWindow({ width: 900, height: 740, resizable: true });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'docs', 'deploy-sheets.html'));
}

// ---------- IPC ----------
ipcMain.handle('settings:get', async () => {
  ensureSettings();
  const characters = new Set([
    ...Object.keys(state.latestZones || {}),
    ...Object.keys(state.covFaction || {}),
    ...Object.keys(state.inventory || {}),
    ...(state.settings.favorites || [])
  ]);
  return { settings: state.settings, characters: Array.from(characters).sort(), env: {} };
});

ipcMain.handle('settings:deriveSheetId', async (evt, url) => {
  try {
    const input = String(url || '').trim();
    const idOnly = /^[A-Za-z0-9_-]{20,}$/;
    if (idOnly.test(input)) return { sheetId: input };
    const u = new URL(/^https?:\/\//i.test(input) ? input : ('https://' + input));
    const parts = u.pathname.split('/').filter(Boolean);
    const idx = parts.indexOf('d');
    const id = (idx >= 0 && parts[idx+1]) ? parts[idx+1] : '';
    return { sheetId: id };
  } catch { return { sheetId: '' }; }
});

ipcMain.handle('settings:set', async (evt, payload) => {
  ensureSettings();
  state.settings = Object.assign({}, state.settings, payload || {});
  saveSettings(); rebuildTray();
  return { ok: true };
});
ipcMain.handle('settings:browseFolder', async (evt, which) => {
  let title = 'Select Folder';
  if (which === 'logsDir') title = 'Select Log Folder';
  else if (which === 'baseDir') title = 'Select Base EverQuest Folder';
  const pick = await dialog.showOpenDialog({ title, properties: ['openDirectory'] });
  if (pick.canceled || !pick.filePaths[0]) return { path: '' };
  const p = pick.filePaths[0];
  if (which === 'logsDir') state.settings.logsDir = p;
  if (which === 'baseDir') state.settings.baseDir = p;
  if (which === 'localSheetsDir') state.settings.localSheetsDir = p;
  saveSettings(); rebuildTray();
  return { path: p };
});
app.whenReady().then(() => {
  const iconPath = path.join(__dirname, 'assets', 'tray.png');
  let img = nativeImage.createFromPath(iconPath);
  if (img.isEmpty()) img = nativeImage.createEmpty();
  tray = new Tray(img);
  rebuildTray();
  startScanning();
});

process.on('unhandledRejection', (err) => { log('Unhandled rejection', err && err.stack ? err.stack : String(err)); });
process.on('uncaughtException', (err) => { log('Uncaught exception', err && err.stack ? err.stack : String(err)); });
