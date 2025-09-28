
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
  remoteSheetsEnabled: true,

  // Backscan configuration
  backscanMaxMB: 0,             // 0 = full file; otherwise clamp 5–20 MB
  backscanRetryMinutes: 10,     // 0 disables periodic retry after first attempt

  // Consideration heuristics
  invisMaxMinutes: 20,          // treat self-invis as potentially active up to this long
  combatRecentMinutes: 5        // treat combat as "recent" for this many minutes
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

function clamp(n, lo, hi){ n = Number(n||0); if (Number.isNaN(n)) n=0; return Math.max(lo, Math.min(hi, n)); }
function getBackscanBytes(){
  ensureSettings();
  const raw = Number(state.settings.backscanMaxMB || 0);
  if (!Number.isFinite(raw) || raw <= 0) return 0; // 0 means unlimited (full file)
  const mb = clamp(raw, 5, 20);
  return mb * 1024 * 1024;
}
function getBackscanRetryMs(){
  ensureSettings();
  const mins = Math.max(0, Number(state.settings.backscanRetryMinutes||0));
  return mins * 60 * 1000;
}

function getInvisMaxMs(){ ensureSettings(); return Math.max(1, Number(state.settings.invisMaxMinutes||20)) * 60 * 1000; }
function getCombatRecentMs(){ ensureSettings(); return Math.max(1, Number(state.settings.combatRecentMinutes||5)) * 60 * 1000; }

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
a cobalt drake
a wyvern`;

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
    const fromSettings = (state.settings.covList||[]).map(s=>String(s).trim()).filter(Boolean);
    const add = (state.settings.covAdditions||[]).map(s=>String(s).trim()).filter(Boolean);
    const rem = new Set((state.settings.covRemovals||[]).map(s => normalizeMobName(String(s))));
    // merge base + user + explicit list + additions
    const mergedRaw = Array.from(new Set([ ...base, ...user, ...fromSettings, ...add ]));
    // normalize and remove removals
    const out = new Set();
    for (const name of mergedRaw){
      const norm = normalizeMobName(name);
      if (!norm) continue;
      if (rem.has(norm)) continue;
      out.add(norm);
    }
    return out;
  } catch(e){ log('buildCovSet error', e.message); return new Set(); }
}
function getCovSet(){ return buildCovSet(); }

// ---------- regexes ----------
const RE_ZONE = /^\[(?<ts>[^\]]+)\]\s+You have entered (?<zone>.+?)\./i;
const RE_CON  = new RegExp(String.raw`^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+(?:regards you as an ally|looks upon you warmly|kindly considers you|judges you amiably|regards you indifferently|looks your way apprehensively|glowers at you dubiously|glares at you threateningly|scowls at you).*?$`, 'i');
const RE_INVIS_ON  = /(You vanish\.|Someone fades away\.|You gather shadows about you\.|Someone steps into the shadows and disappears\.)/i;
const RE_INVIS_OFF = /(You appear\.|Your shadows fade\.)/i;
const RE_SELF_INVIS_ON  = /(You vanish\.|You gather shadows about you\.)/i;
const RE_SELF_INVIS_OFF = /(You appear\.|Your shadows fade\.)/i;
const RE_SNEAK     = /(You are as quiet as a cat stalking it's prey|You are as quiet as a herd of stampeding elephants)/i;
const RE_ATTACK    = /^.*\]\s+You\s+(?:slash|pierce|bash|crush|kick|hit|smash|backstab|strike)\b/i;
const RE_ATTACK_NAMED_HIT  = /^\[(?<ts>[^\]]+)\]\s+You\s+(?:slash|pierce|bash|crush|kick|hit|smash|backstab|strike)\s+(?<mob>.+?)\s+for\s+\d+\s+points of damage\./i;
const RE_ATTACK_NAMED_MISS = /^\[(?<ts>[^\]]+)\]\s+You\s+try to\s+(?:slash|pierce|punch|bash|crush|kick|hit|smash|backstab|strike)\s+(?<mob>.+?),\s+but\s+miss!$/i;
const RE_AUTO_ATTACK_ON = /^\[(?<ts>[^\]]+)\]\s+Auto attack on\./i;
const RE_MOB_HIT_YOU    = /^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+(?:hits|kicks|bashes)\s+YOU\b/i;
const RE_MOB_TRY_HIT_YOU= /^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+tries to\s+(?:hit|bash)\s+YOU\b/i;
const RE_MOB_NON_MELEE  = /^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+was hit by non-melee\b/i;
const RE_MOB_THORNS     = /^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+was pierced by thorns\b/i;
// Spell damage (player to mob)
const RE_SPELL_YOUR_HITS = /^\[(?<ts>[^\]]+)\]\s+Your\s+.+?\s+hits\s+(?<mob>.+?)\s+for\s+\d+\s+points? of (?:\w+\s+)?damage\./i;
const RE_SPELL_YOU_HIT   = /^\[(?<ts>[^\]]+)\]\s+You\s+(?:blast|smite|burn|shock|freeze|immolate|incinerate|strike|hit)\s+(?<mob>.+?)\s+for\s+\d+\s+points? of (?:\w+\s+)?damage\./i;
const RE_SPELL_DOT_TICK  = /^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+has taken\s+\d+\s+damage from your\s+.+?\./i;

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

// On first sight of a log file, backscan from the end to find
// the most recent zone line so we can seed Zone Tracker without
// waiting for the next zone event.
function findLastZoneInFile(filePath, maxBytes = 0){
  try{
    const stat = fs.statSync(filePath);
    const size = stat.size || 0;
    if (size <= 0) return null;
    const fd = fs.openSync(filePath, 'r');
    const block = 512 * 1024; // 512KB chunks
    let endPos = size;
    let readTotal = 0;
    const limit = (!Number.isFinite(maxBytes) || maxBytes <= 0) ? size : Math.min(maxBytes, size);
    let carry = '';
    while (endPos > 0 && readTotal < limit){
      const toRead = Math.min(block, endPos, limit - readTotal);
      const startPos = endPos - toRead;
      const buf = Buffer.alloc(toRead);
      fs.readSync(fd, buf, 0, toRead, startPos);
      let current = buf.toString('utf8', 0, toRead) + carry;

      // Scan lines from end to start for the first zone match
      let idxEnd = current.length;
      while (idxEnd > 0){
        const idxNL = current.lastIndexOf('\n', idxEnd - 1);
        const line = current.substring(idxNL + 1, idxEnd).replace(/\r$/, '');
        const mZ = line.match(RE_ZONE);
        if (mZ){
          const { ts, zone } = mZ.groups;
          const t = parseEqTimestamp(ts);
          fs.closeSync(fd);
          return { zone, utcISO: t.utcISO, localISO: t.localISO };
        }
        if (idxNL < 0) break;
        idxEnd = idxNL;
      }

      // Preserve any incomplete first line at the beginning for the next earlier chunk
      const firstNL = current.indexOf('\n');
      carry = firstNL >= 0 ? current.substring(0, firstNL) : current;

      endPos = startPos;
      readTotal += toRead;
    }
    fs.closeSync(fd);
    return null;
  }catch(e){ log('findLastZoneInFile error', filePath, e.message); return null; }
}

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
      const hadOffset = Object.prototype.hasOwnProperty.call(state.offsets || {}, full);
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
      const status = state._status || (state._status = {});
      const nowState = status[char] || (status[char] = { invisOn: 0, invisOff: 0, lastCombat: 0, attacks: {}, prevBeforeInvis: null, prevBeforeCombat: null });
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

        // Track self invis on/off with timestamps
        if (RE_SELF_INVIS_ON.test(line)){
          const tsStr = (line.match(/^\[([^\]]+)\]/)||[])[1] || '';
          const t = parseEqTimestamp(tsStr);
          // snapshot last known faction before invis begins
          if (state.covFaction[char]) nowState.prevBeforeInvis = Object.assign({}, state.covFaction[char]);
          nowState.invisOn = t.when.getTime();
          continue;
        }
        if (RE_SELF_INVIS_OFF.test(line)){
          const tsStr = (line.match(/^\[([^\]]+)\]/)||[])[1] || '';
          const t = parseEqTimestamp(tsStr);
          nowState.invisOff = t.when.getTime();
          // clear snapshot after invis ends
          // (keep last snapshot until next invis if you prefer historical context)
          // nowState.prevBeforeInvis = null;
          continue;
        }

        // Track combat: named player attacks, auto-attack on (snapshot), mob hits you, thorns/non-melee damage
        let mAH = line.match(RE_ATTACK_NAMED_HIT);
        let mAM = mAH ? null : line.match(RE_ATTACK_NAMED_MISS);
        const mAA = mAH || mAM ? null : line.match(RE_AUTO_ATTACK_ON);
        const mHY = (!mAH && !mAM && !mAA) ? line.match(RE_MOB_HIT_YOU) : null;
        const mTY = (!mAH && !mAM && !mAA && !mHY) ? line.match(RE_MOB_TRY_HIT_YOU) : null;
        const mNM = (!mAH && !mAM && !mAA && !mHY && !mTY) ? line.match(RE_MOB_NON_MELEE) : null;
        const mTH = (!mAH && !mAM && !mAA && !mHY && !mTY && !mNM) ? line.match(RE_MOB_THORNS) : null;
        const mSH = (!mAH && !mAM && !mAA && !mHY && !mTY && !mNM && !mTH) ? line.match(RE_SPELL_YOUR_HITS) : null;
        const mYH = (!mAH && !mAM && !mAA && !mHY && !mTY && !mNM && !mTH && !mSH) ? line.match(RE_SPELL_YOU_HIT) : null;
        const mDT = (!mAH && !mAM && !mAA && !mHY && !mTY && !mNM && !mTH && !mSH && !mYH) ? line.match(RE_SPELL_DOT_TICK) : null;
        if (mAH || mAM){
          const ts = (mAH||mAM).groups.ts;
          const mob = (mAH||mAM).groups.mob;
          const t = parseEqTimestamp(ts);
          // snapshot last known faction before combat engagement
          if (state.covFaction[char]) nowState.prevBeforeCombat = Object.assign({}, state.covFaction[char]);
          nowState.lastCombat = t.when.getTime();
          const key = normalizeMobName(mob);
          if (!nowState.attacks) nowState.attacks = {};
          nowState.attacks[key] = t.when.getTime();
          continue;
        }
        if (mAA){
          const ts = mAA.groups.ts; const t = parseEqTimestamp(ts);
          if (state.covFaction[char]) nowState.prevBeforeCombat = Object.assign({}, state.covFaction[char]);
          nowState.lastCombat = t.when.getTime();
          continue;
        }
        if (mHY || mTY || mNM || mTH || mSH || mYH || mDT){
          const g = (mHY||mTY||mNM||mTH||mSH||mYH||mDT).groups;
          const t = parseEqTimestamp(g.ts);
          nowState.lastCombat = t.when.getTime();
          const key = normalizeMobName(g.mob);
          if (!nowState.attacks) nowState.attacks = {};
          nowState.attacks[key] = t.when.getTime();
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

          // unstable window (time-based): invis can be active for minutes; combat can bias con
          const invisMaxMs = getInvisMaxMs();
          const combatMs = getCombatRecentMs();
          const nowMs = t.when.getTime();
          const invisActive = (nowState.invisOn && (!nowState.invisOff || nowState.invisOn > nowState.invisOff) && (nowMs - nowState.invisOn <= invisMaxMs));
          const attackedThisMobRecently = (() => {
            const key = normalizeMobName(mob);
            const last = (nowState.attacks && nowState.attacks[key]) || 0;
            return last && (nowMs - last <= combatMs);
          })();
          // keep line-based quick heuristics as secondary signals
          const lookBehind = state.settings.strictUnstable ? 10 : 3;
          const startIdx = Math.max(0, i - lookBehind);
          const window = lines.slice(startIdx, i+1);
          const unstableInvis  = invisActive || window.some(l => RE_INVIS_ON.test(l) || RE_INVIS_OFF.test(l) || RE_SNEAK.test(l));
          const unstableCombat = attackedThisMobRecently || window.some(l => RE_ATTACK.test(l));

          // CoV membership
          const mobNorm = normalizeMobName(mob);
          const COV_SET = getCovSet();
          let isCov = false;
          if (state.settings.acceptAllConsiders) isCov = true;
          else if (COV_SET.has(mobNorm)) isCov = true;
          else { for (const name of COV_SET){ if (mobNorm.startsWith(name)) { isCov = true; break; } } }

          if (isCov){
            // Special rules:
            // - Invis bias: if invis is active and standing came out 'Indifferent', avoid incorrectly locking in Indifferent.
            // - Combat bias: if we attacked this mob very recently, a hostile con may be biased; prefer fallback.
            const lastMob = covLast[char][mobNorm];
            const lastChar = state.covFaction[char];
            const preferFallbackForInvis = (invisActive && standing === 'Indifferent');
            const preferFallbackForCombat = attackedThisMobRecently && (standing === 'Threatening' || standing === 'Dubious' || standing === 'Apprehensive');

            if (preferFallbackForCombat || unstableCombat){
              if (lastMob){
                const prev = nowState.prevBeforeCombat;
                const note = prev && prev.standing ? ` (combat; prev=${prev.standing})` : ' (combat)';
                state.covFaction[char] = { standing: lastMob.standing, standingDisplay: lastMob.standing + note, score: lastMob.score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Combat-biased /con, applied mob fallback', char, mob);
              } else if (lastChar){
                const prev = nowState.prevBeforeCombat;
                const note = prev && prev.standing ? ' (combat; prev=' + prev.standing + ')' : ' (combat)';
                state.covFaction[char] = Object.assign({}, lastChar, { standingDisplay: (lastChar.standingDisplay||lastChar.standing||'') + note, mob: lastChar.mob || mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO });
                log('Combat-biased /con, applied char fallback', char, mob);
              } else {
                const prev = nowState.prevBeforeCombat;
                const note = prev && prev.standing ? ' (combat?; prev=' + prev.standing + ')' : ' (combat?)';
                state.covFaction[char] = { standing, standingDisplay: standing + note, score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Combat-biased /con, accepted as baseline', char, mob);
              }
            }
            else if (preferFallbackForInvis || unstableInvis){
              if (lastMob){
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ` (invis; prev=${prev.standing})` : ' (invis)';
                state.covFaction[char] = { standing: lastMob.standing, standingDisplay: lastMob.standing + note, score: lastMob.score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Invis-biased /con, applied mob fallback', char, mob);
              } else if (lastChar){
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ' (invis; prev=' + prev.standing + ')' : ' (invis)';
                state.covFaction[char] = Object.assign({}, lastChar, { standingDisplay: (lastChar.standingDisplay||lastChar.standing||'') + note, mob: lastChar.mob || mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO });
                log('Invis-biased /con, applied char fallback', char, mob);
              } else {
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ' (invis?; prev=' + prev.standing + ')' : ' (invis?)';
                state.covFaction[char] = { standing, standingDisplay: standing + note, score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                log('Invis-biased /con, accepted as baseline', char, mob);
              }
            }
            else {
              state.covFaction[char] = { standing, standingDisplay: standing, score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
              covLast[char][mobNorm] = state.covFaction[char];
            }
          }
        }
      }
      // If we didn't see any zone in the recent tail: backscan from end to find the most recent zone.
      // Try on first sight; if still no zone recorded, retry periodically per settings.
      const rec = state.latestZonesByFile[full];
      const haveZone = !!(rec && rec.zone);
      if (!sawZone && (!hadOffset || !haveZone)){
        const now = Date.now();
        if (!state._backscanNextAtByFile) state._backscanNextAtByFile = {};
        const retryMs = getBackscanRetryMs();
        const nextAt = state._backscanNextAtByFile[full] || 0;
        const shouldTry = (!hadOffset) || (now >= nextAt);
        if (shouldTry){
          const lastZ = findLastZoneInFile(full, getBackscanBytes());
          if (retryMs > 0) state._backscanNextAtByFile[full] = now + retryMs; else state._backscanNextAtByFile[full] = now + (24*60*60*1000);
          if (lastZ){
            state.latestZones[char] = { zone: lastZ.zone, detectedUtcISO: lastZ.utcISO, detectedLocalISO: lastZ.localISO, sourceFile: full };
            state.latestZonesByFile[full] = { character: char, zone: lastZ.zone, detectedUtcISO: lastZ.utcISO, detectedLocalISO: lastZ.localISO, sourceFile: full };
            sawZone = true;
            delete state._backscanNextAtByFile[full];
            log('Backscan seeded last zone', char, lastZ.zone);
          }
        }
      }

      // If we still didn't see any zone line for this file and have no prior record, add a placeholder entry
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
    { label: 'CoV Mob List…', click: openCovWindow },
    { label: 'Settings…', click: openSettingsWindow },
    { label: 'Advanced…', click: openAdvancedWindow },
    { type: 'separator' },
    { label: 'Getting Started', click: openDocsWindow },
    { label: 'Rescan now', click: () => { doScanCycle(); } },
    { label: 'Open data folder', click: () => { shell.openPath(DATA_DIR); } },
    { label: 'Open local CSV folder', click: () => {
        const outDir = (state.settings.localSheetsDir && state.settings.localSheetsDir.trim()) ? state.settings.localSheetsDir.trim() : SHEETS_DIR;
        shell.openPath(outDir);
      }
    },
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
  settingsWin = new BrowserWindow({ width: 900, height: 700, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  settingsWin.setMenu(null);
  settingsWin.loadFile(path.join(__dirname, 'settings.html'));
  makeHidable(settingsWin);
  settingsWin.on('closed', () => settingsWin = null);
}
function openDocsWindow(){
  const win = new BrowserWindow({ width: 900, height: 740, resizable: true, icon: getWindowIconImage() });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'docs', 'deploy-sheets.html'));
}
function openCovWindow(){
  const win = new BrowserWindow({ width: 900, height: 740, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'cov-list.html'));
  makeHidable(win);
}
function openAdvancedWindow(){
  const win = new BrowserWindow({ width: 800, height: 600, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'advanced.html'));
  makeHidable(win);
}

// ---------- Icon helpers (SVG rasterize with fallback) ----------
function tryCreateImageFromSvg(svgPath){
  try {
    if (!fs.existsSync(svgPath)) return null;
    const svg = fs.readFileSync(svgPath, 'utf8');
    const dataUrl = 'data:image/svg+xml;base64,' + Buffer.from(svg, 'utf8').toString('base64');
    const img = nativeImage.createFromDataURL(dataUrl);
    if (img && !img.isEmpty()) return img;
  } catch {}
  return null;
}
function getTrayIconImage(){
  const svgPath = path.join(__dirname, 'assets', 'simple-xp-shield.svg');
  const icoPath = path.join(__dirname, 'assets', 'simple-xp-shield.ico');
  const pngPath = path.join(__dirname, 'assets', 'tray.png');
  // Prefer ICO on Windows if provided
  if (process.platform === 'win32' && fs.existsSync(icoPath)){
    const ico = nativeImage.createFromPath(icoPath);
    if (ico && !ico.isEmpty()) return ico;
  }
  let prefer = tryCreateImageFromSvg(svgPath);
  if (!prefer || prefer.isEmpty()){
    // Embedded minimal 16x16 shield-like SVG for reliable tray rendering
    const inline = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 16 16'><rect width='16' height='16' rx='3' ry='3' fill='#2f7ed8'/><path d='M8 2 l5 2 v4 c0 3-3 5-5 6 c-2-1-5-3-5-6 v-4z' fill='#ffffff'/></svg>";
    const dataUrl = 'data:image/svg+xml;base64,' + Buffer.from(inline, 'utf8').toString('base64');
    try { prefer = nativeImage.createFromDataURL(dataUrl); } catch { prefer = null; }
  }
  if (prefer && !prefer.isEmpty()){
    const resized = prefer.resize({ width: 24, height: 24, quality: 'best' });
    if (!resized.isEmpty()) return resized;
  }
  let img = nativeImage.createFromPath(pngPath);
  if (!img || img.isEmpty()) img = nativeImage.createEmpty();
  return img;
}
function getWindowIconImage(){
  const svgPath = path.join(__dirname, 'assets', 'simple-xp-shield.svg');
  const icoPath = path.join(__dirname, 'assets', 'simple-xp-shield.ico');
  const pngPath = path.join(__dirname, 'assets', 'tray.png');
  if (process.platform === 'win32' && fs.existsSync(icoPath)){
    const ico = nativeImage.createFromPath(icoPath);
    if (ico && !ico.isEmpty()) return ico;
  }
  let prefer = tryCreateImageFromSvg(svgPath);
  if (!prefer || prefer.isEmpty()){
    const inline = "<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64' viewBox='0 0 16 16'><rect width='16' height='16' rx='3' ry='3' fill='#2f7ed8'/><path d='M8 2 l5 2 v4 c0 3-3 5-5 6 c-2-1-5-3-5-6 v-4z' fill='#ffffff'/></svg>";
    const dataUrl = 'data:image/svg+xml;base64,' + Buffer.from(inline, 'utf8').toString('base64');
    try { prefer = nativeImage.createFromDataURL(dataUrl); } catch { prefer = null; }
  }
  if (prefer && !prefer.isEmpty()){
    const resized = prefer.resize({ width: 64, height: 64, quality: 'best' });
    if (!resized.isEmpty()) return resized;
  }
  let img = nativeImage.createFromPath(pngPath);
  if (!img || img.isEmpty()) img = nativeImage.createEmpty();
  return img;
}

// Debug helper: backscan any files that still have no zone recorded
async function forceBackscanMissingZones(){
  ensureSettings();
  const dir = state.settings.logsDir;
  if (!dir || !fs.existsSync(dir)) { log('Force backscan: logsDir not set'); return; }
  const files = fs.readdirSync(dir).filter(f => /^eqlog_.+?\.txt$/i.test(f));
  let updated = 0, missing = 0;
  for (const f of files){
    const full = path.join(dir, f);
    const rec = state.latestZonesByFile[full];
    const haveZone = !!(rec && rec.zone);
    if (haveZone) continue;
    const parsed = parseLogName(f);
    const char = (parsed && parsed.name) ? parsed.name : f.replace(/^eqlog_([^_]+).*$/i,'$1');
    const lastZ = findLastZoneInFile(full, getBackscanBytes());
    if (lastZ){
      state.latestZones[char] = { zone: lastZ.zone, detectedUtcISO: lastZ.utcISO, detectedLocalISO: lastZ.localISO, sourceFile: full };
      state.latestZonesByFile[full] = { character: char, zone: lastZ.zone, detectedUtcISO: lastZ.utcISO, detectedLocalISO: lastZ.localISO, sourceFile: full };
      updated++;
      log('Force backscan seeded last zone', char, lastZ.zone);
    } else {
      missing++;
      log('Force backscan no zone found', char, full);
    }
  }
  saveState();
  await maybeWriteLocalSheets();
  await maybePostWebhook();
  log('Force backscan summary', { updated, missing });
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
ipcMain.handle('cov:getLists', async () => {
  try{
    ensureSettings();
    const defaults = String(RAW_COV||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
    const mergedSet = getCovSet();
    return {
      defaults,
      merged: Array.from(mergedSet),
      additions: Array.from((state.settings.covAdditions||[])),
      removals: Array.from((state.settings.covRemovals||[]))
    };
  } catch(e){ return { defaults: [], merged: [], additions: [], removals: [], error: String(e&&e.message||e) }; }
});
ipcMain.handle('advanced:forceBackscan', async () => {
  try { await forceBackscanMissingZones(); return { ok: true }; }
  catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
});
ipcMain.handle('advanced:replaceAll', async () => {
  try { await sendReplaceAllWebhook(); return { ok: true }; }
  catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
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
  const trayImg = getTrayIconImage();
  tray = new Tray(trayImg);
  rebuildTray();
  startScanning();
});

process.on('unhandledRejection', (err) => { log('Unhandled rejection', err && err.stack ? err.stack : String(err)); });
process.on('uncaughtException', (err) => { log('Uncaught exception', err && err.stack ? err.stack : String(err)); });
