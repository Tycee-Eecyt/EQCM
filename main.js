// EQ Character Manager — v2.0.11 — Author: Tyler A
const { app, Tray, Menu, BrowserWindow, dialog, shell, nativeImage, ipcMain, screen, nativeTheme, clipboard, Notification } = require('electron');

// Compatibility: expose a safe global fallback for older code paths
// that referenced `global.getLogId`. In this file the named helper
// below is hoisted, so this usually no-ops, but keep for integration.
if (typeof getLogId !== 'function') {
  global.getLogId = function(filePath){
    try { return require('path').basename(String(filePath||'').trim()); }
    catch { return String(filePath||''); }
  };
}

// Optional auto-update (electron-updater)
let autoUpdater = null;
try { autoUpdater = require('electron-updater').autoUpdater; }
catch {}
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
    const only = !!(state.settings && state.settings.favoritesOnly);
    const favs = (state.settings && state.settings.favorites) || [];
    if (!only || !favs.length) return rows || [];
    const set = new Set(favs.map(s => String(s).toLowerCase()));
    return (rows || []).filter(r => set.has(String(r && r[0] || '').toLowerCase()));
  } catch { return rows || []; }
}

const os = require('os');
const http = require('http');
const https = require('https');
const crypto = require('crypto');

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
// Debug logging that can be toggled from settings (Advanced)
function dlog(...args){
  try{ ensureSettings(); if (!state.settings || !state.settings.debugLogs) return; }
  catch{}
  log(...args);
}

// ---------- defaults & state ----------
const DEFAULT_SETTINGS = {
  favorites: [],
  favoritesOnly: false,
  perCharSync: {},

  webhookUrl: "",
  webhookSecret: "",
  sheetUrl: "",  // spreadsheet link (not a specific tab)

  covList: [],
  strictUnstable: false,
  debugLogs: false,

  logsDir: "",
  baseDir: "",
  scanIntervalSec: 60,

  localSheetsEnabled: true,
  localSheetsDir: "",
  remoteSheetsEnabled: true,
  remoteSheetsImmediate: true,
  // CoV Faction tab is driven by local CSV exactly (cleared and replaced on sync)

  // Backscan configuration
  backscanMaxMB: 0,             // 0 = full file; otherwise clamp 5–20 MB
  backscanRetryMinutes: 10,     // 0 disables periodic retry after first attempt

  // Consideration heuristics
  invisMaxMinutes: 20,          // treat self-invis as potentially active up to this long
  combatRecentMinutes: 5,       // treat combat as "recent" for this many minutes

  // Raider kit customization (UI in Raid Kit window)
  raidKitItems: [],             // custom additions [{ name, mode: 'present'|'count', pattern? }]
  raidKitHidden: []             // names of defaults to hide
};

let state = {
  tz: Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC',
  offsets: {},             // file -> byte offset
  latestZones: {},         // char -> { zone, detectedUtcISO, detectedLocalISO, sourceFile }
  // New: track per-log (character+server) so same name on multiple servers doesn't collide
  latestZonesByFile: {},   // filePath -> { character, zone, detectedUtcISO, detectedLocalISO, sourceFile }
  covFaction: {},          // char -> { standing, standingDisplay, score, mob, detectedUtcISO, detectedLocalISO }
  inventory: {},           // char -> { filePath, fileCreated, fileModified, items[] }
  // Track last successfully pushed zone timestamp per source log file (ISO string)
  lastPushedZones: {},     // sourceFile -> last UTC ISO sent to Google Sheets
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

function getWebhookUrl(){
  ensureSettings();
  return String(state.settings.webhookUrl || state.settings.appsScriptUrl || '').trim();
}

function getWebhookSecret(){
  ensureSettings();
  return String(state.settings.webhookSecret || state.settings.appsScriptSecret || '');
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

function dedupeStandingDisplay(value){
  const str = String(value || '').trim();
  if (!str) return '';
  const baseMatch = str.match(/^([^(]+)/);
  const base = baseMatch ? baseMatch[1].trim() : '';
  const parens = str.match(/\([^)]+\)/g) || [];
  const seen = new Set();
  const deduped = [];
  parens.forEach(seg => {
    const trimmed = seg.trim();
    if (!trimmed) return;
    if (!seen.has(trimmed)) {
      seen.add(trimmed);
      deduped.push(trimmed);
    }
  });
  if (base) {
    return deduped.length ? base + ' ' + deduped.join(' ') : base;
  }
  return deduped.join(' ');
}

function getSheetIdFromSettings(){
  ensureSettings();
  const raw = String(state.settings.sheetUrl || '').trim();
  if (!raw) return '';
  const idOnly = /^[A-Za-z0-9_-]{20,}$/;
  if (idOnly.test(raw)) return raw;
  try {
    const url = new URL(/^https?:\/\//i.test(raw) ? raw : `https://${raw}`);
    const parts = url.pathname.split('/').filter(Boolean);
    const idx = parts.indexOf('d');
    if (idx >= 0 && parts[idx + 1]) return parts[idx + 1];
  } catch {}
  return '';
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

// ---------- Raider Kit (defaults + user) ----------
const RK = require('./src/raidkit-core');
const DEFAULT_RAID_KIT = RK.DEFAULT_RAID_KIT;
const FIXED_RK_NAMES = RK.FIXED_RK_NAMES;
function getMergedRaidKit(){ ensureSettings(); return RK.getMergedRaidKit(state.settings); }
function countRaidKitForCharacter(character){
  const inv = (state.inventory||{})[character];
  const items = (inv && inv.items) ? inv.items : [];
  return RK.countRaidKitForInventory(state.settings, items);
}

// ---------- regexes ----------
const RE_ZONE = /^\[(?<ts>[^\]]+)\]\s+You have entered (?<zone>.+?)\./i;
const RE_CON  = new RegExp(String.raw`^\[(?<ts>[^\]]+)\]\s+(?<mob>.+?)\s+(?:regards you as an ally|looks upon you warmly|kindly considers you|judges you amiably|regards you indifferently|looks your way apprehensively|glowers at you dubiously|glares at you threateningly|scowls at you).*?$`, 'i');
const RE_INVIS_ON  = /(You vanish\.|Someone fades away\.|You gather shadows about you\.|Someone steps into the shadows and disappears\.)/i;
// Note: "You feel yourself starting to appear." is a transition and should still be treated as invis-on.
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

// ---------- CoV recompute (tail scan, bottom-first, skip invis) ----------
function recalcCovFactionFromTail(){
  try{
    ensureSettings();
    const logsDir = state.settings.logsDir;
    if (!logsDir || !fs.existsSync(logsDir)) return;
    const covSet = getCovSet();
    const preferScore = (a, b) => (a == null ? -1 : a) - (b == null ? -1 : b);
    const byChar = {};
    // Build mapping of character -> candidate log file (latest source if available)
    const files = {};
    Object.values(state.latestZonesByFile||{}).forEach(v => { if (v && v.character) files[v.character] = v.sourceFile; });
    // Fallback: pick any eqlog_<Name>*.txt in logsDir for characters we know
    const allLogs = fs.readdirSync(logsDir).filter(f => /^eqlog_.+?\.txt$/i.test(f));
    function pickFileFor(char){
      if (files[char] && fs.existsSync(files[char])) return files[char];
      const re = new RegExp('^eqlog_'+char.replace(/[^A-Za-z0-9_-]/g,'.')+'.+?\.txt$','i');
      const m = allLogs.find(f => re.test(f));
      return m ? path.join(logsDir, m) : '';
    }
    // For each known character in inventory or zones
    const chars = new Set([ ...Object.keys(state.inventory||{}), ...Object.keys(state.latestZones||{}), ...Object.values(state.latestZonesByFile||{}).map(v=>v.character||'') ]);
    for (const char of Array.from(chars).filter(Boolean)){
      const fp = pickFileFor(char);
      if (!fp || !fs.existsSync(fp)) continue;
      let buf = null;
      try{
        const st = fs.statSync(fp);
        const max = 200*1024; // 200KB tail is plenty for recent considers
        const fd = fs.openSync(fp,'r');
        const size = st.size;
        const start = Math.max(0, size - max);
        const len = size - start;
        const tmp = Buffer.alloc(len);
        fs.readSync(fd, tmp, 0, len, start);
        fs.closeSync(fd);
        buf = tmp.toString('utf8');
      }catch(e){ continue; }
      if (!buf) continue;
      const lines = buf.replace(/\r\n/g,'\n').split('\n');
      let selfInvis = false;
      let found = null; // { standing, score, mob, detectedUtcISO, detectedLocalISO }
      for (let i=lines.length-1; i>=0; i--){
        const line = lines[i];
        if (!line) continue;
        // Self invis markers
        if (RE_SELF_INVIS_ON.test(line)) { selfInvis = true; continue; }
        if (/You feel yourself starting to appear\./i.test(line)) { /* still invis */ continue; }
        if (RE_SELF_INVIS_OFF.test(line)) { selfInvis = false; continue; }
        const m = RE_CON.exec(line);
        if (!m) continue;
        const mob = String(m.groups && m.groups.mob || '').trim();
        const mobNorm = normalizeMobName(mob);
        if (!covSet.has(mobNorm)){
          let ok = false;
          for (const name of covSet){ if (mobNorm.startsWith(name)) { ok = true; break; } }
          if (!ok) continue;
        }
        // Determine standing
        let standing=null, score=null;
        for (const s of STANDINGS){ if (s.test.test(line)){ standing=s.key; score=s.score; break; } }
        if (!standing) continue;
        // Skip considers while invis is active
        if (selfInvis) continue;
        // Timestamps for output
        const ts = String(m.groups && m.groups.ts || '').trim();
        const when = parseEqTimestamp(ts);
        // Bottom-first guarantees first stable hit is the best/latest by position; still prefer higher score if tied by position somehow
        if (!found || (preferScore(score, found.score) > 0)){
          found = { character: char, standing, score, mob, detectedUtcISO: when.utcISO, detectedLocalISO: when.localISO };
          // Early exit if Ally: cannot be improved
          if (score === 1450) break;
        }
      }
      if (found){
        byChar[char] = found;
      }
    }
    // Commit to state.covFaction (do not downgrade existing better scores)
    for (const [char, rec] of Object.entries(byChar)){
      const prev = state.covFaction[char];
      const recScore = (rec.score ?? 0);
      const prevScore = (prev && prev.score != null) ? prev.score : -999999;
      const recWhen = String(rec.detectedUtcISO || '');
      const prevWhen = String((prev && prev.detectedUtcISO) || '');
      const shouldUpdate = (!prev) || (recScore > prevScore) || (recScore === prevScore && recWhen >= prevWhen);
      if (shouldUpdate){
        state.covFaction[char] = {
          standing: rec.standing,
          standingDisplay: rec.standing,
          score: rec.score,
          mob: rec.mob,
          detectedUtcISO: rec.detectedUtcISO,
          detectedLocalISO: rec.detectedLocalISO
        };
      }
    }
  }catch(e){ log('recalcCovFactionFromTail error', e.message); }
}
// ---------- CSV ----------
function writeCsv(filePath, header, rows){
  // Returns true if file content changed
  try{
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    const lines = [header.join(','), ...rows.map(r => r.map(v => {
      const s = String(v ?? '');
      return /[",\n]/.test(s) ? ('"' + s.replace(/"/g,'""') + '"') : s;
    }).join(','))];
    const next = lines.join('\n');
    let prev = '';
    try { if (fs.existsSync(filePath)) prev = fs.readFileSync(filePath, 'utf8'); } catch {}
    if (prev === next) return false;
    fs.writeFileSync(filePath, next, 'utf8');
    return true;
  }catch(e){ log('writeCsv error', filePath, e.message); return false; }
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

function isLikelyWebhookUrl(url){
  try{
    const u = new URL(String(url || '').trim());
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return false;
    if (!u.hostname) return false;
    if (!u.pathname || u.pathname === '/') return false;
    return true;
  } catch { return false; }
}

function getRaidKitSummary(items){
  return RK.getRaidKitSummary(items);
}

// Fixed kit columns (ordered) mapping to summary props
const FIXED_KIT_COLUMN_DEFS = [
  { name: 'Vial of Velium Vapors', header: 'Vial of Velium Vapors', prop: 'vialVeliumVapors' },
  { name: 'Leatherfoot Raider Skullcap', header: 'Leatherfoot Raider Skullcap', prop: 'leatherfootSkullcap' },
  { name: 'Shiny Brass Idol', header: 'Shiny Brass Idol', prop: 'shinyBrassIdol' },
  { name: 'Ring of Shadows', header: 'Ring of Shadows Count', prop: 'ringOfShadowsCount' },
  { name: 'Reaper of the Dead', header: 'Reaper of the Dead', prop: 'reaperOfTheDead' },
  { name: 'Pearl', header: 'Pearl Count', prop: 'pearlCount' },
  { name: 'Peridot', header: 'Peridot Count', prop: 'peridotCount' },
  { name: '10 Dose Potion of Stinging Wort', header: '10 Dose Potion of Stinging Wort Count', prop: 'tenDosePotionOfStingingWortCount' },
  { name: 'Pegasus Feather Cloak', header: 'Pegasus Feather Cloak', prop: 'pegasusFeatherCloak' },
  { name: "Larrikan's Mask", header: "Larrikan's Mask", prop: 'larrikansMask' },
  { name: 'Mana Battery - Class Five', header: 'MB Class Five', prop: 'mbClassFive' },
  { name: 'Mana Battery - Class Four', header: 'MB Class Four', prop: 'mbClassFour' },
  { name: 'Mana Battery - Class Three', header: 'MB Class Three', prop: 'mbClassThree' },
  { name: 'Mana Battery - Class Two', header: 'MB Class Two', prop: 'mbClassTwo' },
  { name: 'Mana Battery - Class One', header: 'MB Class One', prop: 'mbClassOne' }
];
function buildEnabledFixedColumns(){
  ensureSettings();
  const hidden = new Set((state.settings.raidKitHidden||[]).map(String));
  return FIXED_KIT_COLUMN_DEFS.filter(def => !hidden.has(def.name));
}
function buildLatestZoneRows(){
  ensureSettings();
  const tz = state.tz || 'UTC';
  const raw = (state.latestZonesByFile && Object.keys(state.latestZonesByFile).length)
    ? Object.values(state.latestZonesByFile).map(v => ({ character: String(v.character || ''), zone: v.zone || '', utc: String(v.detectedUtcISO || ''), local: String(v.detectedLocalISO || ''), tz, source: v.sourceFile || '' }))
    : Object.entries(state.latestZones || {}).map(([character, v]) => ({ character: String(character || ''), zone: v && v.zone ? v.zone : '', utc: String(v && v.detectedUtcISO ? v.detectedUtcISO : ''), local: String(v && v.detectedLocalISO ? v.detectedLocalISO : ''), tz, source: (v && v.sourceFile) || '' }));

  const byChar = new Map();
  raw.forEach(row => {
    const name = row.character.trim();
    if (!name) return;
    const prev = byChar.get(name);
    if (!prev) { byChar.set(name, row); return; }
    const prevUtc = prev.utc || '';
    const nextUtc = row.utc || '';
    if (!prevUtc) { byChar.set(name, row); return; }
    if (!nextUtc) return;
    if (nextUtc > prevUtc) byChar.set(name, row);
  });

  if (byChar.size === 0) {
    return raw.filter(row => row.character && row.character.trim());
  }
  return Array.from(byChar.values());
}

async function maybePostWebhook(){
  ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) return;
  const url = getWebhookUrl(); if (!url) return;
  if (!isLikelyWebhookUrl(url)) { log('Webhook not attempted: URL does not look like a valid webhook endpoint'); return; }
  const secret = getWebhookSecret();

  // Collect latest zone entries (one per character, based on most recent timestamp)
  const zoneCandidates = buildLatestZoneRows();
  let zoneRows = Array.isArray(zoneCandidates) ? [...zoneCandidates] : [];
  // Guard: only upsert if this zone timestamp is newer than what we last pushed to the sheet for that source
  try{
    const last = (state.lastPushedZones && typeof state.lastPushedZones === 'object') ? state.lastPushedZones : {};
    zoneRows = zoneRows.filter(r => {
      const src = String(r.source||'').trim();
      const when = String(r.utc||'').trim();
      if (!src || !when) return false; // skip incomplete rows
      const prev = String(last[src]||'');
      return !prev || when > prev;
    });
    if (!zoneRows.length){
      // Nothing newer to push; still allow inventory upserts below
    }
  } catch {}
  let covRows  = Object.entries(state.covFaction || {}).map(([character, v]) => ({
    character,
    standing: v.standing || '',
    standingDisplay: dedupeStandingDisplay(v.standingDisplay || ''),
    score: v.score ?? '',
    mob: v.mob || '',
    utc: v.detectedUtcISO || '',
    local: v.detectedLocalISO || ''
  }));
  let invRows  = Object.entries(state.inventory || {}).map(([character, v]) => {
    const kit = getRaidKitSummary(v.items||[]);
    const extras = buildRaidKitExtrasForCharacter(character);
    const exMap = {}; extras.forEach(e => exMap[e.header]=e.value);
    return { character, file: v.filePath||'', logFile: getLatestZoneSourceForChar(character), created: v.fileCreated||'', modified: v.fileModified||'', raidKit: kit, kitExtras: exMap };
  });
  let invDetails = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', created: v.fileCreated||'', modified: v.fileModified||'', items: v.items||[] }));

  // Apply favoritesOnly filtering to webhook payload for parity with CSV
  try{
    const only = !!(state.settings && state.settings.favoritesOnly);
    const favs = (state.settings && state.settings.favorites) || [];
    if (only && favs.length){
      const set = new Set(favs.map(s => String(s).toLowerCase()));
      const keep = (c) => set.has(String(c||'').toLowerCase());
      zoneRows = zoneRows.filter(r => keep(r.character));
      covRows  = covRows.filter(r => keep(r.character));
      invRows  = invRows.filter(r => keep(r.character));
      invDetails = invDetails.filter(r => keep(r.character));
      // To keep the sheet exactly matching the filtered CSV, do a replace when favoritesOnly is on
      try { await sendReplaceAllWebhook({}); } catch {}
      return;
    }
  }catch{}

  // Optionally drive CoV Faction from CSV exactly; when enabled, do not upsert factions JSON
  // Always drive CoV Faction from CSV; do not upsert JSON factions
  const sheetId = getSheetIdFromSettings();
  const immediate = isImmediateSheetsEnabled();
  const upserts = { zones: zoneRows, inventory: invRows, inventoryDetails: invDetails };
  const payload = { secret, upserts, immediate };
  if (sheetId) payload.sheetId = sheetId;
  try {
    const res = await postJson(url, payload);
    const debugWebhook = String(process.env.DEBUG_WEBHOOK || '').toLowerCase();
    const debugOn = debugWebhook === '1' || debugWebhook === 'true' || debugWebhook === 'yes';
    const is2xx = (res.status >= 200 && res.status < 300);
    if (!is2xx || debugOn) {
      const bodyPreview = (res.body || '').slice(0, 180);
      log('Webhook response', res.status, bodyPreview);
    }
    if (is2xx && upserts && Array.isArray(upserts.zones) && upserts.zones.length){
      try{
        if (!state.lastPushedZones || typeof state.lastPushedZones !== 'object') state.lastPushedZones = {};
        upserts.zones.forEach(r => {
          const src = String(r.source||'').trim();
          const when = String(r.utc||'').trim();
          if (!src || !when) return;
          const prev = String(state.lastPushedZones[src]||'');
          if (!prev || when > prev) state.lastPushedZones[src] = when;
        });
        saveState();
      } catch {}
    }
    // No CSV push here; handled unconditionally in scan cycle to ensure sheet matches CSV every interval
  }
  catch(e){ log('Webhook error', e.message); }
}

// Fetch unique characters currently present on the Google Sheet via webhook API
async function fetchSheetCharacters(){
  try{
    ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) return { ok:false, error:'remoteSheetsDisabled' };
  const url = getWebhookUrl();
  if (!url || !isLikelyWebhookUrl(url)) return { ok:false, error:'invalidWebhookUrl' };
  const secret = getWebhookSecret();
  const sheetId = getSheetIdFromSettings();
  const payload = { action: 'listCharacters', secret };
  if (sheetId) payload.sheetId = sheetId;
    const res = await postJson(url, payload);
    if (!(res.status >= 200 && res.status < 300)) return { ok:false, error:'http_'+res.status };
    let body = {};
    try { body = JSON.parse(String(res.body||'{}')); } catch(e){ return { ok:false, error:'parse-error' }; }
    if (body && body.ok === true && Array.isArray(body.characters)){
      return { ok:true, characters: body.characters };
    }
    return { ok:false, error: String(body && body.error || 'unknown') };
  }catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
}

// Replace CoV Faction sheet with the exact contents of local CSV
async function sendReplaceFactionsCsvFromLocal(){
  ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) { log('replaceFactionsCsv skipped: remoteSheetsEnabled=false'); return; }
  const url = getWebhookUrl();
  if (!url || !isLikelyWebhookUrl(url)) { log('replaceFactionsCsv aborted: invalid webhook URL'); throw new Error('Invalid webhook URL'); }
  const secret = getWebhookSecret();
  const sheetId = getSheetIdFromSettings();
  // Determine CSV path from localSheetsDir (or default SHEETS_DIR)
  let dir = (state.settings.localSheetsDir && state.settings.localSheetsDir.trim()) ? state.settings.localSheetsDir.trim() : SHEETS_DIR;
  let filePath = path.join(dir, 'CoV Faction.csv');
  if (!fs.existsSync(filePath)){
    // Fallback to dev relative path
    const alt = path.join(__dirname, 'sheets-out', 'CoV Faction.csv');
    if (fs.existsSync(alt)) filePath = alt;
  }
  if (!fs.existsSync(filePath)) { const msg = 'CoV Faction.csv not found in local CSV folder'; log(msg, dir); throw new Error(msg + ': ' + dir); }
  const csv = fs.readFileSync(filePath, 'utf8');
  try{
    dlog('replaceFactionsCsv path', filePath);
    const immediate = isImmediateSheetsEnabled();
    const payload = { secret, action: 'replaceFactionsCsv', csv, immediate };
    if (sheetId) payload.sheetId = sheetId;
    const res = await postJson(url, payload);
    log('ReplaceFactionsCsv response', res.status, (res.body||'').slice(0, 180));
    if (!(res.status >= 200 && res.status < 300)) throw new Error('HTTP ' + res.status);
    return { ok: true };
  }catch(e){ log('ReplaceFactionsCsv error', e && e.message || e); throw e; }
}
// Extra raid kit beyond fixed columns
function buildRaidKitExtrasForCharacter(character){
  const inv = (state.inventory||{})[character];
  const items = (inv && inv.items) ? inv.items : [];
  const merged = getMergedRaidKit();
  const extras = [];
  for (const k of merged){
    if (FIXED_RK_NAMES.has(k.name)) continue;
    const re = new RegExp(k.pattern||('^'+k.name+'$'), 'i');
    let count=0, present=false;
    for (const it of (items||[])){
      if (re.test(String(it.Name||''))){ present=true; count += Number(it.Count||0); }
    }
    const header = k.mode==='count' ? (k.name + ' Count') : k.name;
    const value = k.mode==='count' ? (count>0 ? count : '') : (present?'Y':'N');
    extras.push({ header, value });
  }
  return extras;
}

// One-shot replace-all import via webhook API
async function sendReplaceAllWebhook(opts){
  const notify = !!(opts && (opts.notify === true));
  const forceEnv = String(process.env.FORCE_REPLACE_ALL || '').toLowerCase();
  const forceFlag = (opts && opts.force === true) || (forceEnv === '1' || forceEnv === 'true' || forceEnv === 'yes');
  ensureSettings();
  if (state.settings.remoteSheetsEnabled === false) { log('ReplaceAll skipped: remoteSheetsEnabled=false'); return; }
  const url = getWebhookUrl();
  if (!url || !isLikelyWebhookUrl(url)) { log('ReplaceAll aborted: invalid webhook URL'); return; }
  const secret = getWebhookSecret();
  const sheetId = getSheetIdFromSettings();

  const zoneCandidates = buildLatestZoneRows();
  let zoneRows = Array.isArray(zoneCandidates) ? [...zoneCandidates] : [];
  let covRows  = Object.entries(state.covFaction || {}).map(([character, v]) => ({
    character,
    standing: v.standing || '',
    standingDisplay: dedupeStandingDisplay(v.standingDisplay || ''),
    score: v.score ?? '',
    mob: v.mob || '',
    utc: v.detectedUtcISO || '',
    local: v.detectedLocalISO || ''
  }));
  let invRows  = Object.entries(state.inventory || {}).map(([character, v]) => {
    const kit = getRaidKitSummary(v.items||[]);
    const extras = buildRaidKitExtrasForCharacter(character);
    const exMap = {}; extras.forEach(e => exMap[e.header]=e.value);
    return { character, file: v.filePath||'', logFile: getLatestZoneSourceForChar(character), created: v.fileCreated||'', modified: v.fileModified||'', raidKit: kit, kitExtras: exMap };
  });
  let invDetails = Object.entries(state.inventory || {}).map(([character, v]) => ({ character, file: v.filePath||'', created: v.fileCreated||'', modified: v.fileModified||'', items: v.items||[] }));

  // Apply favoritesOnly filtering for ReplaceAll as well so sheet matches CSV exactly
  try{
    const only = !!(state.settings && state.settings.favoritesOnly);
    const favs = (state.settings && state.settings.favorites) || [];
    if (only && favs.length){
      const set = new Set(favs.map(s => String(s).toLowerCase()));
      const keep = (c) => set.has(String(c||'').toLowerCase());
      zoneRows = zoneRows.filter(r => keep(r.character));
      covRows  = covRows.filter(r => keep(r.character));
      invRows  = invRows.filter(r => keep(r.character));
      invDetails = invDetails.filter(r => keep(r.character));
    }
  }catch{}

  const enabledFixed = buildEnabledFixedColumns();
  const meta = { invFixedHeaders: enabledFixed.map(d=>d.header), invFixedProps: enabledFixed.map(d=>d.prop) };
  // Always drive CoV Faction from CSV; do not include JSON factions in ReplaceAll
  const upserts = { zones: zoneRows, inventory: invRows, inventoryDetails: invDetails };
  const immediate = isImmediateSheetsEnabled();
  const payload = { secret, action: 'replaceAll', meta, upserts, immediate };
  if (sheetId) payload.sheetId = sheetId;

  // Debounce: compute digest of what would be sent (excluding secret)
  try{
    const sortKeys = (v) => {
      if (!v || typeof v !== 'object') return v;
      if (Array.isArray(v)) return v.map(sortKeys);
      const out = {};
      Object.keys(v).sort().forEach(k => { out[k] = sortKeys(v[k]); });
      return out;
    };
    const digestObj = { action: 'replaceAll', meta, upserts };
    const json = JSON.stringify(sortKeys(digestObj));
    const digest = crypto.createHash('sha256').update(json).digest('hex');
    if (state.lastReplaceAllDigest && state.lastReplaceAllDigest === digest && !forceFlag){
      log('ReplaceAll skipped: no data changes (digest match)');
      return;
    }
    payload.__digest = digest; // optional for debugging
  } catch(e){ /* ignore digest errors */ }
  try{
    const res = await postJson(url, payload);
    log('ReplaceAll response', res.status, (res.body||'').slice(0, 180));
    if (res.status >= 200 && res.status < 300){
      if (payload.__digest) state.lastReplaceAllDigest = payload.__digest;
      try{
        // Update lastPushedZones for all zone rows included in replaceAll
        if (!state.lastPushedZones || typeof state.lastPushedZones !== 'object') state.lastPushedZones = {};
        (upserts.zones||[]).forEach(r => {
          const src = String(r.source||'').trim();
          const when = String(r.utc||'').trim();
          if (!src || !when) return;
          const prev = String(state.lastPushedZones[src]||'');
          if (!prev || when > prev) state.lastPushedZones[src] = when;
        });
      } catch {}
      saveState();
    }
    // Always follow ReplaceAll with exact CSV replace for CoV Faction
    try { await sendReplaceFactionsCsvFromLocal(); } catch(e){ log('factionsFromCsv replaceAll follow-up error', e && e.message || e); }
  }catch(e){
    log('ReplaceAll error', e.message);
  }
}

// New: push a single character inventory to a new sheet tab (like the screenshot)
async function pushInventoryToNewSheet(character){
  ensureSettings();
  if (!character) return;
  const url = getWebhookUrl();
  if (!url || !isLikelyWebhookUrl(url)) { log('Push inventory aborted: invalid webhook URL'); return; }
  const secret = getWebhookSecret();
  const sheetId = getSheetIdFromSettings();
  const inv = state.inventory[character];
  if (!inv || !Array.isArray(inv.items)) { log('Push inventory aborted: no inventory for', character); return; }
  const sheetName = `Inventory - ${character}`;
  const payload = {
    secret,
    ...(sheetId ? { sheetId } : {}),
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
          if (COV_SET.has(mobNorm)) isCov = true;
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
                state.covFaction[char] = { standing: lastMob.standing, standingDisplay: dedupeStandingDisplay(lastMob.standing + note), score: lastMob.score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                dlog('Combat-biased /con, applied mob fallback', char, mob);
              } else if (lastChar){
                const prev = nowState.prevBeforeCombat;
                const note = prev && prev.standing ? ' (combat; prev=' + prev.standing + ')' : ' (combat)';
                state.covFaction[char] = Object.assign({}, lastChar, { standingDisplay: dedupeStandingDisplay((lastChar.standingDisplay||lastChar.standing||'') + note), mob: lastChar.mob || mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO });
                dlog('Combat-biased /con, applied char fallback', char, mob);
              } else {
                const prev = nowState.prevBeforeCombat;
                const note = prev && prev.standing ? ' (combat?; prev=' + prev.standing + ')' : ' (combat?)';
                state.covFaction[char] = { standing, standingDisplay: dedupeStandingDisplay(standing + note), score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                dlog('Combat-biased /con, accepted as baseline', char, mob);
              }
            }
            else if (preferFallbackForInvis || unstableInvis){
              if (lastMob){
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ` (invis; prev=${prev.standing})` : ' (invis)';
                state.covFaction[char] = { standing: lastMob.standing, standingDisplay: dedupeStandingDisplay(lastMob.standing + note), score: lastMob.score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                dlog('Invis-biased /con, applied mob fallback', char, mob);
              } else if (lastChar){
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ' (invis; prev=' + prev.standing + ')' : ' (invis)';
                state.covFaction[char] = Object.assign({}, lastChar, { standingDisplay: dedupeStandingDisplay((lastChar.standingDisplay||lastChar.standing||'') + note), mob: lastChar.mob || mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO });
                dlog('Invis-biased /con, applied char fallback', char, mob);
              } else {
                const prev = nowState.prevBeforeInvis;
                const note = prev && prev.standing ? ' (invis?; prev=' + prev.standing + ')' : ' (invis?)';
                state.covFaction[char] = { standing, standingDisplay: dedupeStandingDisplay(standing + note), score, mob, detectedUtcISO: t.utcISO, detectedLocalISO: t.localISO };
                dlog('Invis-biased /con, accepted as baseline', char, mob);
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
    // Recompute CoV standings bottom-first and skip invis considers
    try { recalcCovFactionFromTail(); } catch {}
    saveState();
    const changed = await maybeWriteLocalSheets();
    if (changed) await maybePostWebhook();
    // Always enforce CoV Faction from CSV each interval so manual CSV edits are reflected
    try { await sendReplaceFactionsCsvFromLocal(); } catch {}
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
  const zoneCandidates = buildLatestZoneRows();
  const zRows = (Array.isArray(zoneCandidates) ? zoneCandidates : []).map(o => [o.character, o.zone, o.utc, o.local, o.tz, o.source]);
  const zRowsOut = filterRowsByFavorites(zRows);
  let changed = false;
  changed = writeCsv(path.join(dir, 'Zone Tracker.csv'), zHead, zRowsOut) || changed;

  const fHead = ['Character','Standing','Score','Mob','Consider Time (UTC)','Consider Time (Local)','Notes'];
  const fRows = Object.entries(state.covFaction || {}).map(([char, v]) => [
    char,
    v.standing || '',
    v.score ?? '',
    v.mob || '',
    v.detectedUtcISO || '',
    v.detectedLocalISO || '',
    dedupeStandingDisplay(v.standingDisplay || '')
  ]);
const fRowsOut = filterRowsByFavorites(fRows);
  changed = writeCsv(path.join(dir, 'CoV Faction.csv'), fHead, fRowsOut) || changed;

  const enabledFixed = buildEnabledFixedColumns();
  const fixedHeaders = enabledFixed.map(d => d.header);
  const iHead = ['Character','Log ID','Inventory File','Source Log File','Created (UTC)','Modified (UTC)']
                 .concat(fixedHeaders)
                 .concat(['Spreadsheet URL','Suggested Sheet Name']);
  // Determine dynamic extra headers from merged raid kit (excluding fixed)
  const extraSet = new Set();
  for (const ch of Object.keys(state.inventory||{})){
    const extras = buildRaidKitExtrasForCharacter(ch);
    extras.forEach(e => extraSet.add(e.header));
  }
  const extraHeaders = Array.from(extraSet);
  const fullHead = iHead.concat(extraHeaders);
const iRows = Object.entries(state.inventory || {}).map(([char, v]) => {
  const kit = getRaidKitSummary(v.items||[]);
  const suggested = `Inventory - ${char}`;
  const fixedVals = enabledFixed.map(d => kit[d.prop] ?? (d.prop.endsWith('Count') || d.prop.startsWith('mbClass') ? 0 : ''));
  const baseRow = [char, getLogId(v.filePath||''), v.filePath||'', getLatestZoneSourceForChar(char), v.fileCreated||'', v.fileModified||'']
          .concat(fixedVals)
          .concat([(state.settings.sheetUrl||''), suggested]);
  const exList = buildRaidKitExtrasForCharacter(char);
  const exMap = {}; exList.forEach(e => exMap[e.header]=e.value);
  const extraVals = extraHeaders.map(h => exMap[h] ?? '');
  return baseRow.concat(extraVals);
});
const iRowsOut = filterRowsByFavorites(iRows);
changed = writeCsv(path.join(dir, 'Raid Kit.csv'), fullHead, iRowsOut) || changed;

  // per-character CSV
  for (const [char, inv] of Object.entries(state.inventory || {})){
    const rows = (inv.items||[]).map(it => [char, inv.filePath||'', inv.fileCreated||'', inv.fileModified||'', it.Location||'', it.Name||'', it.ID||'', it.Count||0, it.Slots||0]);
    const header = ['Character','Inventory File','Created (UTC)','Modified (UTC)','Location','Name','ID','Count','Slots'];
    changed = writeCsv(path.join(dir, `Inventory Items - ${char}.csv`), header, rows) || changed;
  }
  return changed;
}

// ---------- Players-in-zone extraction ----------
function getLastModifiedEqLogFile(){
  try {
    ensureSettings();
    const dir = state.settings.logsDir;
    if (!dir || !fs.existsSync(dir)) return '';
    const files = fs.readdirSync(dir)
      .filter(f => /^eqlog_.+?\.txt$/i.test(f))
      .map(f => path.join(dir, f));
    if (!files.length) return '';
    let best = files[0];
    let bestM = 0;
    for (const fp of files){
      try {
        const st = fs.statSync(fp);
        const m = st.mtimeMs || (st.mtime ? st.mtime.getTime() : 0) || 0;
        if (m >= bestM){ bestM = m; best = fp; }
      } catch {}
    }
    return best;
  } catch { return ''; }
}

function findLatestPlayersBlockInText(text){
  try{
    if (!text) return null;
    const s = String(text).replace(/\r\n/g,'\n');
    const reStart = /\[[^\]]+\]\s*Players on EverQuest:/g;
    let lastIdx = -1, m;
    while ((m = reStart.exec(s))){ lastIdx = m.index; }
    if (lastIdx < 0) return null;
    const tail = s.substring(lastIdx);
    const lines = tail.split('\n');
    const out = [];
    for (let i=0;i<lines.length;i++){
      const line = lines[i];
      if (!line) break;
      out.push(line);
      if (/^\[[^\]]+\]\s*There are\s+\d+\s+players\s+in\s+.+\./i.test(line)) break;
    }
    if (out.length < 3) return null;
    const tsStrip = (l) => l.replace(/^\[[^\]]+\]\s*/, '');
    const clean = out.map(tsStrip).join('\n');
    return { raw: out.join('\n'), clean };
  }catch{ return null; }
}

function copyLatestPlayersToClipboard(){
  try{
    const filePath = getLastModifiedEqLogFile();
    if (!filePath){
      try { dialog.showMessageBox({ type:'warning', buttons:['OK'], title:'Logs Not Found', message:'No EverQuest logs found', detail:'Set your Logs folder in Settings and try again.' }); } catch {}
      return { ok:false, error:'no-logs' };
    }
    let text = '';
    try{
      const st = fs.statSync(filePath);
      const max = 512*1024; // 512KB tail
      const fd = fs.openSync(filePath,'r');
      const size = st.size;
      const start = Math.max(0, size - max);
      const len = size - start;
      const buf = Buffer.alloc(len);
      fs.readSync(fd, buf, 0, len, start);
      fs.closeSync(fd);
      text = buf.toString('utf8');
    }catch(e){
      try { dialog.showMessageBox({ type:'error', buttons:['OK'], title:'Read Error', message:'Could not read latest log file', detail:String(e && e.message || e) }); } catch {}
      return { ok:false, error:'read-error' };
    }
    const block = findLatestPlayersBlockInText(text);
    if (!block){
      try { dialog.showMessageBox({ type:'info', buttons:['OK'], title:'No Players Block', message:'No recent "Players on EverQuest" block found in the log tail.' }); } catch {}
      return { ok:false, error:'not-found' };
    }
    clipboard.writeText(block.clean);
    try { (new Notification({ title: 'Players copied', body: 'Latest players-in-zone list copied to clipboard.' })).show(); } catch {}
    return { ok:true };
  }catch(e){ return { ok:false, error:String(e&&e.message||e) }; }
}

// ---------- UI ----------
let tray = null, settingsWin = null, favoritesWin = null;
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
  const version = (() => { try { return app.getVersion && app.getVersion(); } catch { return ''; } })();
  const scanChoices = [30,60,120,300];
  const scanSub = { label: 'Scan interval', submenu: scanChoices.map(sec => ({ label: `${sec}s${(state.settings.scanIntervalSec===sec) ? ' ✓':''}`, click: () => { state.settings.scanIntervalSec = sec; saveSettings(); restartScanning(); rebuildTray(); } })) };
  return Menu.buildFromTemplate([
    { label: version ? `EQ Character Manager v${version}` : 'EQ Character Manager', enabled: false },
    { type: 'separator' },
    { label: 'Favorite Characters…', click: openFavoritesWindow },
    { label: 'Copy Last Log', click: () => { try { copyLatestPlayersToClipboard(); } catch {} } },
    buildPushInventorySubmenu(),
    { label: 'Raid Kit…', click: openRaidKitWindow },
    { label: 'CoV Mob List…', click: openCovWindow },
    { label: 'Settings…', click: openSettingsWindow },
    { type: 'separator' },
    { label: 'Open Google Sheet', click: openGoogleSheet },
    { label: 'Open data folder', click: () => { shell.openPath(DATA_DIR); } },
    { label: 'Open local CSV folder', click: () => {
        const outDir = (state.settings.localSheetsDir && state.settings.localSheetsDir.trim()) ? state.settings.localSheetsDir.trim() : SHEETS_DIR;
        shell.openPath(outDir);
      }
    },
    { type: 'separator' },
    { label: 'Rescan now', click: () => { doScanCycle(); } },
    { label: scanTimer ? 'Pause scanning' : 'Start scanning', click: () => { scanTimer ? stopScanning() : startScanning(); rebuildTray(); } },
    scanSub,
    { type: 'separator' },
    { label: 'Getting Started', click: openDocsWindow },
    { label: 'Check for updates…', click: () => { try { manualCheckForUpdates(); } catch {} } },
    { label: 'Open Releases…', click: () => { try { shell.openExternal('https://github.com/Tycee-Eecyt/EQCM/releases'); } catch {} } },
    { label: 'Advanced…', click: openAdvancedWindow },
    { type: 'separator' },
    { label: 'Quit', click: () => { quitting = true; app.quit(); } }
  ]);
}
function rebuildTray(){ if (!tray) return; tray.setContextMenu(buildMenu()); tray.setToolTip(withVersionTooltip(buildTrayTooltip())); }
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
function openFavoritesWindow(){
  if (favoritesWin && !favoritesWin.isDestroyed()) { favoritesWin.show(); favoritesWin.focus(); return; }
  favoritesWin = new BrowserWindow({ width: 700, height: 640, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  favoritesWin.setMenu(null);
  favoritesWin.loadFile(path.join(__dirname, 'favorites.html'));
  makeHidable(favoritesWin);
  favoritesWin.on('closed', () => favoritesWin = null);
}
function openGoogleSheet(){
  ensureSettings();
  const url = String(state.settings.sheetUrl || '').trim();
  if (url){
    try { shell.openExternal(url); } catch (err) { log('Open sheet error', err && err.message || err); }
    return;
  }
  try {
    dialog.showMessageBox({ type: 'info', buttons: ['OK'], defaultId: 0, title: 'Google Sheet', message: 'No Google Sheet URL configured.', detail: 'Set the Sheet URL on the Settings tab first.' });
  } catch (err) {}
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
function openRaidKitWindow(){
  const win = new BrowserWindow({ width: 900, height: 740, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'raider-kit.html'));
  makeHidable(win);
}

// ---------- Auto Updates ----------
function setupAutoUpdates(){
  try{
    if (!autoUpdater) return;
    // Basic listeners for visibility + logging
    autoUpdater.on('checking-for-update', () => log('AutoUpdate checking for update'));
    autoUpdater.on('update-available', (info) => log('AutoUpdate update available', info && (info.version||'')));
    autoUpdater.on('update-not-available', () => log('AutoUpdate no update available'));
    autoUpdater.on('error', (err) => log('AutoUpdate error', err && err.message || String(err)));
    autoUpdater.on('download-progress', (p) => {
      try{ if (p && typeof p.percent === 'number') dlog('AutoUpdate progress', Math.round(p.percent) + '%'); } catch{}
    });
    autoUpdater.on('update-downloaded', async (info) => {
      log('AutoUpdate update downloaded', info && (info.version||''));
      try{
        const res = await dialog.showMessageBox({
          type: 'question', buttons: ['Restart Now','Later'], defaultId: 0, cancelId: 1,
          title: 'Update Ready', message: 'An update has been downloaded. Restart now to apply it?'
        });
        if (res.response === 0) {
          setImmediate(() => { try { autoUpdater.quitAndInstall(); } catch {} });
        }
      }catch{}
    });
  }catch{}
}

// Manual update check with user feedback dialogs
async function manualCheckForUpdates(){
  try{
    if (!autoUpdater){
      const res = await dialog.showMessageBox({
        type: 'info', buttons: ['Open Releases','Close'], defaultId: 0, cancelId: 1,
        title: 'Updates', message: 'Auto-updater not available.',
        detail: 'Open the Releases page to download updates manually.'
      });
      if (res.response === 0){ try { shell.openExternal('https://github.com/Tycee-Eecyt/EQCM/releases'); } catch {} }
      return;
    }
    const current = app.getVersion();
    let info = null;
    try {
      const result = await autoUpdater.checkForUpdates();
      info = result && result.updateInfo ? result.updateInfo : null;
    } catch (e) {
      // The library can throw on "no update" in some cases; fall back to event path
      info = null;
    }
    const next = info && info.version ? String(info.version) : '';
    if (next && next !== current){
      await dialog.showMessageBox({
        type: 'info', buttons: ['OK'], defaultId: 0,
        title: 'Update Available',
        message: `Version ${next} is available`,
        detail: 'The update will download in the background. You will be prompted to restart when it is ready.'
      });
    } else {
      await dialog.showMessageBox({
        type: 'info', buttons: ['OK'], defaultId: 0,
        title: 'Up to Date',
        message: `You are on the latest version (${current}).`
      });
    }
  }catch(e){
    try{
      await dialog.showMessageBox({
        type: 'error', buttons: ['OK'], defaultId: 0,
        title: 'Update Check Failed',
        message: 'Could not check for updates.',
        detail: String(e && e.message || e)
      });
    }catch{}
  }
}

// Prefix app version into the tray tooltip title
function withVersionTooltip(s){
  try {
    const v = app.getVersion && app.getVersion();
    if (v) return String(s||'').replace(/^EQ Character Manager\b/, `EQ Character Manager v${v}`);
  } catch {}
  return s;
}

// ---------- Installer / Onboarding ----------
function needsOnboarding(){
  try{
    ensureSettings();
    if (state.settings.onboardingCompleted) return false;
    const logsOk = !!(state.settings.logsDir && fs.existsSync(state.settings.logsDir));
    const baseOk = !!(state.settings.baseDir && fs.existsSync(state.settings.baseDir));
    const sheetOk = !!String(state.settings.sheetUrl||'').trim();
    const webhookOk = isLikelyWebhookUrl(getWebhookUrl());
    return !(sheetOk && webhookOk && logsOk && baseOk);
  }catch{ return true; }
}
function openInstallerWindow(){
  try{
    const win = new BrowserWindow({ width: 880, height: 720, resizable: true, icon: getWindowIconImage(), webPreferences: { contextIsolation: true, preload: path.join(__dirname, 'renderer.js') } });
    win.setMenu(null);
    win.loadFile(path.join(__dirname, 'installer.html'));
    makeHidable(win);
  }catch(e){ log('openInstallerWindow error', e && e.message || e); }
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
  // 1) Prefer size-specific PNGs if present: assets/tray-16.png, tray-24.png, tray-32.png, ...
  try {
    const sizes = [16, 20, 24, 32, 48, 64, 128, 256];
    const available = sizes
      .map(sz => ({ sz, fp: path.join(__dirname, 'assets', `tray-${sz}.png`) }))
      .filter(x => fs.existsSync(x.fp));
    if (available.length) {
      const sf = (() => { try { return (screen && screen.getPrimaryDisplay && screen.getPrimaryDisplay().scaleFactor) || 1; } catch { return 1; } })();
      const targetPx = Math.max(16, Math.round(16 * sf));
      let pick = null;
      for (const a of available.sort((a,b) => a.sz - b.sz)) {
        if (!pick) pick = a;
        if (a.sz >= targetPx) { pick = a; break; }
      }
      const img = nativeImage.createFromPath(pick.fp);
      if (img && !img.isEmpty()) return img;
    }
  } catch {}

  // 2) Theme-specific fallbacks if present
  try {
    const dark = nativeTheme && nativeTheme.shouldUseDarkColors;
    const themed = path.join(__dirname, 'assets', dark ? 'tray-light.png' : 'tray-dark.png');
    if (fs.existsSync(themed)){
      const img = nativeImage.createFromPath(themed);
      if (img && !img.isEmpty()) return img;
    }
  } catch {}

  // 3) Windows ICO if provided (best multi-size support on Windows)
  const icoPath = path.join(__dirname, 'assets', 'simple-xp-shield.ico');
  if (process.platform === 'win32' && fs.existsSync(icoPath)){
    const ico = nativeImage.createFromPath(icoPath);
    if (ico && !ico.isEmpty()) return ico;
  }

  // 4) Try the shield SVG (or embedded), resized to a sensible tray size
  const svgPath = path.join(__dirname, 'assets', 'simple-xp-shield.svg');
  let prefer = tryCreateImageFromSvg(svgPath);
  if (!prefer || prefer.isEmpty()){
    const inline = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 16 16'><rect width='16' height='16' rx='3' ry='3' fill='#2f7ed8'/><path d='M8 2 l5 2 v4 c0 3-3 5-5 6 c-2-1-5-3-5-6 v-4z' fill='#ffffff'/></svg>";
    const dataUrl = 'data:image/svg+xml;base64,' + Buffer.from(inline, 'utf8').toString('base64');
    try { prefer = nativeImage.createFromDataURL(dataUrl); } catch { prefer = null; }
  }
  if (prefer && !prefer.isEmpty()){
    const resized = prefer.resize({ width: 24, height: 24, quality: 'best' });
    if (!resized.isEmpty()) return resized;
  }

  // 5) Final fallback: assets/tray.png if present, otherwise empty image
  const pngPath = path.join(__dirname, 'assets', 'tray.png');
  let img = nativeImage.createFromPath(pngPath);
  if (!img || img.isEmpty()) img = nativeImage.createEmpty();
  return img;
}
function getWindowIconImage(){
  // Prefer tray-256.png explicitly to match tray art
  try {
    const tray256 = path.join(__dirname, 'assets', 'tray-256.png');
    if (fs.existsSync(tray256)){
      const img = nativeImage.createFromPath(tray256);
      if (img && !img.isEmpty()) return img;
    }
  } catch {}
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

// Generate a minimal .ico (PNG-based) from shield SVG if not present
function ensureShieldIco(){
  try {
    const icoPath = path.join(__dirname, 'assets', 'simple-xp-shield.ico');
    if (fs.existsSync(icoPath)) return;
    const svgPath = path.join(__dirname, 'assets', 'simple-xp-shield.svg');
    const imgSvg = tryCreateImageFromSvg(svgPath);
    if (!imgSvg || imgSvg.isEmpty()) return;
    const icon32 = imgSvg.resize({ width: 32, height: 32, quality: 'best' });
    const png = icon32.toPNG();
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0,0); // reserved
    header.writeUInt16LE(1,2); // type icon
    header.writeUInt16LE(1,4); // count
    const dir = Buffer.alloc(16);
    dir.writeUInt8(32,0); // width
    dir.writeUInt8(32,1); // height
    dir.writeUInt8(0,2);  // colors
    dir.writeUInt8(0,3);  // reserved
    dir.writeUInt16LE(1,4); // planes
    dir.writeUInt16LE(32,6); // bitcount
    dir.writeUInt32LE(png.length,8);
    dir.writeUInt32LE(6+16,12);
    const out = Buffer.concat([header, dir, png]);
    fs.writeFileSync(icoPath, out);
  } catch {}
}

// Generate a tray.png raster fallback from the shield SVG if missing
function ensureTrayPng(){
  try {
    const pngPath = path.join(__dirname, 'assets', 'tray.png');
    if (fs.existsSync(pngPath)) return;
    const svgPath = path.join(__dirname, 'assets', 'simple-xp-shield.svg');
    const imgSvg = tryCreateImageFromSvg(svgPath);
    if (!imgSvg || imgSvg.isEmpty()) return;
    const icon24 = imgSvg.resize({ width: 24, height: 24, quality: 'best' });
    const buf = icon24.toPNG();
    fs.writeFileSync(pngPath, Buffer.from(buf));
  } catch {}
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
  const changed = await maybeWriteLocalSheets();
  if (changed) await maybePostWebhook();
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

// Centralize external URL opening in main for reliability
ipcMain.handle('shell:openExternal', async (_evt, url) => {
  try {
    const u = String(url || '').trim();
    if (!u) return false;
    const result = await Promise.resolve(shell.openExternal(u));
    return (typeof result === 'undefined') ? true : !!result;
  } catch (e) {
    try { log('shell:openExternal error', e && e.message || e); } catch {}
    return false;
  }
});

ipcMain.handle('settings:set', async (evt, payload) => {
  ensureSettings();
  // Capture previous favorites state to detect changes
  const prevOnly = !!state.settings.favoritesOnly;
  const prevFavs = Array.isArray(state.settings.favorites) ? state.settings.favorites.map(String).sort() : [];

  state.settings = Object.assign({}, state.settings, payload || {});
  saveSettings(); rebuildTray();

  try {
    const nextOnly = !!state.settings.favoritesOnly;
    const nextFavs = Array.isArray(state.settings.favorites) ? state.settings.favorites.map(String).sort() : [];
    const changedOnly = prevOnly !== nextOnly;
    const changedFavs = prevFavs.length !== nextFavs.length || prevFavs.some((v,i)=>v!==nextFavs[i]);
    if (changedOnly || changedFavs) {
      // Kick a Replace All so the Google Sheet reflects the new favorites filter immediately
      setImmediate(() => { try { sendReplaceAllWebhook({ force: true }); } catch {} });
    }
  } catch {}

  return { ok: true };
});
ipcMain.handle('favorites:listFromSheet', async () => {
  try { return await fetchSheetCharacters(); }
  catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
});
ipcMain.handle('players:copyLatest', async () => {
  try { return copyLatestPlayersToClipboard(); }
  catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
});
ipcMain.handle('raidkit:get', async () => {
  try{
    ensureSettings();
    return { defaults: DEFAULT_RAID_KIT, custom: state.settings.raidKitItems||[], hidden: state.settings.raidKitHidden||[], merged: getMergedRaidKit() };
  } catch(e){ return { defaults: [], custom: [], hidden: [], merged: [], error: String(e&&e.message||e) }; }
});
ipcMain.handle('raidkit:set', async (evt, payload) => {
  try{
    ensureSettings();
    const items = Array.isArray(payload && payload.items) ? payload.items : (state.settings.raidKitItems||[]);
    const hidden = Array.isArray(payload && payload.hidden) ? payload.hidden : (state.settings.raidKitHidden||[]);
    state.settings.raidKitItems = items;
    state.settings.raidKitHidden = hidden;
    saveSettings();
    return { ok: true };
  } catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
});
ipcMain.handle('raidkit:saveAndPush', async (evt, payload) => {
  try{
    ensureSettings();
    const items = Array.isArray(payload && payload.items) ? payload.items : (state.settings.raidKitItems||[]);
    const hidden = Array.isArray(payload && payload.hidden) ? payload.hidden : (state.settings.raidKitHidden||[]);
    state.settings.raidKitItems = items;
    state.settings.raidKitHidden = hidden;
    saveSettings();
    // Trigger push in background without blocking the UI
    setImmediate(() => { sendReplaceAllWebhook().catch(() => {}); });
    return { ok: true };
  } catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
});
ipcMain.handle('raidkit:counts', async (evt, character) => {
  try{ return { ok:true, character, rows: countRaidKitForCharacter(String(character||'')) }; }
  catch(e){ return { ok:false, error: String(e&&e.message||e) }; }
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
ipcMain.handle('advanced:replaceAll', async (_evt, opts) => {
  try { await sendReplaceAllWebhook(opts||{}); return { ok: true }; }
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

// Copy bundled docs/Code.gs to clipboard and notify
ipcMain.handle('docs:copyCodeGs', async () => {
  try {
    const fp = path.join(__dirname, 'docs', 'Code.gs');
    const code = fs.readFileSync(fp, 'utf8');
    clipboard.writeText(code);
    try { (new Notification({ title: 'Copied', body: 'Code.gs copied to clipboard.' })).show(); } catch {}
    return { ok: true };
  } catch (e) {
    try { dialog.showMessageBox({ type:'error', buttons:['OK'], title:'Copy Failed', message:'Could not copy Code.gs', detail:String(e && e.message || e) }); } catch {}
    return { ok:false, error: String(e && e.message || e) };
  }
});
app.whenReady().then(() => {
  try { ensureShieldIco(); } catch {}
  try { ensureTrayPng(); } catch {}
  const trayImg = getTrayIconImage();
  tray = new Tray(trayImg);
  // Left and right click both show the menu; left avoids accidental window popups
  try { tray.on('click', () => { tray.popUpContextMenu(buildMenu()); }); } catch {}
  try { tray.on('right-click', () => { tray.popUpContextMenu(buildMenu()); }); } catch {}
  rebuildTray();
  try { if (needsOnboarding()) openInstallerWindow(); } catch {}
  try { setupAutoUpdates(); setTimeout(() => { try { if (autoUpdater) autoUpdater.checkForUpdatesAndNotify(); } catch{} }, 8000); } catch {}
  startScanning();
});

process.on('unhandledRejection', (err) => { log('Unhandled rejection', err && err.stack ? err.stack : String(err)); });
process.on('uncaughtException', (err) => { log('Uncaught exception', err && err.stack ? err.stack : String(err)); });
// Compute discovered character names from current state
function getDiscoveredCharacters(){
  try{
    const names = new Set();
    Object.keys(state.inventory||{}).forEach(n => names.add(String(n)));
    Object.values(state.latestZonesByFile||{}).forEach(v => { if (v && v.character) names.add(String(v.character)); });
    Object.keys(state.latestZones||{}).forEach(n => names.add(String(n)));
    return Array.from(names).sort();
  }catch{ return []; }
}
