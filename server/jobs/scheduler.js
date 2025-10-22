const { getDb } = require('../db/client');
const { listCharactersForSpreadsheet } = require('../db/summaries');
const { enqueueJob } = require('./job-queue');

const SCAN_INTERVAL_MS = Number(process.env.SYNC_SCAN_INTERVAL_MS || 60_000);
const STALE_MINUTES = Number(process.env.SYNC_THRESHOLD_MINUTES || 30);

async function scheduleCharacterSync(spreadsheetId, character, opts = {}) {
  await enqueueJob({
    spreadsheetId,
    character,
    kind: opts.kind || 'character-sync',
    delayMs: opts.delay || 0,
    maxAttempts: opts.attempts,
    backoff: opts.backoff,
    metadata: opts.metadata || {}
  });
}

async function pullCharactersNeedingSync(limit = 100) {
  const db = await getDb();
  const threshold = new Date(Date.now() - STALE_MINUTES * 60 * 1000);
  const cursor = db.collection('character_summaries')
    .find({
      $or: [
        { needs_sync: true },
        { last_sheet_push: { $exists: false } },
        { last_sheet_push: { $lt: threshold } }
      ]
    })
    .sort({ updated_at: -1 })
    .limit(limit);
  const results = [];
  await cursor.forEach((doc) => {
    if (!doc || !doc.character) return;
    results.push({
      spreadsheetId: doc.spreadsheet_id,
      character: doc.character
    });
  });
  return results;
}

async function enqueuePendingJobs() {
  const pending = await pullCharactersNeedingSync();
  if (!pending.length) return;
  await Promise.all(pending.map(({ spreadsheetId, character }) =>
    scheduleCharacterSync(spreadsheetId, character, { delay: randomDelay() })
  ));
}

function randomDelay() {
  const maxSpread = Number(process.env.SYNC_STAGGER_MS || 120_000);
  if (!Number.isFinite(maxSpread) || maxSpread <= 0) return 0;
  return Math.floor(Math.random() * maxSpread);
}

async function periodicallyScan() {
  try {
    await enqueuePendingJobs();
  } catch (err) {
    console.error('Scheduler scan error', err);
  } finally {
    setTimeout(periodicallyScan, SCAN_INTERVAL_MS).unref();
  }
}

async function startScheduler() {
  setImmediate(periodicallyScan);
}

async function enqueueFullBackfill(spreadsheetId) {
  const characters = await listCharactersForSpreadsheet(spreadsheetId);
  if (!characters.length) return 0;
  await Promise.all(characters.map((character) =>
    scheduleCharacterSync(spreadsheetId, character, { delay: randomDelay() })
  ));
  return characters.length;
}

module.exports = {
  startScheduler,
  scheduleCharacterSync,
  enqueueFullBackfill
};
