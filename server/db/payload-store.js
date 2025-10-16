const { ObjectId } = require('mongodb');
const { getDb } = require('./client');

let indexesEnsured = false;

async function ensureIndexes() {
  if (indexesEnsured) return;
  const db = await getDb();
  await Promise.all([
    db.collection('zones').createIndex({ character: 1 }, { background: true }),
    db.collection('zones').createIndex({ needs_sync: 1, last_sheet_push: 1 }, { background: true }),
    db.collection('factions').createIndex({ character: 1 }, { background: true }),
    db.collection('factions').createIndex({ needs_sync: 1, last_sheet_push: 1 }, { background: true }),
    db.collection('inventory').createIndex({ character: 1 }, { background: true }),
    db.collection('inventory').createIndex({ needs_sync: 1, last_sheet_push: 1 }, { background: true }),
    db.collection('inventory_details').createIndex({ character: 1 }, { background: true }),
    db.collection('inventory_details').createIndex({ needs_sync: 1, last_sheet_push: 1 }, { background: true }),
    db.collection('sync_jobs').createIndex({ needs_sync: 1, scheduled_for: 1 }, { background: true })
  ]);
  indexesEnsured = true;
}

function nowDate() {
  return new Date();
}

function asDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

async function upsertZone(db, spreadsheetId, row) {
  const character = String(row.character || '').trim();
  if (!character) return;
  await db.collection('zones').updateOne(
    { spreadsheet_id: spreadsheetId, character },
    {
      $set: {
        spreadsheet_id: spreadsheetId,
        character,
        zone: row.zone || '',
        zone_time_utc: row.utc || '',
        zone_time_local: row.local || '',
        tz: row.tz || '',
        source: row.source || '',
        needs_sync: true,
        updated_at: nowDate()
      },
      $setOnInsert: {
        created_at: nowDate()
      }
    },
    { upsert: true }
  );
}

async function upsertFaction(db, spreadsheetId, row) {
  const character = String(row.character || '').trim();
  if (!character) return;
  await db.collection('factions').updateOne(
    { spreadsheet_id: spreadsheetId, character },
    {
      $set: {
        spreadsheet_id: spreadsheetId,
        character,
        standing: row.standing || '',
        standing_display: row.standingDisplay || '',
        score: row.score ?? '',
        mob: row.mob || '',
        consider_time_utc: row.utc || '',
        consider_time_local: row.local || '',
        needs_sync: true,
        updated_at: nowDate()
      },
      $setOnInsert: { created_at: nowDate() }
    },
    { upsert: true }
  );
}

async function upsertInventorySummary(db, spreadsheetId, row) {
  const character = String(row.character || '').trim();
  if (!character) return;
  await db.collection('inventory').updateOne(
    { spreadsheet_id: spreadsheetId, character },
    {
      $set: {
        spreadsheet_id: spreadsheetId,
        character,
        file: row.file || '',
        log_file: row.logFile || '',
        created: row.created || '',
        modified: row.modified || '',
        raid_kit: row.raidKit || null,
        kit_extras: row.kitExtras || {},
        needs_sync: true,
        updated_at: nowDate()
      },
      $setOnInsert: { created_at: nowDate() }
    },
    { upsert: true }
  );
}

async function upsertInventoryDetails(db, spreadsheetId, row) {
  const character = String(row.character || '').trim();
  if (!character) return;
  const items = Array.isArray(row.items) ? row.items : [];
  await db.collection('inventory_details').updateOne(
    { spreadsheet_id: spreadsheetId, character, file: row.file || '' },
    {
      $set: {
        spreadsheet_id: spreadsheetId,
        character,
        file: row.file || '',
        created: row.created || '',
        modified: row.modified || '',
        items,
        needs_sync: true,
        updated_at: nowDate()
      },
      $setOnInsert: { created_at: nowDate() }
    },
    { upsert: true }
  );
}

async function enqueueSyncJob(db, spreadsheetId, kind, character) {
  await db.collection('sync_jobs').insertOne({
    spreadsheet_id: spreadsheetId,
    kind,
    character,
    needs_sync: true,
    scheduled_for: nowDate(),
    created_at: nowDate()
  });
}

async function storeWebhookPayload(spreadsheetId, upserts = {}, meta = {}) {
  await ensureIndexes();
  const db = await getDb();
  const ops = [];
  if (Array.isArray(upserts.zones)) {
    upserts.zones.forEach((row) => ops.push(upsertZone(db, spreadsheetId, row)));
  }
  if (Array.isArray(upserts.factions)) {
    upserts.factions.forEach((row) => ops.push(upsertFaction(db, spreadsheetId, row)));
  }
  if (Array.isArray(upserts.inventory)) {
    upserts.inventory.forEach((row) => ops.push(upsertInventorySummary(db, spreadsheetId, row)));
  }
  if (Array.isArray(upserts.inventoryDetails)) {
    upserts.inventoryDetails.forEach((row) => ops.push(upsertInventoryDetails(db, spreadsheetId, row)));
  }
  await Promise.all(ops);

  const characters = new Set();
  (upserts.zones || []).forEach((r) => { if (r.character) characters.add(r.character); });
  (upserts.factions || []).forEach((r) => { if (r.character) characters.add(r.character); });
  (upserts.inventory || []).forEach((r) => { if (r.character) characters.add(r.character); });
  (upserts.inventoryDetails || []).forEach((r) => { if (r.character) characters.add(r.character); });

  await Promise.all(
    Array.from(characters).map((character) =>
      enqueueSyncJob(db, spreadsheetId, 'webhook_upsert', character)
    )
  );
}

async function recordRawWebhook(spreadsheetId, body) {
  await ensureIndexes();
  const db = await getDb();
  await db.collection('webhook_events').insertOne({
    spreadsheet_id: spreadsheetId,
    payload: body,
    created_at: nowDate()
  });
}

module.exports = {
  storeWebhookPayload,
  recordRawWebhook,
};
