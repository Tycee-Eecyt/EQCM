const { getDb } = require('./client');

function nowDate() {
  return new Date();
}

async function rebuildSummaryEntry(db, spreadsheetId, character) {
  const [zone, faction, inventory, inventoryDetails] = await Promise.all([
    db.collection('zones').findOne({ spreadsheet_id: spreadsheetId, character }),
    db.collection('factions').findOne({ spreadsheet_id: spreadsheetId, character }),
    db.collection('inventory').findOne({ spreadsheet_id: spreadsheetId, character }),
    db.collection('inventory_details').findOne({ spreadsheet_id: spreadsheetId, character })
  ]);

  const summary = {
    spreadsheet_id: spreadsheetId,
    character,
    zone: zone ? {
      zone: zone.zone || '',
      utc: zone.zone_time_utc || '',
      local: zone.zone_time_local || '',
      tz: zone.tz || '',
      source: zone.source || ''
    } : null,
    faction: faction ? {
      standing: faction.standing || '',
      standingDisplay: faction.standing_display || '',
      score: faction.score ?? '',
      mob: faction.mob || '',
      utc: faction.consider_time_utc || '',
      local: faction.consider_time_local || ''
    } : null,
    inventory: inventory ? {
      file: inventory.file || '',
      logFile: inventory.log_file || '',
      created: inventory.created || '',
      modified: inventory.modified || '',
      raidKit: inventory.raid_kit || null,
      kitExtras: inventory.kit_extras || {}
    } : null,
    inventoryDetails: inventoryDetails ? {
      file: inventoryDetails.file || '',
      created: inventoryDetails.created || '',
      modified: inventoryDetails.modified || '',
      items: inventoryDetails.items || []
    } : null,
    needs_sync: true,
    updated_at: nowDate()
  };

  await db.collection('character_summaries').updateOne(
    { spreadsheet_id: spreadsheetId, character },
    {
      $set: summary,
      $setOnInsert: { created_at: nowDate() }
    },
    { upsert: true }
  );
}

async function updateCharacterSummary(spreadsheetId, character) {
  const db = await getDb();
  await rebuildSummaryEntry(db, spreadsheetId, character);
}

async function getCharacterSummary(spreadsheetId, character) {
  const db = await getDb();
  const summary = await db.collection('character_summaries').findOne({ spreadsheet_id: spreadsheetId, character });
  if (!summary) {
    await rebuildSummaryEntry(db, spreadsheetId, character);
    return db.collection('character_summaries').findOne({ spreadsheet_id: spreadsheetId, character });
  }
  return summary;
}

async function markCharacterSynced(spreadsheetId, character) {
  const db = await getDb();
  const now = nowDate();
  await Promise.all([
    db.collection('zones').updateMany(
      { spreadsheet_id: spreadsheetId, character },
      { $set: { needs_sync: false, last_sheet_push: now } }
    ),
    db.collection('factions').updateMany(
      { spreadsheet_id: spreadsheetId, character },
      { $set: { needs_sync: false, last_sheet_push: now } }
    ),
    db.collection('inventory').updateMany(
      { spreadsheet_id: spreadsheetId, character },
      { $set: { needs_sync: false, last_sheet_push: now } }
    ),
    db.collection('inventory_details').updateMany(
      { spreadsheet_id: spreadsheetId, character },
      { $set: { needs_sync: false, last_sheet_push: now } }
    ),
    db.collection('character_summaries').updateOne(
      { spreadsheet_id: spreadsheetId, character },
      { $set: { needs_sync: false, last_sheet_push: now } }
    ),
    db.collection('sync_jobs').updateMany(
      { spreadsheet_id: spreadsheetId, character, status: 'running' },
      { $set: { needs_sync: false, completed_at: now, status: 'completed', lease_expires: null, updated_at: now } }
    )
  ]);
}

async function listCharactersForSpreadsheet(spreadsheetId) {
  const db = await getDb();
  const cursor = db.collection('character_summaries').find({ spreadsheet_id: spreadsheetId }).project({ character: 1 });
  const chars = new Set();
  await cursor.forEach((doc) => {
    if (doc && doc.character) chars.add(String(doc.character));
  });
  return Array.from(chars);
}

module.exports = {
  updateCharacterSummary,
  getCharacterSummary,
  markCharacterSynced,
  listCharactersForSpreadsheet,
};
