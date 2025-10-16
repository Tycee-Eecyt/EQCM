const { Worker } = require('bullmq');
const { syncQueue, connectionOptions } = require('./scheduler');
const { getCharacterSummary, markCharacterSynced } = require('../db/summaries');

function buildUpsertsFromSummary(character, summary) {
  const upserts = {};
  if (summary.zone) {
    upserts.zones = [Object.assign({ character }, summary.zone)];
  }
  if (summary.faction) {
    upserts.factions = [Object.assign({ character }, {
      standingDisplay: summary.faction.standingDisplay || summary.faction.standing || ''
    }, summary.faction)];
  }
  if (summary.inventory) {
    upserts.inventory = [Object.assign({ character }, summary.inventory)];
  }
  if (summary.inventoryDetails) {
    upserts.inventoryDetails = [Object.assign({ character }, summary.inventoryDetails)];
  }
  return upserts;
}

function startWorker(sheetHelpers) {
  const concurrency = Number(process.env.SYNC_WORKER_CONCURRENCY || 2);
  const worker = new Worker(
    syncQueue.name,
    async (job) => {
      const { spreadsheetId, character } = job.data || {};
      if (!spreadsheetId || !character) return;
      const summary = await getCharacterSummary(spreadsheetId, character);
      if (!summary) return;
      const upserts = buildUpsertsFromSummary(character, summary);
      const sheets = await sheetHelpers.getSheetsClient();
      if (upserts.zones && upserts.zones.length) {
        await sheetHelpers.upsertZones(sheets, spreadsheetId, upserts.zones);
      }
      if (upserts.factions && upserts.factions.length) {
        await sheetHelpers.upsertFactions(sheets, spreadsheetId, upserts.factions);
      }
      if (upserts.inventory && upserts.inventory.length) {
        await sheetHelpers.upsertInventorySummary(sheets, spreadsheetId, upserts.inventory, {});
      }
      if (upserts.inventoryDetails && upserts.inventoryDetails.length) {
        await sheetHelpers.upsertInventoryDetails(sheets, spreadsheetId, upserts.inventoryDetails);
      }
      await markCharacterSynced(spreadsheetId, character);
    },
    Object.assign({}, connectionOptions, { concurrency })
  );

  worker.on('failed', (job, err) => {
    console.error('Sync worker failed', job?.id, err?.message || err);
  });

  return worker;
}

module.exports = {
  startWorker,
};
