const { leaseNextJob, markJobSuccess, markJobFailure } = require('./job-queue');
const { getCharacterSummary, markCharacterSynced } = require('../db/summaries');

const POLL_INTERVAL_MS = Number(process.env.SYNC_POLL_INTERVAL_MS || 5_000);

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

async function processJob(job, sheetHelpers) {
  const { spreadsheet_id: spreadsheetId, character } = job;
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
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function startWorker(sheetHelpers) {
  const concurrency = Math.max(1, Number(process.env.SYNC_WORKER_CONCURRENCY || 2));
  let stopRequested = false;

  async function workerLoop(id) {
    while (!stopRequested) {
      const job = await leaseNextJob();
      if (!job) {
        await sleep(POLL_INTERVAL_MS);
        continue;
      }
      try {
        await processJob(job, sheetHelpers);
        await markJobSuccess(job);
      } catch (err) {
        const willRetry = await markJobFailure(job, err);
        if (!willRetry) {
          console.error('Sync job permanently failed', job?._id?.toString(), err?.message || err);
        } else if (process.env.NODE_ENV !== 'test') {
          console.warn('Sync job failed, requeued', job?._id?.toString(), err?.message || err);
        }
      }
    }
  }

  const loops = Array.from({ length: concurrency }, (_, idx) => workerLoop(idx).catch((err) => {
    console.error('Worker loop crashed', err);
  }));

  return {
    async stop() {
      stopRequested = true;
      await Promise.allSettled(loops);
    }
  };
}

module.exports = {
  startWorker,
};
