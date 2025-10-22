const assert = require('assert');
const { MongoMemoryServer } = require('mongodb-memory-server');

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function run() {
  const mongod = await MongoMemoryServer.create();
  const uri = mongod.getUri('eqcm-scheduler-test');
  process.env.MONGODB_URI = uri;

  // Ensure we use a fresh cached client per test run.
  const clientPath = require.resolve('../server/db/client');
  delete require.cache[clientPath];

  const jobQueue = require('../server/jobs/job-queue');
  const dbModule = require('../server/db/client');
  const { getDb } = dbModule;

  try {
    await jobQueue.enqueueJob({
      spreadsheetId: 'sheet1',
      character: 'Testy',
      kind: 'unit-test',
      metadata: { from: 'test' },
      backoff: 25
    });

    const db = await getDb();
    const stored = await db.collection('sync_jobs').findOne({ character: 'Testy' });
    assert(stored, 'job is stored');
    assert.strictEqual(stored.status, 'queued');
    assert.strictEqual(stored.attempts, 0);

    const leased = await jobQueue.leaseNextJob();
    assert(leased, 'job leased');
    assert.strictEqual(leased.status, 'running');
    assert.strictEqual(leased.attempts, 1);

    const retrying = await jobQueue.markJobFailure(leased, new Error('boom'));
    assert.strictEqual(retrying, true, 'job should retry');

    // wait for retry delay to elapse
    await sleep(60);

    const leasedAgain = await jobQueue.leaseNextJob();
    assert(leasedAgain, 'job leased again');
    assert.strictEqual(String(leasedAgain._id), String(leased._id));
    assert.strictEqual(leasedAgain.attempts, 2);

    await jobQueue.markJobSuccess(leasedAgain);

    const final = await db.collection('sync_jobs').findOne({ _id: leased._id });
    assert(final, 'final job exists');
    assert.strictEqual(final.status, 'completed');
    assert.strictEqual(final.needs_sync, false);
  } finally {
    const db = await dbModule.getDb();
    if (db && db.client && typeof db.client.close === 'function') {
      await db.client.close();
    }
    await mongod.stop();
  }
}

run().then(
  () => {
    console.log('ok - job queue integration');
  },
  (err) => {
    console.error('not ok - job queue integration ->', err && err.stack ? err.stack : err);
    process.exitCode = 1;
  }
);
