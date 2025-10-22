const { getDb } = require('../db/client');

const JOB_COLLECTION = 'sync_jobs';
const DEFAULT_MAX_ATTEMPTS = Number(process.env.SYNC_MAX_ATTEMPTS || 3);
const DEFAULT_BACKOFF_MS = Number(process.env.SYNC_BACKOFF_MS || 30_000);
const DEFAULT_BACKOFF_CAP_MS = Number(process.env.SYNC_BACKOFF_CAP_MS || 15 * 60 * 1000);
const JOB_LEASE_MS = Number(process.env.SYNC_LEASE_MS || 5 * 60 * 1000);

function now() {
  return new Date();
}

function normalizeDelayMs(delay) {
  const asNumber = Number(delay);
  if (!Number.isFinite(asNumber) || asNumber <= 0) return 0;
  return Math.floor(asNumber);
}

function normalizeBackoff(backoff) {
  if (typeof backoff === 'number') {
    return backoff > 0 ? Math.floor(backoff) : DEFAULT_BACKOFF_MS;
  }
  if (backoff && typeof backoff === 'object' && typeof backoff.delay === 'number') {
    return backoff.delay > 0 ? Math.floor(backoff.delay) : DEFAULT_BACKOFF_MS;
  }
  return DEFAULT_BACKOFF_MS;
}

function formatError(err) {
  if (!err) return null;
  if (err instanceof Error) {
    return `${err.message}${err.stack ? `\n${err.stack}` : ''}`;
  }
  if (typeof err === 'string') return err;
  try {
    return JSON.stringify(err);
  } catch (_) {
    return String(err);
  }
}

function computeRetryDelay(job) {
  const base = Number(job.backoff_ms || DEFAULT_BACKOFF_MS);
  const attempt = Math.max(1, Number(job.attempts || 1));
  const exponent = Math.max(0, attempt - 1);
  const delay = base * Math.pow(2, exponent);
  const cap = Number.isFinite(Number(process.env.SYNC_BACKOFF_CAP_MS))
    ? Number(process.env.SYNC_BACKOFF_CAP_MS)
    : DEFAULT_BACKOFF_CAP_MS;
  return Math.min(cap, delay);
}

async function getJobsCollection() {
  const db = await getDb();
  return db.collection(JOB_COLLECTION);
}

async function enqueueJob({
  spreadsheetId,
  character,
  kind = 'character-sync',
  delayMs = 0,
  maxAttempts = DEFAULT_MAX_ATTEMPTS,
  backoff = DEFAULT_BACKOFF_MS,
  metadata = {}
}) {
  if (!spreadsheetId || !character) return null;
  const jobs = await getJobsCollection();
  const nowDate = now();
  const scheduledFor = new Date(nowDate.getTime() + normalizeDelayMs(delayMs));
  const backoffMs = normalizeBackoff(backoff);
  const maxAttemptsInt = Math.max(1, Math.floor(Number(maxAttempts) || DEFAULT_MAX_ATTEMPTS));

  const update = await jobs.updateOne(
    {
      spreadsheet_id: spreadsheetId,
      character,
      status: 'queued'
    },
    {
      $set: {
        spreadsheet_id: spreadsheetId,
        character,
        kind,
        needs_sync: true,
        status: 'queued',
        scheduled_for: scheduledFor,
        backoff_ms: backoffMs,
        max_attempts: maxAttemptsInt,
        metadata,
        updated_at: nowDate
      },
      $setOnInsert: {
        attempts: 0,
        created_at: nowDate
      }
    },
    { upsert: true }
  );

  if (update.matchedCount || update.upsertedCount) {
    return update.upsertedId ? update.upsertedId._id || update.upsertedId : null;
  }

  const inserted = await jobs.insertOne({
    spreadsheet_id: spreadsheetId,
    character,
    kind,
    needs_sync: true,
    status: 'queued',
    scheduled_for: scheduledFor,
    backoff_ms: backoffMs,
    max_attempts: maxAttemptsInt,
    metadata,
    attempts: 0,
    created_at: nowDate,
    updated_at: nowDate
  });
  return inserted.insertedId;
}

async function leaseNextJob() {
  const jobs = await getJobsCollection();
  const nowDate = now();
  const leaseExpires = new Date(nowDate.getTime() + JOB_LEASE_MS);
  const unwrap = (result) => {
    if (!result) return null;
    if (Object.prototype.hasOwnProperty.call(result, 'value')) {
      return result.value;
    }
    return result;
  };

  const pending = await jobs.findOneAndUpdate(
    {
      status: 'queued',
      scheduled_for: { $lte: nowDate },
      $expr: { $lt: ['$attempts', '$max_attempts'] }
    },
    {
      $set: {
        status: 'running',
        lease_expires: leaseExpires,
        last_attempt_started_at: nowDate,
        updated_at: nowDate
      },
      $inc: { attempts: 1 }
    },
    {
      sort: { priority: -1, scheduled_for: 1, created_at: 1 },
      returnDocument: 'after'
    }
  );
  const pendingDoc = unwrap(pending);
  if (pendingDoc) return pendingDoc;

  const stale = await jobs.findOneAndUpdate(
    {
      status: 'running',
      lease_expires: { $lte: nowDate },
      $expr: { $lt: ['$attempts', '$max_attempts'] }
    },
    {
      $set: {
        status: 'running',
        lease_expires: leaseExpires,
        last_attempt_started_at: nowDate,
        updated_at: nowDate
      },
      $inc: { attempts: 1 }
    },
    {
      sort: { lease_expires: 1, created_at: 1 },
      returnDocument: 'after'
    }
  );
  return unwrap(stale);
}

async function markJobSuccess(job) {
  if (!job || !job._id) return;
  const jobs = await getJobsCollection();
  const nowDate = now();
  await jobs.updateOne(
    { _id: job._id },
    {
      $set: {
        status: 'completed',
        needs_sync: false,
        completed_at: nowDate,
        updated_at: nowDate,
        lease_expires: null,
        last_error: null
      }
    }
  );
}

async function markJobFailure(job, error) {
  if (!job || !job._id) return false;
  const jobs = await getJobsCollection();
  const nowDate = now();
  const maxAttempts = Number(job.max_attempts || DEFAULT_MAX_ATTEMPTS);
  const attempts = Number(job.attempts || 0);

  if (attempts >= maxAttempts) {
    await jobs.updateOne(
      { _id: job._id },
      {
        $set: {
          status: 'failed',
          needs_sync: false,
          failed_at: nowDate,
          updated_at: nowDate,
          lease_expires: null,
          last_error: formatError(error)
        }
      }
    );
    return false;
  }

  const retryDelay = computeRetryDelay(job);
  await jobs.updateOne(
    { _id: job._id },
    {
      $set: {
        status: 'queued',
        needs_sync: true,
        scheduled_for: new Date(nowDate.getTime() + retryDelay),
        lease_expires: null,
        updated_at: nowDate,
        last_error: formatError(error)
      }
    }
  );
  return true;
}

module.exports = {
  enqueueJob,
  leaseNextJob,
  markJobSuccess,
  markJobFailure,
  normalizeDelayMs,
  normalizeBackoff,
  DEFAULT_MAX_ATTEMPTS,
  DEFAULT_BACKOFF_MS,
  JOB_COLLECTION
};
