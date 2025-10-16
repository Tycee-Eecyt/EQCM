const { MongoClient } = require('mongodb');

const DEFAULT_URI = 'mongodb://localhost:27017/eqcm';
let clientPromise = null;

function getMongoUri() {
  const uri = process.env.MONGODB_URI || DEFAULT_URI;
  return uri;
}

async function getClient() {
  if (!clientPromise) {
    const uri = getMongoUri();
    const client = new MongoClient(uri, {
      maxPoolSize: Number(process.env.MONGODB_POOL_SIZE || 10),
      serverSelectionTimeoutMS: 5000
    });
    clientPromise = client.connect().catch((err) => {
      clientPromise = null;
      throw err;
    });
  }
  return clientPromise;
}

async function getDb() {
  const client = await getClient();
  const uri = getMongoUri();
  // Extract DB name from URI or fall back to eqcm
  try {
    const parsed = new URL(uri);
    const pathname = parsed.pathname.replace(/^\//, '');
    if (pathname) return client.db(pathname);
  } catch (_) { /* ignore */ }
  return client.db('eqcm');
}

module.exports = {
  getDb,
};
