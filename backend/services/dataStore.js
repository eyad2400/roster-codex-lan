const fs = require('fs/promises');
const path = require('path');

const STORE_PATH = path.join(__dirname, '..', 'data', 'rosterStore.json');

async function readStore() {
  try {
    const raw = await fs.readFile(STORE_PATH, 'utf-8');
    const parsed = JSON.parse(raw || '{}');
    if (!parsed || typeof parsed !== 'object') {
      return { data: {}, updatedAt: null };
    }
    if (!('data' in parsed)) {
      return { data: parsed, updatedAt: null };
    }
    return parsed;
  } catch (error) {
    if (error.code === 'ENOENT') {
      return { data: {}, updatedAt: null };
    }
    throw error;
  }
}

async function writeStore(payload) {
  await fs.mkdir(path.dirname(STORE_PATH), { recursive: true });
  await fs.writeFile(STORE_PATH, JSON.stringify(payload, null, 2), 'utf-8');
  return payload;
}

module.exports = {
  readStore,
  writeStore,
};