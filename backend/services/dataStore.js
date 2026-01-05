const fs = require('fs/promises');
const path = require('path');

const DATA_DIR = path.join(__dirname, '..', 'data');
const STORE_PATH = path.join(DATA_DIR, 'rosterStore.json');
const DATA_FILES_DIR = path.join(DATA_DIR, 'roster');
const DEFAULT_BACKUP_FILES = [
  path.join(DATA_DIR, 'roster_backup_2026-01-03.json'),
  path.join(__dirname, '..', '..', 'roster_backup_2026-01-03.json'),
];

const DATA_FILE_MAP = {
  officers: 'officers.json',
  officerLimits: 'officerLimits.json',
  departments: 'departments.json',
  jobTitles: 'jobTitles.json',
  duties: 'duties.json',
  roster: 'roster.json',
  archivedRoster: 'archivedRoster.json',
  exceptions: 'exceptions.json',
  ranks: 'ranks.json',
  settings: 'settings.json',
  activityLog: 'activityLog.json',
  supportRequests: 'supportRequests.json',
  officerLimitFilters: 'officerLimitFilters.json',
  meta: 'meta.json',
};

const DEFAULT_DATA = {
  officers: [],
  officerLimits: {},
  departments: [],
  jobTitles: [],
  duties: [],
  roster: {},
  archivedRoster: {},
  exceptions: [],
  ranks: [],
  settings: {},
  activityLog: [],
  supportRequests: [],
 officerLimitFilters: {},
};
const SNAPSHOT_FILE_PATTERN = /^roster_backup_.+\.json$/i;

function buildSnapshotLabel(fileName) {
  if (fileName === 'rosterStore.json') {
    return 'النسخة الحالية على الخادم';
  }
  const match = fileName.match(/^roster_backup_(.+)\.json$/i);
  if (match) {
    return `نسخة احتياطية ${match[1]}`;
  }
  return fileName;
}
function hasData(payload) {
  if (!payload || typeof payload !== 'object') {
    return false;
  }
  const ignoredKeys = new Set(['meta']);
  return Object.keys(payload).some((key) => {
    if (ignoredKeys.has(key)) {
      return false;
    }
    const value = payload[key];
    if (Array.isArray(value)) {
      return value.length > 0;
    }
    if (value && typeof value === 'object') {
      return Object.keys(value).length > 0;
    }
    return false;
  });
}

async function readStore() {
  try {
    const raw = await fs.readFile(STORE_PATH, 'utf-8');
    const parsed = JSON.parse(raw || '{}');
    if (!parsed || typeof parsed !== 'object') {
      return { data: {}, updatedAt: null };
    }
    if (!('data' in parsed)) {
      if (!hasData(parsed)) {
        const fallback = await readDataFiles();
        if (hasData(fallback.data)) {
          return { data: fallback.data, updatedAt: fallback.updatedAt };
        }
        const hydrated = await ingestBackupFileIfPresent();
        if (hydrated) {
          return hydrated;
        }
      }
      return { data: parsed, updatedAt: null };
    }
    if (!hasData(parsed.data)) {
      const fallback = await readDataFiles();
      if (hasData(fallback.data)) {
        return { data: fallback.data, updatedAt: fallback.updatedAt };
      }
      const hydrated = await ingestBackupFileIfPresent();
      if (hydrated) {
        return hydrated;
      }
    }
    return parsed;
  } catch (error) {
    if (error.code === 'ENOENT') {
      const fallback = await readDataFiles();
      if (hasData(fallback.data)) {
        return { data: fallback.data, updatedAt: fallback.updatedAt };
      }
      const hydrated = await ingestBackupFileIfPresent();
      if (hydrated) {
        return hydrated;
      }
      return { data: {}, updatedAt: null };
    }
    throw error;
  }
}

async function readJsonFile(filePath, fallback) {
  try {
    const raw = await fs.readFile(filePath, 'utf-8');
    const parsed = JSON.parse(raw || 'null');
    return { value: parsed ?? fallback, found: true };
  } catch (error) {
    if (error.code === 'ENOENT') {
      return { value: fallback, found: false };
    }
    return { value: fallback, found: false };
  }
}

async function readDataFiles() {
  const data = {};
  let hasFiles = false;
  for (const [key, fileName] of Object.entries(DATA_FILE_MAP)) {
    const filePath = path.join(DATA_FILES_DIR, fileName);
    const fallback = key === 'meta' ? {} : DEFAULT_DATA[key] ?? null;
    const { value, found } = await readJsonFile(filePath, fallback);
    if (found) {
      hasFiles = true;
    }
    if (key === 'meta') {
      data.meta = value || {};
    } else {
      data[key] = value ?? fallback;
    }
  }
  const updatedAt = data.meta?.updatedAt || null;
  if (!hasFiles) {
    return { data: {}, updatedAt: null };
  }
  return { data, updatedAt };
}

async function areDataFilesPresent() {
  try {
    await fs.access(DATA_FILES_DIR);
  } catch (error) {
    return false;
  }
  const checks = await Promise.all(
    Object.values(DATA_FILE_MAP).map(async (fileName) => {
      try {
        await fs.access(path.join(DATA_FILES_DIR, fileName));
        return true;
      } catch (error) {
        return false;
      }
    })
  );
  return checks.every(Boolean);
}

async function writeDataFiles(payload, updatedAt) {
  await fs.mkdir(DATA_FILES_DIR, { recursive: true });
  const meta = Object.assign({}, payload?.meta || {});
  if (updatedAt) {
    meta.updatedAt = updatedAt;
  }
  const data = Object.assign({}, DEFAULT_DATA, payload || {}, { meta });
  const writes = [];
  for (const [key, fileName] of Object.entries(DATA_FILE_MAP)) {
    const filePath = path.join(DATA_FILES_DIR, fileName);
    const value = key === 'meta' ? meta : data[key];
    writes.push(fs.writeFile(filePath, JSON.stringify(value ?? null, null, 2), 'utf-8'));
  }
  await Promise.all(writes);
}

async function readBackupPayload() {
  for (const filePath of DEFAULT_BACKUP_FILES) {
    const { value, found } = await readJsonFile(filePath, null);
    if (found && value && typeof value === 'object') {
      return value;
    }
  }
  return null;
}

async function listSnapshotSummaries() {
  let dirents = [];
  try {
    dirents = await fs.readdir(DATA_DIR, { withFileTypes: true });
  } catch (error) {
    if (error.code === 'ENOENT') {
      return [];
    }
    throw error;
  }
  const entries = await Promise.all(
    dirents
      .filter((dirent) => dirent.isFile())
      .map(async (dirent) => {
        const fileName = dirent.name;
        if (fileName !== 'rosterStore.json' && !SNAPSHOT_FILE_PATTERN.test(fileName)) {
          return null;
        }
        const filePath = path.join(DATA_DIR, fileName);
        const { value } = await readJsonFile(filePath, null);
        if (!value || typeof value !== 'object') {
          return null;
        }
        const data = value.data && typeof value.data === 'object' ? value.data : value;
        return {
          id: fileName,
          label: buildSnapshotLabel(fileName),
          updatedAt: value.updatedAt || data?.meta?.updatedAt || null,
          hasData: hasData(data),
        };
      })
  );
  return entries
    .filter((entry) => entry && entry.hasData)
    .sort((a, b) => {
      const next = Date.parse(b.updatedAt || '') || 0;
      const current = Date.parse(a.updatedAt || '') || 0;
      if (next !== current) {
        return next - current;
      }
      return String(a.label).localeCompare(String(b.label), 'ar');
    });
}

async function readSnapshotById(snapshotId) {
  if (!snapshotId || typeof snapshotId !== 'string') {
    return null;
  }
  const snapshots = await listSnapshotSummaries();
  const target = snapshots.find((entry) => entry.id === snapshotId);
  if (!target) {
    return null;
  }
  const filePath = path.join(DATA_DIR, snapshotId);
  const { value } = await readJsonFile(filePath, null);
  if (!value || typeof value !== 'object') {
    return null;
  }
  const data = value.data && typeof value.data === 'object' ? value.data : value;
  return {
    data,
    updatedAt: value.updatedAt || data?.meta?.updatedAt || null,
  };
}

async function ingestBackupFileIfPresent() {
  const value = await readBackupPayload();
  if (!value) {
    return null;
  }
  const payload = value.data && typeof value.data === 'object' ? value.data : value;
  if (!hasData(payload)) {
    return null;
  }
  const updatedAt = value.updatedAt || payload?.meta?.updatedAt || new Date().toISOString();
  const store = { data: payload, updatedAt };
  await writeStore(store);
  return store;
}

async function writeStore(payload) {
  const updatedAt = payload?.updatedAt || payload?.data?.meta?.updatedAt || new Date().toISOString();
  await fs.mkdir(path.dirname(STORE_PATH), { recursive: true });
  await fs.writeFile(STORE_PATH, JSON.stringify(payload, null, 2), 'utf-8');
  const data = payload?.data ?? payload;
  await writeDataFiles(data, updatedAt);
  return payload;
}

async function ensureDataFilesFromStore(store) {
  const data = store?.data ?? store;
  if (!hasData(data)) {
    return;
  }
  const hasAllFiles = await areDataFilesPresent();
  if (!hasAllFiles) {
    const updatedAt = store?.updatedAt || data?.meta?.updatedAt || new Date().toISOString();
    await writeDataFiles(data, updatedAt);
  }
}

async function isStoreEmpty() {
  const store = await readStore();
  const data = store?.data ?? store;
  return !hasData(data);
}

module.exports = {
  hasData,
  ensureDataFilesFromStore,
  isStoreEmpty,
  listSnapshotSummaries,
  readStore,
  readSnapshotById,
  writeStore,
};
