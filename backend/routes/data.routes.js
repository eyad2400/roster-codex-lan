const {
  readStore,
  writeStore,
  hasData,
  ensureDataFilesFromStore,
  listSnapshotSummaries,
  readSnapshotById,
} = require('../services/dataStore');
const router = express.Router();

router.get('/snapshots', async (_req, res) => {
  try {
    const snapshots = await listSnapshotSummaries();
    res.json({ snapshots });
  } catch (error) {
    console.error('Failed to list roster snapshots', error);
    res.status(500).json({ error: 'Failed to list roster snapshots.' });
  }
});

router.get('/snapshots/:id', async (req, res) => {
  try {
    const snapshot = await readSnapshotById(req.params.id);
    if (!snapshot) {
      return res.status(404).json({ error: 'Snapshot not found.' });
    }
    return res.json(snapshot);
  } catch (error) {
    console.error('Failed to read roster snapshot', error);
    return res.status(500).json({ error: 'Failed to read roster snapshot.' });
  }
});

router.get('/', async (req, res) => {
  try {
    const store = await readStore();
    await ensureDataFilesFromStore(store);
    res.json(store);
  } catch (error) {
    console.error('Failed to read roster store', error);
    res.status(500).json({ error: 'Failed to read roster store.' });
  }
});

router.post('/', async (req, res) => {
  try {
    const body = req.body;
    if (!body || typeof body !== 'object') {
      return res.status(400).json({ error: 'Invalid payload.' });
    }
    const data = body.data ?? body;
    const store = {
      data,
      updatedAt: new Date().toISOString(),
    };
       await writeStore(store);
    return res.json(store);
  } catch (error) {
    console.error('Failed to write roster store', error);
    return res.status(500).json({ error: 'Failed to write roster store.' });
  }
});

router.post('/import', async (req, res) => {
  try {
    const body = req.body;
    if (!body || typeof body !== 'object') {
      return res.status(400).json({ error: 'Invalid payload.' });
    }
    if (req.get('x-admin-import') !== 'true') {
      return res.status(403).json({ error: 'Admin import required.' });
    }
    const payload = body.data ?? body;
    if (!payload || typeof payload !== 'object') {
      return res.status(400).json({ error: 'Invalid backup payload.' });
    }
    if (!hasData(payload)) {
      return res.status(400).json({ error: 'Backup payload is empty.' });
    }
    const updatedAt = body.updatedAt || payload?.meta?.updatedAt || new Date().toISOString();
    const store = {
      data: payload,
      updatedAt,
    };
    await writeStore(store);
    return res.json(store);
  } catch (error) {
    console.error('Failed to import roster store', error);
    return res.status(500).json({ error: 'Failed to import roster store.' });
  }
});

module.exports = router;
