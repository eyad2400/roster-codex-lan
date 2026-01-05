const express = require('express');

const { readStore, writeStore } = require('../services/dataStore');

const router = express.Router();

router.get('/', async (req, res) => {
  try {
    const store = await readStore();
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

module.exports = router;