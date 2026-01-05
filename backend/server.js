require('dotenv').config();

const express = require('express');
const cors = require('cors');
const path = require('path');

const app = express();

/* =========================
   Middleware
   ========================= */
app.use(express.json({ limit: '10mb' }));

app.use(cors({
  origin: process.env.CORS_ORIGIN || '*',
  credentials: true
}));

/* =========================
   Health check (IMPORTANT)
   ========================= */
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    service: 'roster-codex-backend',
    time: new Date().toISOString()
  });
});

/* =========================
   API routes (add later)
   ========================= */
// app.use('/api/auth', require('./routes/auth.routes'));
app.use('/api/data', require('./routes/data.routes'));
app.get('/vendor/html2canvas.min.js', (req, res) => {
  res.sendFile(path.join(__dirname, 'vendor', 'html2canvas.min.js'));
});

/* =========================
   Server start
   ========================= */
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '197.2.0.241';

app.listen(PORT, HOST, () => {
  console.log(`RosterCodex backend running on http://${HOST}:${PORT}`);
});
