require('dotenv').config();

const express = require('express');
const cors = require('cors');

const app = express();

/* =========================
   Middleware
   ========================= */
app.use(express.json());

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
// app.use('/api/roster', require('./routes/roster.routes'));

/* =========================
   Server start
   ========================= */
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '197.2.0.241';

app.listen(PORT, HOST, () => {
  console.log(`RosterCodex backend running on http://${HOST}:${PORT}`);
});
