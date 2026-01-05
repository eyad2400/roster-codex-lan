require('dotenv').config();

const express = require('express');
const cors = require('cors');
const path = require('path');

const app = express();
app.disable('x-powered-by');

/* =========================
   Middleware
   ========================= */
app.use(express.json({ limit: '10mb' }));

const configuredOrigins = (process.env.CORS_ORIGIN || '')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);
app.use(cors({
 origin: configuredOrigins.length ? configuredOrigins : false,
  credentials: configuredOrigins.length > 0
}));

app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
  res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
  next();
});

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
