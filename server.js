const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const os = require('os');

const app = express();
const DATA_FILE = path.join(__dirname, 'data.json');

// PostgreSQL（本番環境用）
let pgPool = null;
if (process.env.DATABASE_URL) {
  const { Pool } = require('pg');
  pgPool = new Pool({ connectionString: process.env.DATABASE_URL, ssl: { rejectUnauthorized: false } });
  pgPool.query(`
    CREATE TABLE IF NOT EXISTS events (
      id TEXT PRIMARY KEY,
      data JSONB NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `).catch(err => console.error('DB init error:', err));
}

async function getEvent(id) {
  if (pgPool) {
    const r = await pgPool.query('SELECT data FROM events WHERE id = $1', [id]);
    return r.rows[0]?.data || null;
  }
  return readLocalData().events[id] || null;
}

async function saveEvent(event) {
  if (pgPool) {
    await pgPool.query(
      'INSERT INTO events (id, data) VALUES ($1, $2) ON CONFLICT (id) DO UPDATE SET data = $2',
      [event.id, JSON.stringify(event)]
    );
    return;
  }
  const data = readLocalData();
  data.events[event.id] = event;
  writeLocalData(data);
}

function readLocalData() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return { events: {} }; }
}

function writeLocalData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/events', async (req, res) => {
  const { title, description, dates } = req.body;
  if (!title || !dates || dates.length === 0) {
    return res.status(400).json({ error: '必須項目が不足しています' });
  }

  const id = crypto.randomBytes(4).toString('hex');
  const event = {
    id,
    title,
    description: description || '',
    dates: dates.map((label, i) => ({ id: i + 1, label })),
    responses: [],
    created_at: new Date().toISOString()
  };

  await saveEvent(event);
  res.json({ id });
});

app.get('/api/events/:id', async (req, res) => {
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  res.json({
    id: event.id,
    title: event.title,
    description: event.description,
    dates: event.dates.map(d => ({ id: d.id, date_label: d.label }))
  });
});

app.post('/api/events/:id/responses', async (req, res) => {
  const { participant_name, answers } = req.body;
  if (!participant_name || !answers) {
    return res.status(400).json({ error: '必須項目が不足しています' });
  }

  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  event.responses = event.responses.filter(r => r.name !== participant_name);
  const answerMap = {};
  answers.forEach(({ date_id, availability }) => { answerMap[String(date_id)] = availability; });
  event.responses.push({ name: participant_name, answers: answerMap });

  await saveEvent(event);
  res.json({ success: true });
});

app.get('/api/events/:id/results', async (req, res) => {
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  const dates = event.dates.map(d => ({ id: d.id, date_label: d.label }));
  const responses = [];
  event.responses.forEach(r => {
    Object.entries(r.answers).forEach(([date_id, availability]) => {
      responses.push({ participant_name: r.name, date_id: Number(date_id), availability });
    });
  });

  res.json({ id: event.id, title: event.title, description: event.description, dates, responses });
});

function getLocalIP() {
  for (const ifaces of Object.values(os.networkInterfaces())) {
    for (const iface of ifaces) {
      if (iface.family === 'IPv4' && !iface.internal) return iface.address;
    }
  }
  return 'localhost';
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  if (process.env.DATABASE_URL) {
    console.log(`\n✅ スケジュール調整アプリ起動中 (port ${PORT})\n`);
  } else {
    const ip = getLocalIP();
    console.log(`\n✅ スケジュール調整アプリ起動中`);
    console.log(`   PC から開く:     http://localhost:${PORT}`);
    console.log(`   スマホから開く:  http://${ip}:${PORT}`);
    console.log(`\n   ※ PC とスマホが同じ Wi-Fi に繋がっている必要があります\n`);
  }
});
