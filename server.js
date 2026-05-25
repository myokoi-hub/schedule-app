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
  pgPool.query(`
    CREATE TABLE IF NOT EXISTS contacts (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      email TEXT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `).catch(err => console.error('DB contacts init error:', err));
}

async function getEvents() {
  if (pgPool) {
    const r = await pgPool.query('SELECT id, data, created_at FROM events ORDER BY created_at DESC');
    return r.rows.map(row => ({ ...row.data, db_created_at: row.created_at }));
  }
  const data = readLocalData();
  return Object.values(data.events).sort((a, b) =>
    new Date(b.created_at) - new Date(a.created_at)
  );
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
  catch { return { events: {}, contacts: [] }; }
}

async function getContacts() {
  if (pgPool) {
    const r = await pgPool.query('SELECT id, name, email FROM contacts ORDER BY name');
    return r.rows;
  }
  const data = readLocalData();
  return data.contacts || [];
}

async function addContact(contact) {
  if (pgPool) {
    await pgPool.query(
      'INSERT INTO contacts (id, name, email) VALUES ($1, $2, $3)',
      [contact.id, contact.name, contact.email]
    );
    return;
  }
  const data = readLocalData();
  if (!data.contacts) data.contacts = [];
  data.contacts.push(contact);
  data.contacts.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  writeLocalData(data);
}

async function updateContact(id, name, email) {
  if (pgPool) {
    await pgPool.query(
      'UPDATE contacts SET name = $1, email = $2 WHERE id = $3',
      [name, email, id]
    );
    return;
  }
  const data = readLocalData();
  const contact = (data.contacts || []).find(c => c.id === id);
  if (contact) {
    contact.name = name;
    contact.email = email;
    data.contacts.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    writeLocalData(data);
  }
}

async function removeContact(id) {
  if (pgPool) {
    await pgPool.query('DELETE FROM contacts WHERE id = $1', [id]);
    return;
  }
  const data = readLocalData();
  data.contacts = (data.contacts || []).filter(c => c.id !== id);
  writeLocalData(data);
}

function writeLocalData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/admin-auth', (req, res) => {
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword) return res.json({ ok: true });
  const { password } = req.body;
  if (password === adminPassword) {
    res.json({ ok: true });
  } else {
    res.status(401).json({ ok: false, error: 'パスワードが違います' });
  }
});

app.get('/api/events', async (req, res) => {
  const events = await getEvents();
  const summary = events.map(e => {
    const responsers = [...new Set(e.responses.map(r => r.name))];
    // 日程ごとの○数を集計
    const okCounts = {};
    e.dates.forEach(d => { okCounts[d.id] = 0; });
    e.responses.forEach(r => {
      Object.entries(r.answers || {}).forEach(([did, av]) => {
        if (av === '○' && okCounts[did] !== undefined) okCounts[did]++;
      });
    });
    const bestOk = Math.max(0, ...Object.values(okCounts));
    return {
      id: e.id,
      title: e.title,
      description: e.description || '',
      deadline: e.deadline || null,
      created_at: e.db_created_at || e.created_at,
      date_count: e.dates.length,
      response_count: responsers.length,
      invited_count: (e.invited || []).length,
      best_ok: bestOk,
    };
  });
  res.json(summary);
});

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
    deadline: req.body.deadline || null,
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
  const { participant_name, answers, comment } = req.body;
  if (!participant_name || !answers) {
    return res.status(400).json({ error: '必須項目が不足しています' });
  }

  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  event.responses = event.responses.filter(r => r.name !== participant_name);
  const answerMap = {};
  answers.forEach(({ date_id, availability }) => { answerMap[String(date_id)] = availability; });
  event.responses.push({ name: participant_name, answers: answerMap, comment: comment || '' });

  await saveEvent(event);
  res.json({ success: true });
});

app.get('/api/events/:id/results', async (req, res) => {
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  const dates = event.dates.map(d => ({ id: d.id, date_label: d.label }));
  const responses = [];
  const comments = {};
  event.responses.forEach(r => {
    comments[r.name] = r.comment || '';
    Object.entries(r.answers).forEach(([date_id, availability]) => {
      responses.push({ participant_name: r.name, date_id: Number(date_id), availability });
    });
  });

  res.json({
    id: event.id,
    title: event.title,
    description: event.description,
    deadline: event.deadline || null,
    decided_date_id: event.decided_date_id || null,
    dates,
    responses,
    comments,
    invited_count: (event.invited || []).length
  });
});

app.put('/api/events/:id', async (req, res) => {
  const { title, description, deadline } = req.body;
  if (!title) return res.status(400).json({ error: 'タイトルは必須です' });
  if (!deadline) return res.status(400).json({ error: '締切日は必須です' });
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });
  event.title = title;
  event.description = description || '';
  event.deadline = deadline;
  await saveEvent(event);
  res.json({ success: true });
});

app.delete('/api/events/:id', async (req, res) => {
  if (pgPool) {
    await pgPool.query('DELETE FROM events WHERE id = $1', [req.params.id]);
  } else {
    const data = readLocalData();
    delete data.events[req.params.id];
    writeLocalData(data);
  }
  res.json({ success: true });
});

app.post('/api/events/:id/decide', async (req, res) => {
  const { date_id } = req.body;
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });
  event.decided_date_id = date_id ?? null;
  await saveEvent(event);
  res.json({ success: true });
});

app.post('/api/events/:id/invited', async (req, res) => {
  const { contacts } = req.body;
  if (!Array.isArray(contacts)) return res.status(400).json({ error: '不正なデータです' });
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });

  const existing = event.invited || [];
  contacts.forEach(c => {
    if (!existing.some(e => e.id === c.id)) existing.push({ id: c.id, name: c.name, email: c.email });
  });
  event.invited = existing;
  await saveEvent(event);
  res.json({ success: true, invited_count: existing.length });
});

// ---- 連絡先API ----
app.get('/api/contacts', async (req, res) => {
  const contacts = await getContacts();
  res.json(contacts);
});

app.post('/api/contacts', async (req, res) => {
  const { name, email } = req.body;
  if (!name || !email) return res.status(400).json({ error: '名前とメールアドレスは必須です' });
  const contacts = await getContacts();
  if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase())) {
    return res.status(409).json({ error: 'そのメールアドレスはすでに登録されています' });
  }
  const contact = { id: crypto.randomBytes(4).toString('hex'), name, email };
  await addContact(contact);
  res.json(contact);
});

app.put('/api/contacts/:id', async (req, res) => {
  const { name, email } = req.body;
  if (!name || !email) return res.status(400).json({ error: '名前とメールアドレスは必須です' });
  const contacts = await getContacts();
  if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase() && c.id !== req.params.id)) {
    return res.status(409).json({ error: 'そのメールアドレスはすでに登録されています' });
  }
  await updateContact(req.params.id, name, email);
  res.json({ success: true });
});

app.delete('/api/contacts/:id', async (req, res) => {
  await removeContact(req.params.id);
  res.json({ success: true });
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
