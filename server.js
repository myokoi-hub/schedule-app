const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const os = require('os');
const cookieParser = require('cookie-parser');
const nodemailer = require('nodemailer');

const app = express();
const DATA_FILE = path.join(__dirname, 'data.json');

// ---- Azure AD / Graph API 設定 ----
const AZURE_TENANT_ID    = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_ID    = process.env.AZURE_CLIENT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback';
const GRAPH_SCOPES = 'Calendars.ReadWrite User.Read offline_access';
const isOAuthConfigured = () => !!(AZURE_TENANT_ID && AZURE_CLIENT_ID && AZURE_CLIENT_SECRET);

// ---- SMTP 設定 ----
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
const SMTP_FROM_NAME = process.env.SMTP_FROM_NAME || 'スケジュール調整';

// Microsoft Graph API でメール送信（client_credentials フロー・MFA対応）
async function getSmtpAccessToken() {
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    console.error('Graph token: 環境変数が不足しています (AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)');
    return null;
  }
  try {
    const params = new URLSearchParams({
      client_id:     AZURE_CLIENT_ID,
      client_secret: AZURE_CLIENT_SECRET,
      scope:         'https://graph.microsoft.com/.default',
      grant_type:    'client_credentials'
    });
    const res = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString()
    });
    const data = await res.json();
    if (data.error) {
      console.error('Graph token エラー:', data.error, data.error_description);
      return null;
    }
    console.log('Graph token 取得成功 (client_credentials)');
    return data.access_token || null;
  } catch (e) {
    console.error('Graph token 例外:', e.message);
    return null;
  }
}

async function sendMailViaGraph({ accessToken, from, recipients, subject, bodyText, icsContent }) {
  const toRecipients = recipients.map(r => ({ emailAddress: { address: r.email, name: r.name || r.email } }));
  const message = {
    subject,
    body: { contentType: 'Text', content: bodyText },
    toRecipients,
    ...(icsContent ? {
      attachments: [{
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'invite.ics',
        contentType: 'text/calendar; method=REQUEST',
        contentBytes: Buffer.from(icsContent).toString('base64')
      }]
    } : {})
  };
  console.log('Graph sendMail 送信先:', recipients.map(r => r.email).join(', '));
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(from)}/sendMail`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, saveToSentItems: false })
  });
  console.log('Graph sendMail ステータス:', res.status);
  if (!res.ok) {
    const err = await res.text();
    console.error('Graph sendMail エラー詳細:', err);
    throw new Error(`Graph sendMail failed: ${res.status} ${err}`);
  }
  console.log('Graph sendMail 成功');
}

function parseJpDateSrv(label) {
  const m  = label.match(/(\d+)月(\d+)日[^\s　]*[\s　]+(\d+):(\d+)/);
  const m2 = label.match(/(\d+)月(\d+)日/);
  if (!m2) return null;
  const month = parseInt(m ? m[1] : m2[1]);
  const day   = parseInt(m ? m[2] : m2[2]);
  const hour  = m ? parseInt(m[3]) : 10;
  const min   = m ? parseInt(m[4]) : 0;
  const now = new Date();
  let year = now.getFullYear();
  if (new Date(year, month-1, day) < new Date(now.getFullYear(), now.getMonth(), now.getDate())) year++;
  const p = n => String(n).padStart(2,'0');
  const fmt = d => `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
  return { start: fmt(new Date(year, month-1, day, hour, min)), end: fmt(new Date(year, month-1, day, hour+1, min)) };
}

function generateICS({ uid, title, location, description, dtstart, dtend, organizer, attendees }) {
  const dtstamp = new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d{3}/,'');
  const fmtLocal = s => s.replace(/[-:]/g,'').replace('T','T'); // "2026-06-15T14:00:00" → "20260615T140000"
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//ScheduleApp//JP',
    'CALSCALE:GREGORIAN',
    'METHOD:REQUEST',
    'BEGIN:VTIMEZONE',
    'TZID:Asia/Tokyo',
    'BEGIN:STANDARD',
    'DTSTART:19700101T000000',
    'TZOFFSETFROM:+0900',
    'TZOFFSETTO:+0900',
    'TZNAME:JST',
    'END:STANDARD',
    'END:VTIMEZONE',
    'BEGIN:VEVENT',
    `UID:${uid}`,
    `DTSTAMP:${dtstamp}`,
    `ORGANIZER;CN=${SMTP_FROM_NAME}:mailto:${organizer}`,
    `DTSTART;TZID=Asia/Tokyo:${fmtLocal(dtstart)}`,
    `DTEND;TZID=Asia/Tokyo:${fmtLocal(dtend)}`,
    `SUMMARY:${title.replace(/[,;\\]/g, s => '\\'+s)}`,
  ];
  if (location) lines.push(`LOCATION:${location.replace(/[,;\\]/g, s => '\\'+s)}`);
  if (description) lines.push(`DESCRIPTION:${description.replace(/\n/g,'\\n').replace(/[,;\\]/g, s => '\\'+s)}`);
  attendees.forEach(a => {
    lines.push(`ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${a.name}:mailto:${a.email}`);
  });
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines.join('\r\n');
}

// ---- PostgreSQL（本番環境用）----
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
  pgPool.query(`
    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      email TEXT NOT NULL,
      display_name TEXT,
      access_token TEXT NOT NULL,
      refresh_token TEXT,
      expires_at BIGINT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `).catch(err => console.error('DB sessions init error:', err));
}

// ---- イベントCRUD ----
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

// ---- 連絡先CRUD ----
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

// ---- セッション管理 ----
async function getSession(sessionId) {
  if (pgPool) {
    const r = await pgPool.query('SELECT * FROM sessions WHERE id = $1', [sessionId]);
    return r.rows[0] || null;
  }
  const data = readLocalData();
  return (data.sessions || {})[sessionId] || null;
}

async function saveSession(session) {
  if (pgPool) {
    await pgPool.query(
      `INSERT INTO sessions (id, email, display_name, access_token, refresh_token, expires_at)
       VALUES ($1, $2, $3, $4, $5, $6)
       ON CONFLICT (id) DO UPDATE SET
         access_token=$4, refresh_token=$5, expires_at=$6`,
      [session.id, session.email, session.display_name,
       session.access_token, session.refresh_token, session.expires_at]
    );
    return;
  }
  const data = readLocalData();
  if (!data.sessions) data.sessions = {};
  data.sessions[session.id] = session;
  writeLocalData(data);
}

async function deleteSession(sessionId) {
  if (pgPool) {
    await pgPool.query('DELETE FROM sessions WHERE id = $1', [sessionId]);
    return;
  }
  const data = readLocalData();
  if (data.sessions) delete data.sessions[sessionId];
  writeLocalData(data);
}

// ---- トークンリフレッシュ ----
async function getValidToken(sessionId) {
  const session = await getSession(sessionId);
  if (!session) return null;

  const expiresAt = Number(session.expires_at);
  // 5分余裕を持って判定
  if (Date.now() < expiresAt - 5 * 60 * 1000) {
    return session.access_token;
  }

  // リフレッシュトークンで更新
  if (!session.refresh_token) return null;
  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: AZURE_CLIENT_ID,
          client_secret: AZURE_CLIENT_SECRET,
          refresh_token: session.refresh_token,
          grant_type: 'refresh_token',
          scope: GRAPH_SCOPES
        })
      }
    );
    const tokens = await res.json();
    if (tokens.error) return null;

    await saveSession({
      ...session,
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token || session.refresh_token,
      expires_at: Date.now() + (tokens.expires_in || 3600) * 1000
    });
    return tokens.access_token;
  } catch {
    return null;
  }
}

// ---- Graph API: カレンダーイベント作成 ----
function parseJpDateForGraph(label) {
  const m  = label.match(/(\d+)月(\d+)日[^　\s]*[\s　]+(\d+):(\d+)/);
  const m2 = label.match(/(\d+)月(\d+)日/);
  if (!m2) return null;

  const month = parseInt(m ? m[1] : m2[1]);
  const day   = parseInt(m ? m[2] : m2[2]);
  const hour  = m ? parseInt(m[3]) : 10;
  const min   = m ? parseInt(m[4]) : 0;

  const now = new Date();
  let year = now.getFullYear();
  if (new Date(year, month-1, day) < new Date(now.getFullYear(), now.getMonth(), now.getDate())) year++;

  const pad = n => String(n).padStart(2, '0');
  const fmt = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;

  return {
    start: fmt(new Date(year, month-1, day, hour, min)),
    end:   fmt(new Date(year, month-1, day, hour+1, min))
  };
}

async function createCalendarEvent(event, dateId, accessToken) {
  try {
    const decidedDate = event.dates.find(d => d.id == dateId);
    if (!decidedDate) return false;

    const dateInfo = parseJpDateForGraph(decidedDate.label);
    if (!dateInfo) return false;

    const attendees = (event.invited || []).map(p => ({
      emailAddress: { address: p.email, name: p.name },
      type: 'required'
    }));

    const body = [
      event.description || '',
      '',
      `スケジュール調整ページ: ${process.env.APP_URL || ''}` +
        `/result.html?id=${event.id}`
    ].join('\n').trim();

    const graphEvent = {
      subject: event.title,
      body: { contentType: 'Text', content: body },
      start: { dateTime: dateInfo.start, timeZone: 'Asia/Tokyo' },
      end:   { dateTime: dateInfo.end,   timeZone: 'Asia/Tokyo' },
      attendees
    };

    const res = await fetch('https://graph.microsoft.com/v1.0/me/events', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(graphEvent)
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      console.error('Graph API error:', err);
    }
    return res.ok;
  } catch (err) {
    console.error('createCalendarEvent error:', err);
    return false;
  }
}

// ---- ローカルデータ ----
function readLocalData() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return { events: {}, contacts: [], sessions: {} }; }
}

function writeLocalData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// ---- ミドルウェア ----
app.use(express.json());
app.use(cookieParser());

// HTMLファイルを動的ルートで配信（Railwayキャッシュ回避）
const NO_CACHE = {
  'Cache-Control': 'no-cache, no-store, must-revalidate',
  'Pragma': 'no-cache',
  'Expires': '0'
};
const HTML_PAGES = {
  '/':              'index.html',
  '/index.html':    'index.html',
  '/event.html':    'event.html',
  '/result.html':   'result.html',
  '/events.html':   'events.html',
  '/contacts.html': 'contacts.html',
  '/sw.js':         'sw.js',
};
Object.entries(HTML_PAGES).forEach(([route, file]) => {
  app.get(route, (req, res) => {
    res.set(NO_CACHE);
    res.sendFile(path.join(__dirname, 'public', file));
  });
});

app.use(express.static(path.join(__dirname, 'public')));

// ---- 管理者認証 ----
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

// ======== OAuth 認証 ========

app.get('/auth/login', (req, res) => {
  if (!isOAuthConfigured()) return res.status(503).send('OAuth未設定');

  const state = crypto.randomBytes(8).toString('hex');
  const returnTo = req.query.returnTo || '/';
  res.cookie('oauth_state', `${state}|${returnTo}`, {
    httpOnly: true, maxAge: 10 * 60 * 1000, sameSite: 'lax'
  });

  const authUrl =
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${AZURE_CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&scope=${encodeURIComponent(GRAPH_SCOPES)}` +
    `&response_mode=query` +
    `&state=${state}`;

  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const { code, state, error } = req.query;

  if (error) {
    return res.redirect('/?auth_error=' + encodeURIComponent(req.query.error_description || error));
  }

  const stateCookie = req.cookies?.oauth_state || '';
  const [savedState, returnTo] = stateCookie.split('|');
  res.clearCookie('oauth_state');

  if (!state || state !== savedState) {
    return res.redirect('/?auth_error=invalid_state');
  }

  try {
    // code → token 交換
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: AZURE_CLIENT_ID,
          client_secret: AZURE_CLIENT_SECRET,
          code,
          redirect_uri: REDIRECT_URI,
          grant_type: 'authorization_code',
          scope: GRAPH_SCOPES
        })
      }
    );
    const tokens = await tokenRes.json();
    if (tokens.error) {
      return res.redirect('/?auth_error=' + encodeURIComponent(tokens.error_description || tokens.error));
    }

    // ユーザー情報取得
    const userRes = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${tokens.access_token}` }
    });
    const user = await userRes.json();

    // セッション保存
    const sessionId = crypto.randomBytes(16).toString('hex');
    await saveSession({
      id: sessionId,
      email: user.mail || user.userPrincipalName || '',
      display_name: user.displayName || '',
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token || null,
      expires_at: Date.now() + (tokens.expires_in || 3600) * 1000
    });

    res.cookie('schedule_session', sessionId, {
      httpOnly: true,
      maxAge: 30 * 24 * 60 * 60 * 1000,
      sameSite: 'lax'
    });

    res.redirect(returnTo || '/');
  } catch (err) {
    console.error('Auth callback error:', err);
    res.redirect('/?auth_error=server_error');
  }
});

app.get('/auth/me', async (req, res) => {
  if (!isOAuthConfigured()) return res.json({ loggedIn: false, oauthEnabled: false });
  const sessionId = req.cookies?.schedule_session;
  if (!sessionId) return res.json({ loggedIn: false, oauthEnabled: true });
  const session = await getSession(sessionId);
  if (!session) return res.json({ loggedIn: false, oauthEnabled: true });
  res.json({ loggedIn: true, oauthEnabled: true, email: session.email, name: session.display_name });
});

app.get('/auth/logout', async (req, res) => {
  const sessionId = req.cookies?.schedule_session;
  if (sessionId) await deleteSession(sessionId);
  res.clearCookie('schedule_session');
  res.redirect('/');
});

// ======== イベント API ========

app.get('/api/ping', (req, res) => {
  res.json({ version: 27, deployed: new Date().toISOString() });
});

// ---- GAL（全社アドレス帳）取得 ----
app.get('/api/gal-users', async (req, res) => {
  try {
    const token = await getSmtpAccessToken();
    if (!token) return res.status(503).json({ error: 'トークン取得失敗' });
    const usersRes = await fetch(
      'https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName&$top=999',
      { headers: { 'Authorization': `Bearer ${token}` } }
    );
    if (!usersRes.ok) {
      const err = await usersRes.text();
      return res.status(500).json({ error: `Graph API error: ${err}` });
    }
    const data = await usersRes.json();
    const users = (data.value || [])
      .map(u => ({ name: u.displayName || '', email: u.mail || u.userPrincipalName || '' }))
      .filter(u => u.email && u.email.includes('@') && !u.email.includes('#EXT#'))
      .sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    res.json(users);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/events', async (req, res) => {
  const events = await getEvents();
  const summary = events.map(e => {
    const responsers = [...new Set(e.responses.map(r => r.name))];
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
    invited_count: (event.invited || []).length,
    invited: (event.invited || [])
  });
});

app.put('/api/events/:id', async (req, res) => {
  const { title, description, deadline, dates } = req.body;
  if (!title) return res.status(400).json({ error: 'タイトルは必須です' });
  if (!deadline) return res.status(400).json({ error: '締切日は必須です' });
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).json({ error: 'イベントが見つかりません' });
  event.title = title;
  event.description = description || '';
  event.deadline = deadline;
  if (Array.isArray(dates)) {
    const oldDates = event.dates;
    let nextId = Math.max(0, ...oldDates.map(d => d.id)) + 1;
    event.dates = dates.map(label => {
      const existing = oldDates.find(d => d.label === label);
      return existing || { id: nextId++, label };
    });
    const validIds = new Set(event.dates.map(d => d.id));
    event.responses = event.responses.map(r => ({
      ...r,
      answers: Object.fromEntries(Object.entries(r.answers || {}).filter(([did]) => validIds.has(Number(did))))
    }));
    if (event.decided_date_id != null && !validIds.has(event.decided_date_id)) {
      event.decided_date_id = null;
    }
  }
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

  // ログイン中ならOutlookカレンダーに予定追加
  let calendarCreated = false;
  if (date_id != null && isOAuthConfigured()) {
    const sessionId = req.cookies?.schedule_session;
    if (sessionId) {
      const accessToken = await getValidToken(sessionId);
      if (accessToken) {
        calendarCreated = await createCalendarEvent(event, date_id, accessToken);
      }
    }
  }

  res.json({ success: true, calendarCreated });
});

app.get('/api/events/:id/calendar.ics', async (req, res) => {
  const event = await getEvent(req.params.id);
  if (!event) return res.status(404).send('Not found');
  if (event.decided_date_id == null) return res.status(404).send('確定日が設定されていません');
  const decidedDate = event.dates.find(d => d.id == event.decided_date_id);
  if (!decidedDate) return res.status(404).send('日付が見つかりません');

  const label = decidedDate.label;
  const m  = label.match(/(\d+)月(\d+)日[^　\s]*[\s　]+(\d+):(\d+)/);
  const m2 = label.match(/(\d+)月(\d+)日/);
  if (!m2) return res.status(400).send('日付の解析に失敗しました');

  const month = parseInt(m ? m[1] : m2[1]);
  const day   = parseInt(m ? m[2] : m2[2]);
  const hour  = m ? parseInt(m[3]) : 10;
  const min   = m ? parseInt(m[4]) : 0;

  const now = new Date();
  let year = now.getFullYear();
  if (new Date(year, month-1, day) < new Date(now.getFullYear(), now.getMonth(), now.getDate())) year++;

  const pad = n => String(n).padStart(2,'0');
  const fmtUtc = d => `${d.getUTCFullYear()}${pad(d.getUTCMonth()+1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}00Z`;
  const dtStart = new Date(Date.UTC(year, month-1, day, hour-9, min, 0));
  const dtEnd   = new Date(Date.UTC(year, month-1, day, hour-9+1, min, 0));
  const stamp = fmtUtc(new Date());

  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Schedule App//Schedule App//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    `UID:${event.id}-${event.decided_date_id}@schedule-app`,
    `DTSTAMP:${stamp}`,
    `DTSTART:${fmtUtc(dtStart)}`,
    `DTEND:${fmtUtc(dtEnd)}`,
    `SUMMARY:${event.title.replace(/[,;\\]/g, s => '\\' + s)}`,
    event.description ? `DESCRIPTION:${event.description.replace(/\n/g,'\\n')}` : null,
    'END:VEVENT',
    'END:VCALENDAR'
  ].filter(Boolean).join('\r\n');

  res.setHeader('Content-Type', 'text/calendar; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="event.ics"`);
  res.send(ics);
});

// ---- 日程調整招待メール送信 ----
app.post('/api/events/:id/send-invite', async (req, res) => {
  const { subject, body, recipients } = req.body;
  if (!subject || !Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({ error: '件名と送信先は必須です' });
  }
  if (!SMTP_USER) return res.status(503).json({ error: 'メール送信が未設定です' });
  try {
    const token = await getSmtpAccessToken();
    if (!token) return res.status(503).json({ error: 'メール認証に失敗しました。SMTP_USER/PASSを確認してください。' });
    await sendMailViaGraph({ accessToken: token, from: SMTP_USER, recipients, subject, bodyText: body || '' });
    res.json({ ok: true, sent: recipients.length });
  } catch (err) {
    console.error('Graph invite error:', err.message);
    res.status(500).json({ error: `送信失敗: ${err.message}` });
  }
});

// ---- 確定メール送信 ----
app.post('/api/events/:id/send-confirm', async (req, res) => {
  const { subject, location, memo, recipients, date_label } = req.body;
  if (!subject || !Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({ error: '件名と送信先は必須です' });
  }
  if (!SMTP_USER) return res.status(503).json({ error: 'メール送信が未設定です' });

  try {
    const token = await getSmtpAccessToken();
    if (!token) return res.status(503).json({ error: 'メール認証に失敗しました' });

    const dateTime = date_label ? parseJpDateSrv(date_label) : null;

    if (dateTime) {
      // Calendar Events API で正式な会議出席依頼を送信（Outlookに承諾/仮承諾/辞退ボタン表示）
      const attendees = recipients.map(r => ({
        emailAddress: { address: r.email, name: r.name || r.email },
        type: 'required'
      }));
      const graphEvent = {
        subject,
        body: { contentType: 'Text', content: memo || '' },
        start: { dateTime: dateTime.start, timeZone: 'Asia/Tokyo' },
        end:   { dateTime: dateTime.end,   timeZone: 'Asia/Tokyo' },
        attendees,
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness',
        ...(location ? { location: { displayName: location } } : {})
      };
      console.log('Graph createEvent 送信先:', recipients.map(r => r.email).join(', '));
      const evRes = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SMTP_USER)}/events`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(graphEvent)
      });
      console.log('Graph createEvent ステータス:', evRes.status);
      if (!evRes.ok) {
        const err = await evRes.text();
        console.error('Graph createEvent エラー:', err);
        throw new Error(`Graph createEvent failed: ${evRes.status} ${err}`);
      }
      console.log('Graph createEvent 成功（会議出席依頼送信済み）');
    } else {
      // 日時が解析できない場合は通常メール送信
      const bodyLines = [`【${subject}】の日程が確定しました。`, ''];
      if (date_label) bodyLines.push(`日時: ${date_label}`);
      if (location)   bodyLines.push(`場所: ${location}`);
      if (memo)       { bodyLines.push(''); bodyLines.push(memo); }
      await sendMailViaGraph({ accessToken: token, from: SMTP_USER, recipients, subject, bodyText: bodyLines.join('\n') });
    }

    res.json({ ok: true, sent: recipients.length });
  } catch (err) {
    console.error('Graph confirm error:', err.message);
    res.status(500).json({ error: `送信失敗: ${err.message}` });
  }
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

// ---- サーバー起動 ----
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
