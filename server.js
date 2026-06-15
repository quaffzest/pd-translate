const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const session = require('express-session');
const passport = require('passport');
const GoogleStrategy = require('passport-google-oauth20').Strategy;
const googleConfig = require('./config/google');
const driveService = require('./services/driveService');

const PORT = process.env.PORT || 3000;
console.log('Starting pd-translate, GOOGLE_CLIENT_ID set:', !!process.env.GOOGLE_CLIENT_ID);
const DATA_DIR = process.env.DATA_DIR || (fs.existsSync('/var/data') ? '/var/data' : path.join(__dirname, 'data'));
const LEGACY_FILE = path.join(DATA_DIR, '第01话_协作主表.xlsx');
const WORKBOOK_DIR = path.join(DATA_DIR, 'workbooks');
const META_FILE = path.join(WORKBOOK_DIR, '_meta.json');

const workbooks = new Map();
const clients = new Map();
const saveTimers = new Map();
let meta = {};

function ensureSeedWorkbook() {
  fs.mkdirSync(WORKBOOK_DIR, { recursive: true });
  const existing = walkFiles(WORKBOOK_DIR).filter(file => /\.xlsx?$/i.test(file));
  if (!existing.length && fs.existsSync(LEGACY_FILE)) {
    fs.copyFileSync(LEGACY_FILE, path.join(WORKBOOK_DIR, 'default.xlsx'));
  }
}

function readMeta() {
  try {
    meta = JSON.parse(fs.readFileSync(META_FILE, 'utf8'));
  } catch (e) {
    meta = {};
  }
}

function writeMeta() {
  fs.writeFileSync(META_FILE, JSON.stringify(meta, null, 2));
}

function walkFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === '_meta.json') continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...walkFiles(full));
    else out.push(full);
  }
  return out;
}

function normalizeFolder(folder) {
  const clean = String(folder || '').replace(/\\/g, '/').split('/').filter(Boolean).join('/');
  if (clean.includes('..')) throw new Error('invalid_folder');
  return clean;
}

function resolveFolder(folder) {
  const clean = normalizeFolder(folder);
  const full = path.resolve(WORKBOOK_DIR, clean);
  if (full !== WORKBOOK_DIR && !full.startsWith(WORKBOOK_DIR + path.sep)) throw new Error('invalid_folder');
  return { clean, full };
}

function relativeFolder(filePath) {
  const rel = path.relative(WORKBOOK_DIR, path.dirname(filePath)).replace(/\\/g, '/');
  return rel === '.' ? '' : rel;
}

function safeName(name) {
  return String(name || '协作主表.xlsx').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
}

function makeId(name) {
  const base = safeName(name).replace(/\.[^.]+$/, '').replace(/[^\w\u4e00-\u9fa5-]+/g, '_').slice(0, 40) || 'workbook';
  return Date.now().toString(36) + '_' + base;
}

function parseWorkbookBuffer(buffer) {
  return parseWorkbook(XLSX.read(buffer, { type: 'buffer' }));
}

function parseWorkbookFile(filePath) {
  return parseWorkbook(XLSX.readFile(filePath));
}

function parseWorkbook(wb) {
  const sheets = {};
  const sheetNames = [];
  let currentSheet = '';

  for (const sn of wb.SheetNames) {
    if (sn === '使用说明·流程' || sn === '说明') continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
    if (rows.length < 2) continue;

    const h = rows[0];
    const dr = rows.slice(1);
    const si = Math.max(h.findIndex(x => String(x).trim() === '序'), 0);
    const ki = Math.max(h.findIndex(x => String(x).trim() === '韩文原文'), 0);
    const ti = Math.max(h.findIndex(x => String(x).trim() === '译者'), 0);
    const mi = h.findIndex(x => String(x).trim() === '合并');
    const headers = [];

    for (let ci = ti; ci < h.length; ci++) {
      const hn = String(h[ci]).trim();
      if (!hn || hn === '序' || hn === '韩文原文' || hn === '合并' || hn === '终稿') continue;
      headers.push({ name: hn, idx: ci });
    }

    const sheetRows = dr.map(r => {
      const content = headers.map(c => String(r[c.idx] || '').trim());
      const mergeValue = mi >= 0 ? String(r[mi] || '').trim().toUpperCase() : '';
      return {
        seq: String(r[si] || '').trim(),
        kr: String(r[ki] || '').trim(),
        content,
        styles: { seq: {}, kr: {}, content: headers.map(() => ({})) },
        merge: mergeValue === 'Y' || mergeValue === 'TRUE' || mergeValue === '1',
      };
    });

    if (sheetRows.length) {
      sheets[sn] = { headers, rows: sheetRows };
      sheetNames.push(sn);
    }
  }

  if (sheetNames.length) currentSheet = sheetNames[0];
  return { sheets, sheetNames, currentSheet };
}

function writeWorkbook(wbInfo) {
  const wb = XLSX.utils.book_new();
  for (const sn of wbInfo.state.sheetNames) {
    const info = wbInfo.state.sheets[sn];
    const header = ['序', '韩文原文', ...info.headers.map(c => c.name), '终稿', '合并'];
    const data = [header];
    for (const row of info.rows) {
      const values = [row.seq, row.kr];
      for (let ci = 0; ci < info.headers.length; ci++) values.push(row.content[ci] || '');
      let finalText = '';
      for (let i = row.content.length - 1; i >= 0; i--) {
        if (String(row.content[i] || '').trim()) {
          finalText = String(row.content[i]).trim();
          break;
        }
      }
      values.push(finalText, row.merge ? 'Y' : '');
      data.push(values);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), sn);
  }
  
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  
  // 如果是 Google Drive 文件，同步到 Drive
  if (wbInfo._driveFileId && wbInfo._user) {
    try {
      const drive = driveService.getDriveService(wbInfo._user);
      driveService.updateFileContent(drive, wbInfo._driveFileId, buffer)
        .then(result => {
          console.log('Synced to Drive:', wbInfo._driveFileId);
        })
        .catch(err => {
          console.error('Drive sync failed:', err.message);
        });
    } catch (e) {
      console.error('Drive sync error:', e.message);
    }
  }
  
  // 本地文件保存（如果有 filePath）
  if (wbInfo.filePath) {
    fs.writeFileSync(wbInfo.filePath, buffer);
  }
  
  wbInfo.updatedAt = new Date().toISOString();
}

function scheduleSave(id) {
  clearTimeout(saveTimers.get(id));
  saveTimers.set(id, setTimeout(() => {
    const wbInfo = workbooks.get(id);
    if (wbInfo) writeWorkbook(wbInfo);
  }, 3000));
}

function scheduleMetaSave() {
  clearTimeout(scheduleMetaSave.timer);
  scheduleMetaSave.timer = setTimeout(writeMeta, 1000);
}

function applyCellEdit(wbInfo, msg) {
  const info = wbInfo.state.sheets[msg.sheet];
  if (!info || !info.rows[msg.row]) return false;
  const row = info.rows[msg.row];
  if (msg.field === 'seq') row.seq = String(msg.value || '');
  else if (msg.field === 'kr') row.kr = String(msg.value || '');
  else row.content[msg.col] = String(msg.value || '');
  return true;
}

function ensureRowStyles(row, headerCount) {
  if (!row.styles) row.styles = {};
  if (!row.styles.seq) row.styles.seq = {};
  if (!row.styles.kr) row.styles.kr = {};
  if (!Array.isArray(row.styles.content)) row.styles.content = Array.from({ length: headerCount }, () => ({}));
  while (row.styles.content.length < headerCount) row.styles.content.push({});
  return row.styles;
}

function getCellStyle(row, headers, msg) {
  const styles = ensureRowStyles(row, headers.length);
  if (msg.field === 'seq') return styles.seq;
  if (msg.field === 'kr') return styles.kr;
  return styles.content[msg.col] || (styles.content[msg.col] = {});
}

function applyStyle(wbInfo, msg) {
  const info = wbInfo.state.sheets[msg.sheet];
  if (!info || !info.rows[msg.row]) return false;
  const style = getCellStyle(info.rows[msg.row], info.headers, msg);
  Object.assign(style, msg.style || {});
  Object.keys(style).forEach(key => {
    if (style[key] === '' || style[key] === null || style[key] === false) delete style[key];
  });
  meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
  scheduleMetaSave();
  return true;
}

function extractStyles(state) {
  const out = {};
  for (const sn of state.sheetNames) {
    out[sn] = state.sheets[sn].rows.map(row => row.styles || {});
  }
  return out;
}

function applySavedStyles(state, styles) {
  if (!styles) return;
  for (const sn of state.sheetNames) {
    const sheetStyles = styles[sn] || [];
    state.sheets[sn].rows.forEach((row, i) => {
      if (sheetStyles[i]) row.styles = sheetStyles[i];
      ensureRowStyles(row, state.sheets[sn].headers.length);
    });
  }
}

function applyRowOp(wbInfo, msg) {
  const info = wbInfo.state.sheets[msg.sheet];
  if (!info) return false;
  const idx = Math.max(0, Math.min(Number(msg.row) || 0, info.rows.length));
  if (msg.action === 'insert') {
    const contentLen = info.headers.length;
    info.rows.splice(idx + 1, 0, { seq: '', kr: '', content: Array(contentLen).fill(''), styles: { seq: {}, kr: {}, content: Array.from({ length: contentLen }, () => ({})) }, merge: false });
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
    scheduleMetaSave();
    return true;
  }
  if (msg.action === 'duplicate' && info.rows[idx]) {
    const row = info.rows[idx];
    info.rows.splice(idx + 1, 0, { seq: row.seq, kr: row.kr, content: row.content.slice(), styles: JSON.parse(JSON.stringify(row.styles || {})), merge: false });
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
    scheduleMetaSave();
    return true;
  }
  if (msg.action === 'delete' && info.rows[idx]) {
    info.rows.splice(idx, 1);
    if (info.rows[idx] && info.rows[idx].merge) info.rows[idx].merge = false;
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
    scheduleMetaSave();
    return true;
  }
  return false;
}

function loadAllWorkbooks() {
  ensureSeedWorkbook();
  readMeta();
  workbooks.clear();
  for (const filePath of walkFiles(WORKBOOK_DIR).filter(name => /\.xlsx?$/i.test(name))) {
    const file = path.basename(filePath);
    const stat = fs.statSync(filePath);
    const id = file.replace(/\.[^.]+$/, '');
    const saved = meta[id] || {};
    const state = parseWorkbookFile(filePath);
    applySavedStyles(state, saved.styles);
    workbooks.set(id, {
      id,
      name: saved.name || (file === 'default.xlsx' ? '第01话_协作主表.xlsx' : file),
      folder: relativeFolder(filePath),
      filePath,
      uploadedAt: saved.uploadedAt || stat.birthtime.toISOString(),
      updatedAt: stat.mtime.toISOString(),
      state,
    });
  }
  console.log('Loaded workbooks: ' + workbooks.size);
}

function workbookSummary(wbInfo) {
  const rows = Object.values(wbInfo.state.sheets).reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const translated = Object.values(wbInfo.state.sheets).reduce((sum, sheet) => (
    sum + sheet.rows.filter(row => row.content.some(v => String(v || '').trim())).length
  ), 0);
  return {
    id: wbInfo.id,
    name: wbInfo.name,
    folder: wbInfo.folder || '',
    sheetCount: wbInfo.state.sheetNames.length,
    rowCount: rows,
    translated,
    uploadedAt: wbInfo.uploadedAt,
    updatedAt: wbInfo.updatedAt,
  };
}

function listWorkbooks() {
  return [...workbooks.values()].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt))).map(workbookSummary);
}

function listItems(folder) {
  const { clean, full } = resolveFolder(folder);
  fs.mkdirSync(full, { recursive: true });
  const folders = fs.readdirSync(full, { withFileTypes: true })
    .filter(entry => entry.isDirectory())
    .map(entry => {
      const folderPath = clean ? clean + '/' + entry.name : entry.name;
      const stat = fs.statSync(path.join(full, entry.name));
      return { type: 'folder', name: entry.name, path: folderPath, updatedAt: stat.mtime.toISOString() };
    });
  const files = [...workbooks.values()]
    .filter(wb => (wb.folder || '') === clean)
    .map(workbookSummary)
    .map(item => ({ ...item, type: 'file' }));
  return { folder: clean, folders, files };
}

function broadcast(msg, sender, workbookId) {
  const data = JSON.stringify(msg);
  for (const [ws, info] of clients) {
    if (ws === sender || ws.readyState !== 1) continue;
    if (workbookId && info.workbookId !== workbookId) continue;
    ws.send(data);
  }
}

function broadcastList() {
  const msg = { type: 'workbookList', workbooks: listWorkbooks() };
  broadcast(msg);
}

function safeReturnTo(value, fallback = '/latest') {
  const target = String(value || '').trim();
  if (!target.startsWith('/') || target.startsWith('//')) return fallback;
  if (target.startsWith('/auth/')) return fallback;
  return target;
}

const app = express();
app.set('trust proxy', true); // Trust Cloudflare/Render proxy headers so secure cookies work.
const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: '/ws' });

// ---- Session & Authentication ----

app.use(session({
  name: 'pd.sid',
  secret: process.env.SESSION_SECRET || 'dev-secret-change-me-in-production',
  resave: false,
  saveUninitialized: false,
  proxy: true,
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    sameSite: 'lax',
    maxAge: 24 * 60 * 60 * 1000, // 24 hours
  },
}));

app.use(passport.initialize());
app.use(passport.session());

// 只在 Google OAuth 环境变量配置好时才初始化
if (googleConfig.clientID && googleConfig.clientSecret) {
  passport.use(new GoogleStrategy({
    clientID: googleConfig.clientID,
    clientSecret: googleConfig.clientSecret,
    callbackURL: googleConfig.callbackURL,
  }, (accessToken, refreshToken, profile, done) => {
    const user = {
      id: profile.id,
      displayName: profile.displayName,
      email: profile.emails[0].value,
      avatar: profile.photos[0].value,
      accessToken,
      refreshToken,
    };
    return done(null, user);
  }));
  console.log('Google OAuth enabled');
} else {
  console.log('Google OAuth NOT configured — skipping (set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET)');
}

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

wss.on('connection', ws => {
  const cid = 'user_' + Date.now().toString(36);
  clients.set(ws, { name: cid, workbookId: '' });
  ws.send(JSON.stringify({ type: 'workbookList', workbooks: listWorkbooks(), clientId: cid }));
  ws.send(JSON.stringify({ type: 'userCount', count: clients.size }));
  broadcast({ type: 'userCount', count: clients.size });

  ws.on('message', raw => {
    try {
      const msg = JSON.parse(raw);
      const client = clients.get(ws);

      if (msg.type === 'setName') {
        client.name = String(msg.name || cid).slice(0, 40);
        return;
      }

      if (msg.type === 'open') {
        const wbInfo = workbooks.get(msg.workbookId);
        if (!wbInfo) return ws.send(JSON.stringify({ type: 'error', error: 'workbook_not_found' }));
        client.workbookId = msg.workbookId;
        return ws.send(JSON.stringify({ type: 'init', workbook: workbookSummary(wbInfo), state: wbInfo.state }));
      }

      const wbInfo = workbooks.get(msg.workbookId || client.workbookId);
      if (!wbInfo) return;

      if (msg.type === 'edit') {
        applyCellEdit(wbInfo, msg);
        scheduleSave(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'merge') {
        const info = wbInfo.state.sheets[msg.sheet];
        if (info && info.rows[msg.row]) info.rows[msg.row].merge = !!msg.value;
        scheduleSave(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'row') {
        if (applyRowOp(wbInfo, msg)) {
          scheduleSave(wbInfo.id);
          broadcast(msg, ws, wbInfo.id);
        }
      } else if (msg.type === 'style') {
        if (applyStyle(wbInfo, msg)) {
          broadcast(msg, ws, wbInfo.id);
        }
      } else if (msg.type === 'switchSheet') {
        wbInfo.state.currentSheet = msg.sheet;
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'save') {
        writeWorkbook(wbInfo);
        ws.send(JSON.stringify({ type: 'saved', time: new Date().toLocaleTimeString() }));
        broadcastList();
      }
    } catch (e) {
      console.error('Msg err:', e);
    }
  });

  ws.on('close', () => {
    clients.delete(ws);
    broadcast({ type: 'userCount', count: clients.size });
  });
});

// ---- Auth Routes ----

// 发起 Google 登录
app.get('/auth/google', (req, res, next) => {
  if (!googleConfig.clientID) return res.status(503).send('Google OAuth not configured');
  req.session.returnTo = safeReturnTo(req.query.returnTo, '/latest');
  passport.authenticate('google', {
    scope: [
      'profile',
      'email',
      'https://www.googleapis.com/auth/drive',
    ],
    accessType: 'offline',
    prompt: 'consent',
  })(req, res, next);
});

// Google 回调
app.get('/auth/google/callback', (req, res, next) => {
  passport.authenticate('google', (err, user) => {
    if (err) return next(err);
    if (!user) return res.redirect('/');

    req.logIn(user, loginErr => {
      if (loginErr) return next(loginErr);

      req.session.save(saveErr => {
        if (saveErr) return next(saveErr);
        const returnTo = safeReturnTo(req.session.returnTo, '/latest');
        delete req.session.returnTo;
        res.redirect(returnTo);
      });
    });
  })(req, res, next);
});

// 登出
app.get('/auth/logout', (req, res) => {
  const returnTo = safeReturnTo(req.query.returnTo, '/latest');
  req.logout(() => {
    req.session.destroy(() => {
      res.clearCookie('pd.sid');
      res.redirect(returnTo);
    });
  });
});

// 获取当前用户信息
app.get('/api/user', (req, res) => {
  if (req.isAuthenticated()) {
    res.json({
      loggedIn: true,
      user: {
        displayName: req.user.displayName,
        email: req.user.email,
        avatar: req.user.avatar,
      },
    });
  } else {
    res.json({ loggedIn: false });
  }
});

app.use((req, res, next) => {
  if (req.path === '/' || req.path.endsWith('.html')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

app.get(['/latest', '/drive'], (req, res) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/drive-tools', (req, res) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  res.sendFile(path.join(__dirname, 'public', 'drive-tools.html'));
});

app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders(res, filePath) {
    if (filePath.endsWith('.html')) {
      res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
      res.set('Pragma', 'no-cache');
      res.set('Expires', '0');
    }
  },
}));
app.get('/api/status', (req, res) => res.json({ clients: clients.size, wsPath: '/ws' }));
app.get('/api/workbooks', (req, res) => res.json({ workbooks: listWorkbooks() }));
app.get('/api/items', (req, res) => {
  try {
    res.json(listItems(req.query.folder || ''));
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.get('/api/properties', (req, res) => {
  try {
    if (req.query.type === 'folder') {
      const folder = resolveFolder(req.query.path || '');
      const stat = fs.statSync(folder.full);
      const childCount = fs.readdirSync(folder.full).length;
      return res.json({ type: 'folder', name: path.basename(folder.full) || '根目录', path: folder.clean, childCount, createdAt: stat.birthtime.toISOString(), updatedAt: stat.mtime.toISOString() });
    }
    const wbInfo = workbooks.get(String(req.query.id || ''));
    if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
    const stat = fs.statSync(wbInfo.filePath);
    res.json({ ...workbookSummary(wbInfo), type: 'file', size: stat.size, createdAt: stat.birthtime.toISOString() });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/folder', express.json(), (req, res) => {
  try {
    const parent = resolveFolder(req.body && req.body.folder || '');
    const name = safeName(req.body && req.body.name || '新建文件夹');
    if (!name) throw new Error('invalid_name');
    fs.mkdirSync(path.join(parent.full, name), { recursive: false });
    res.json({ ok: true, items: listItems(parent.clean) });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/delete', express.json(), (req, res) => {
  try {
    const body = req.body || {};
    if (body.type === 'folder') {
      const folder = resolveFolder(body.path || '');
      if (!folder.clean) throw new Error('cannot_delete_root');
      fs.rmSync(folder.full, { recursive: true, force: true });
      loadAllWorkbooks();
      return res.json({ ok: true });
    }
    const wbInfo = workbooks.get(String(body.id || ''));
    if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
    fs.rmSync(wbInfo.filePath, { force: true });
    delete meta[wbInfo.id];
    writeMeta();
    workbooks.delete(wbInfo.id);
    broadcastList();
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/rename', express.json(), (req, res) => {
  try {
    const body = req.body || {};
    const name = safeName(body.name || '');
    if (!name) throw new Error('invalid_name');
    if (body.type === 'folder') {
      const folder = resolveFolder(body.path || '');
      if (!folder.clean) throw new Error('cannot_rename_root');
      const dest = path.join(path.dirname(folder.full), name);
      fs.renameSync(folder.full, dest);
      loadAllWorkbooks();
      return res.json({ ok: true });
    }
    const wbInfo = workbooks.get(String(body.id || ''));
    if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
    wbInfo.name = name;
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name, uploadedAt: wbInfo.uploadedAt };
    writeMeta();
    broadcastList();
    res.json({ ok: true, workbook: workbookSummary(wbInfo) });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/move', express.json(), (req, res) => {
  try {
    const body = req.body || {};
    const target = resolveFolder(body.target || '');
    fs.mkdirSync(target.full, { recursive: true });
    if (body.type === 'folder') {
      const folder = resolveFolder(body.path || '');
      if (!folder.clean) throw new Error('cannot_move_root');
      if (target.full === folder.full || target.full.startsWith(folder.full + path.sep)) throw new Error('cannot_move_folder_into_itself');
      const dest = path.join(target.full, path.basename(folder.full));
      fs.renameSync(folder.full, dest);
      loadAllWorkbooks();
      return res.json({ ok: true });
    }
    const wbInfo = workbooks.get(String(body.id || ''));
    if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
    const dest = path.join(target.full, path.basename(wbInfo.filePath));
    fs.renameSync(wbInfo.filePath, dest);
    wbInfo.filePath = dest;
    wbInfo.folder = target.clean;
    workbooks.set(wbInfo.id, wbInfo);
    broadcastList();
    res.json({ ok: true, workbook: workbookSummary(wbInfo) });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.get('/api/download/:id', (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  res.download(wbInfo.filePath, wbInfo.name);
});
app.post('/api/data/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  res.json({ workbook: workbookSummary(wbInfo), state: wbInfo.state });
});
app.post('/api/upload', express.raw({ type: '*/*', limit: '80mb' }), (req, res) => {
  if (!req.body || !req.body.length) return res.status(400).json({ ok: false, error: 'empty_file' });
  try {
    const originalName = safeName(decodeURIComponent(req.header('X-File-Name') || '协作主表.xlsx'));
    const folder = resolveFolder(decodeURIComponent(req.header('X-Folder') || ''));
    const id = makeId(originalName);
    const fileName = id + '.xlsx';
    const filePath = path.join(folder.full, fileName);
    const state = parseWorkbookBuffer(req.body);
    if (!state.sheetNames.length) throw new Error('no usable sheets');
    fs.writeFileSync(filePath, req.body);
    const now = new Date().toISOString();
    const wbInfo = { id, name: originalName, folder: folder.clean, filePath, uploadedAt: now, updatedAt: now, state };
    workbooks.set(id, wbInfo);
    meta[id] = { name: originalName, uploadedAt: now };
    writeMeta();
    broadcastList();
    res.json({ ok: true, workbook: workbookSummary(wbInfo), state });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/edit/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  if (!wbInfo || !applyCellEdit(wbInfo, msg)) return res.status(404).json({ ok: false, error: 'row_not_found' });
  scheduleSave(wbInfo.id);
  broadcast({ type: 'edit', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, col: msg.col, field: msg.field, value: String(msg.value || '') }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/merge/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  const info = wbInfo && wbInfo.state.sheets[msg.sheet];
  if (!info || !info.rows[msg.row]) return res.status(404).json({ ok: false, error: 'row_not_found' });
  info.rows[msg.row].merge = !!msg.value;
  scheduleSave(wbInfo.id);
  broadcast({ type: 'merge', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, value: !!msg.value }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/row/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  if (!wbInfo || !applyRowOp(wbInfo, msg)) return res.status(400).json({ ok: false, error: 'row_op_failed' });
  scheduleSave(wbInfo.id);
  broadcast({ type: 'row', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, action: msg.action }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/style/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  if (!wbInfo || !applyStyle(wbInfo, msg)) return res.status(400).json({ ok: false, error: 'style_failed' });
  broadcast({ type: 'style', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, col: msg.col, field: msg.field, style: msg.style || {} }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/save/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  writeWorkbook(wbInfo);
  broadcastList();
  res.json({ ok: true, time: new Date().toLocaleTimeString() });
});

// ---- Google Drive API Routes ----

// 获取 Google Drive 上的 XLSX 文件列表（含共享文件）
app.get('/api/drive/files', (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  try {
    const drive = driveService.getDriveService(req.user);
    driveService.findOrCreateRootFolder(drive).then(rootId => {
      return driveService.listDriveFiles(drive, rootId);
    }).then(files => {
      const items = files
        .filter(f => {
          const name = (f.name || '').toLowerCase();
          return f.mimeType === 'application/vnd.google-apps.folder' ||
                 f.mimeType === 'application/vnd.google-apps.spreadsheet' ||
                 f.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                 f.mimeType === 'application/vnd.ms-excel' ||
                 name.endsWith('.xlsx') || name.endsWith('.xls');
        })
        .map(f => ({
          id: f.id,
          name: f.name,
          mimeType: f.mimeType,
          isFolder: f.mimeType === 'application/vnd.google-apps.folder',
          size: f.size,
          modifiedTime: f.modifiedTime,
          owners: f.owners,
        }));
      res.json({ ok: true, files: items });
    }).catch(err => {
      console.error('Drive list error:', err);
      res.status(500).json({ ok: false, error: err.message });
    });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// 上传本地 XLSX 到 Google Drive，并加载到当前工作台
app.post('/api/drive/upload', express.raw({ type: '*/*', limit: '80mb' }), async (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  if (!req.body || !req.body.length) {
    return res.status(400).json({ ok: false, error: 'empty_file' });
  }

  try {
    const originalName = safeName(decodeURIComponent(req.header('X-File-Name') || '协作主表.xlsx'));
    const state = parseWorkbookBuffer(req.body);
    if (!state.sheetNames.length) throw new Error('no usable sheets');

    const drive = driveService.getDriveService(req.user);
    const rootId = await driveService.findOrCreateRootFolder(drive);
    const uploaded = await driveService.uploadFile(drive, rootId, originalName, req.body);
    const now = new Date().toISOString();
    const wbInfo = {
      id: uploaded.id,
      name: uploaded.name || originalName,
      folder: '',
      filePath: null,
      uploadedAt: uploaded.modifiedTime || now,
      updatedAt: now,
      state,
      _driveFileId: uploaded.id,
      _user: req.user,
    };

    workbooks.set(wbInfo.id, wbInfo);
    broadcastList();
    res.json({ ok: true, workbook: workbookSummary(wbInfo), state });
  } catch (err) {
    console.error('Drive upload error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// 从 Google Drive 打开文件并加载到内存
app.post('/api/drive/open/:fileId', (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  try {
    const drive = driveService.getDriveService(req.user);
    driveService.downloadFile(drive, req.params.fileId)
      .then(buffer => {
        const state = parseWorkbookBuffer(buffer);
        if (!state.sheetNames.length) throw new Error('no usable sheets');
        // 获取文件信息
        return driveService.getFileInfo(drive, req.params.fileId).then(fileInfo => {
          const id = req.params.fileId;
          workbooks.set(id, {
            id,
            name: fileInfo.name || decodeURIComponent(req.query.name || 'GoogleDrive文件'),
            folder: '',
            filePath: null,
            uploadedAt: fileInfo.modifiedTime || new Date().toISOString(),
            updatedAt: new Date().toISOString(),
            state,
            _driveFileId: id,
            _user: req.user, // 保存用户引用以用于后续保存
          });
          res.json({ ok: true, workbook: workbookSummary(workbooks.get(id)), state });
        });
      })
      .catch(err => {
        console.error('Drive open error:', err);
        res.status(500).json({ ok: false, error: err.message });
      });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

loadAllWorkbooks();
server.listen(PORT, () => console.log('Server: http://localhost:' + PORT));
