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
const driveSyncTimers = new Map();
const driveRefreshInFlight = new Set();
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
      sheets[sn] = { headers, rows: sheetRows, seqIdx: si, krIdx: ki, mergeIdx: mi };
      sheetNames.push(sn);
    }
  }

  if (sheetNames.length) currentSheet = sheetNames[0];
  return { sheets, sheetNames, currentSheet };
}

function workbookToBuffer(wbInfo) {
  const wb = wbInfo._sourceBuffer ? XLSX.read(wbInfo._sourceBuffer, { type: 'buffer' }) : XLSX.utils.book_new();
  for (const sn of wbInfo.state.sheetNames) {
    const info = wbInfo.state.sheets[sn];
    let ws = wb.Sheets[sn];
    if (!ws) {
      const header = ['序', '韩文原文', ...info.headers.map(c => c.name), '终稿', '合并'];
      ws = XLSX.utils.aoa_to_sheet([header]);
      XLSX.utils.book_append_sheet(wb, ws, sn);
    }
    const columns = [
      { idx: info.seqIdx || 0, values: info.rows.map(row => row.seq || '') },
      { idx: info.krIdx || 1, values: info.rows.map(row => row.kr || '') },
      ...info.headers.map((header, contentIndex) => ({
        idx: header.idx,
        values: info.rows.map(row => (row.content || [])[contentIndex] || ''),
      })),
    ];
    if (Number.isInteger(info.mergeIdx) && info.mergeIdx >= 0) {
      columns.push({ idx: info.mergeIdx, values: info.rows.map(row => row.merge ? 'Y' : '') });
    }
    for (const col of columns) {
      if (!Number.isInteger(col.idx) || col.idx < 0) continue;
      const colName = XLSX.utils.encode_col(col.idx);
      col.values.forEach((value, rowIndex) => {
        XLSX.utils.sheet_add_aoa(ws, [[value]], { origin: `${colName}${rowIndex + 2}` });
      });
    }
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

async function syncWorkbookToDrive(wbInfo) {
  if (!wbInfo._driveFileId || !wbInfo._user) return null;
  wbInfo.driveSync = { status: 'syncing', time: new Date().toISOString(), error: '' };
  const drive = driveService.getDriveService(wbInfo._user);
  if (wbInfo._driveMimeType === 'application/vnd.google-apps.spreadsheet') {
    const sheets = driveService.getSheetsService(wbInfo._user);
    await driveService.updateGoogleSheetValues(sheets, wbInfo._driveFileId, wbInfo.state);
    const updatedInfo = await driveService.getFileInfo(drive, wbInfo._driveFileId);
    wbInfo._driveModifiedTime = updatedInfo.modifiedTime;
  } else {
    const buffer = workbookToBuffer(wbInfo);
    const updatedInfo = await driveService.updateFileContent(drive, wbInfo._driveFileId, buffer);
    wbInfo._driveModifiedTime = updatedInfo && updatedInfo.modifiedTime || wbInfo._driveModifiedTime;
    wbInfo._sourceBuffer = buffer;
  }
  wbInfo.driveSync = { status: 'synced', time: new Date().toISOString(), error: '' };
  wbInfo._lastSyncedRevision = wbInfo._revision || 0;
  return wbInfo.driveSync;
}

function markWorkbookChanged(wbInfo) {
  wbInfo._revision = (wbInfo._revision || 0) + 1;
  wbInfo.updatedAt = new Date().toISOString();
}

function hasActiveClients(workbookId) {
  for (const client of clients.values()) {
    if (client.workbookId === workbookId) return true;
  }
  return false;
}

async function refreshWorkbookFromDrive(wbInfo) {
  if (!wbInfo || !wbInfo._driveFileId || !wbInfo._user) return false;
  if (wbInfo.driveSync && wbInfo.driveSync.status === 'syncing') return false;
  if (driveRefreshInFlight.has(wbInfo.id)) return false;

  driveRefreshInFlight.add(wbInfo.id);
  try {
    const drive = driveService.getDriveService(wbInfo._user);
    const fileInfo = await driveService.getFileInfo(drive, wbInfo._driveFileId);
    if (!fileInfo.modifiedTime || fileInfo.modifiedTime === wbInfo._driveModifiedTime) return false;

    const localRevision = wbInfo._revision || 0;
    const syncedRevision = wbInfo._lastSyncedRevision || 0;
    if (localRevision !== syncedRevision) {
      wbInfo.driveSync = {
        status: 'conflict',
        time: new Date().toISOString(),
        error: 'Drive 文件已在外部更新，但工作台也有未同步修改。请先保存或重新打开文件。',
      };
      broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id);
      return false;
    }

    if (fileInfo.mimeType !== 'application/vnd.google-apps.spreadsheet' && !driveItemSummary(fileInfo).isSpreadsheet) {
      return false;
    }

    const sourceBuffer = await driveService.downloadFile(drive, wbInfo._driveFileId);
    const state = parseWorkbookBuffer(sourceBuffer);
    if (!state.sheetNames.length) return false;

    wbInfo.name = fileInfo.name || wbInfo.name;
    wbInfo.state = state;
    wbInfo._driveMimeType = fileInfo.mimeType;
    wbInfo._driveModifiedTime = fileInfo.modifiedTime;
    wbInfo._sourceBuffer = fileInfo.mimeType === 'application/vnd.google-apps.spreadsheet' ? null : sourceBuffer;
    wbInfo.updatedAt = new Date().toISOString();
    wbInfo.uploadedAt = fileInfo.modifiedTime || wbInfo.uploadedAt;
    wbInfo.driveSync = { status: 'remote_updated', time: new Date().toISOString(), error: '' };
    wbInfo._revision = localRevision;
    wbInfo._lastSyncedRevision = localRevision;
    broadcast({ type: 'init', workbook: workbookSummary(wbInfo), state: wbInfo.state }, null, wbInfo.id);
    broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id);
    return true;
  } catch (err) {
    console.error('Drive refresh failed:', err.message);
    return false;
  } finally {
    driveRefreshInFlight.delete(wbInfo.id);
  }
}

setInterval(() => {
  for (const wbInfo of workbooks.values()) {
    if (wbInfo._driveFileId && hasActiveClients(wbInfo.id)) {
      refreshWorkbookFromDrive(wbInfo);
    }
  }
}, 10000);

function writeWorkbook(wbInfo) {
  const buffer = workbookToBuffer(wbInfo);

  if (wbInfo._driveFileId && wbInfo._user) {
    syncWorkbookToDrive(wbInfo)
      .then(() => broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id))
      .catch(err => {
        wbInfo.driveSync = { status: 'failed', time: new Date().toISOString(), error: err.message };
        console.error('Drive sync failed:', err.message);
        broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id);
      });
  }
  
  // 本地文件保存（如果有 filePath）
  if (wbInfo.filePath) {
    fs.writeFileSync(wbInfo.filePath, buffer);
  }
  
  wbInfo.updatedAt = new Date().toISOString();
}

async function forceSaveWorkbook(wbInfo) {
  clearTimeout(saveTimers.get(wbInfo.id));
  clearTimeout(driveSyncTimers.get(wbInfo.id));
  const buffer = workbookToBuffer(wbInfo);
  if (wbInfo.filePath) {
    fs.writeFileSync(wbInfo.filePath, buffer);
  }
  if (wbInfo._driveFileId && wbInfo._user) {
    await syncWorkbookToDrive(wbInfo);
  }
  wbInfo.updatedAt = new Date().toISOString();
  return wbInfo.driveSync || null;
}

function scheduleSave(id) {
  clearTimeout(saveTimers.get(id));
  saveTimers.set(id, setTimeout(() => {
    const wbInfo = workbooks.get(id);
    if (wbInfo) writeWorkbook(wbInfo);
  }, 3000));
}

function scheduleDriveSync(id) {
  clearTimeout(driveSyncTimers.get(id));
  driveSyncTimers.set(id, setTimeout(() => {
    const wbInfo = workbooks.get(id);
    if (!wbInfo || !wbInfo._driveFileId) return;
    syncWorkbookToDrive(wbInfo)
      .then(() => broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id))
      .catch(err => {
        wbInfo.driveSync = { status: 'failed', time: new Date().toISOString(), error: err.message };
        broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id);
      });
  }, 8000));
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
  markWorkbookChanged(wbInfo);
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
  markWorkbookChanged(wbInfo);
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
    markWorkbookChanged(wbInfo);
    return true;
  }
  if (msg.action === 'duplicate' && info.rows[idx]) {
    const row = info.rows[idx];
    info.rows.splice(idx + 1, 0, { seq: row.seq, kr: row.kr, content: row.content.slice(), styles: JSON.parse(JSON.stringify(row.styles || {})), merge: false });
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
    scheduleMetaSave();
    markWorkbookChanged(wbInfo);
    return true;
  }
  if (msg.action === 'delete' && info.rows[idx]) {
    info.rows.splice(idx, 1);
    if (info.rows[idx] && info.rows[idx].merge) info.rows[idx].merge = false;
    meta[wbInfo.id] = { ...(meta[wbInfo.id] || {}), name: wbInfo.name, uploadedAt: wbInfo.uploadedAt, styles: extractStyles(wbInfo.state) };
    scheduleMetaSave();
    markWorkbookChanged(wbInfo);
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
    source: wbInfo._driveFileId ? 'drive' : 'server',
    driveFileId: wbInfo._driveFileId || '',
    driveSync: wbInfo.driveSync || null,
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
    .filter(wb => !wb._driveFileId && (wb.folder || '') === clean)
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

function safeReturnTo(value, fallback = '/workbench') {
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

      if (msg.type === 'leave') {
        client.workbookId = '';
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
        scheduleDriveSync(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'merge') {
        const info = wbInfo.state.sheets[msg.sheet];
        if (info && info.rows[msg.row]) {
          info.rows[msg.row].merge = !!msg.value;
          markWorkbookChanged(wbInfo);
        }
        scheduleSave(wbInfo.id);
        scheduleDriveSync(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'row') {
        if (applyRowOp(wbInfo, msg)) {
          scheduleSave(wbInfo.id);
          scheduleDriveSync(wbInfo.id);
          broadcast(msg, ws, wbInfo.id);
        }
      } else if (msg.type === 'style') {
        if (applyStyle(wbInfo, msg)) {
          scheduleDriveSync(wbInfo.id);
          broadcast(msg, ws, wbInfo.id);
        }
      } else if (msg.type === 'switchSheet') {
        wbInfo.state.currentSheet = msg.sheet;
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'save') {
        forceSaveWorkbook(wbInfo)
          .then(sync => {
            ws.send(JSON.stringify({ type: 'saved', time: new Date().toLocaleTimeString() }));
            if (sync) broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync }, null, wbInfo.id);
            broadcastList();
          })
          .catch(err => {
            wbInfo.driveSync = { status: 'failed', time: new Date().toISOString(), error: err.message };
            ws.send(JSON.stringify({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }));
            console.error('Save failed:', err.message);
          });
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
  req.session.returnTo = safeReturnTo(req.query.returnTo, '/workbench');
  passport.authenticate('google', {
    scope: [
      'profile',
      'email',
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/spreadsheets',
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
        const returnTo = safeReturnTo(req.session.returnTo, '/workbench');
        delete req.session.returnTo;
        res.redirect(returnTo);
      });
    });
  })(req, res, next);
});

// 登出
app.get('/auth/logout', (req, res) => {
  const returnTo = safeReturnTo(req.query.returnTo, '/workbench');
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
  if (req.path === '/' || req.path.endsWith('.html') || req.path.startsWith('/api/drive/')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

app.get(['/latest', '/drive', '/workbench'], (req, res) => {
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
  scheduleDriveSync(wbInfo.id);
  broadcast({ type: 'edit', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, col: msg.col, field: msg.field, value: String(msg.value || '') }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/merge/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  const info = wbInfo && wbInfo.state.sheets[msg.sheet];
  if (!info || !info.rows[msg.row]) return res.status(404).json({ ok: false, error: 'row_not_found' });
  info.rows[msg.row].merge = !!msg.value;
  markWorkbookChanged(wbInfo);
  scheduleSave(wbInfo.id);
  scheduleDriveSync(wbInfo.id);
  broadcast({ type: 'merge', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, value: !!msg.value }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/row/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  if (!wbInfo || !applyRowOp(wbInfo, msg)) return res.status(400).json({ ok: false, error: 'row_op_failed' });
  scheduleSave(wbInfo.id);
  scheduleDriveSync(wbInfo.id);
  broadcast({ type: 'row', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, action: msg.action }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/style/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  if (!wbInfo || !applyStyle(wbInfo, msg)) return res.status(400).json({ ok: false, error: 'style_failed' });
  scheduleDriveSync(wbInfo.id);
  broadcast({ type: 'style', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, col: msg.col, field: msg.field, style: msg.style || {} }, null, wbInfo.id);
  res.json({ ok: true });
});
app.post('/api/save/:id', express.json(), async (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  try {
    const sync = await forceSaveWorkbook(wbInfo);
    if (sync) broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync }, null, wbInfo.id);
    broadcastList();
    res.json({ ok: true, time: new Date().toLocaleTimeString(), driveSync: sync });
  } catch (e) {
    wbInfo.driveSync = { status: 'failed', time: new Date().toISOString(), error: e.message };
    broadcast({ type: 'driveSync', workbookId: wbInfo.id, sync: wbInfo.driveSync }, null, wbInfo.id);
    res.status(500).json({ ok: false, error: e.message, driveSync: wbInfo.driveSync });
  }
});

// ---- Google Drive API Routes ----

function driveItemSummary(file) {
  const name = file.name || '';
  const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
  const isSpreadsheet = isFolder ? false : (
    file.isSpreadsheet ||
    file.mimeType === 'application/vnd.google-apps.spreadsheet' ||
    file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.mimeType === 'application/vnd.ms-excel' ||
    /\.xlsx?$/i.test(name)
  );
  return {
    id: file.id,
    name,
    mimeType: file.mimeType,
    isFolder,
    isSpreadsheet,
    size: file.size,
    modifiedTime: file.modifiedTime,
    createdTime: file.createdTime,
    owners: file.owners,
    webViewLink: file.webViewLink,
  };
}

// 浏览当前 Google Drive 文件夹
app.get('/api/drive/items', async (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  try {
    const drive = driveService.getDriveService(req.user);
    const folderId = String(req.query.folderId || 'root');
    const files = await driveService.listDriveItems(drive, folderId);
    res.json({ ok: true, folderId, items: files.map(driveItemSummary) });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// 兼容旧 Drive 文件选择器
app.get('/api/drive/files', async (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  try {
    const drive = driveService.getDriveService(req.user);
    const files = await driveService.listDriveItems(drive, String(req.query.folderId || 'root'));
    res.json({ ok: true, files: files.map(driveItemSummary) });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.post('/api/drive/folder', express.json(), async (req, res) => {
  if (!req.isAuthenticated()) return res.status(401).json({ ok: false, error: 'login_required' });
  try {
    const drive = driveService.getDriveService(req.user);
    const name = safeName(req.body && req.body.name || '新建文件夹');
    const folder = await driveService.createFolder(drive, req.body && req.body.folderId || 'root', name);
    res.json({ ok: true, item: driveItemSummary(folder) });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/api/drive/rename', express.json(), async (req, res) => {
  if (!req.isAuthenticated()) return res.status(401).json({ ok: false, error: 'login_required' });
  try {
    const drive = driveService.getDriveService(req.user);
    const item = await driveService.renameFile(drive, req.body && req.body.fileId, safeName(req.body && req.body.name || ''));
    res.json({ ok: true, item: driveItemSummary(item) });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/api/drive/move', express.json(), async (req, res) => {
  if (!req.isAuthenticated()) return res.status(401).json({ ok: false, error: 'login_required' });
  try {
    const drive = driveService.getDriveService(req.user);
    const item = await driveService.moveFile(drive, req.body && req.body.fileId, req.body && req.body.targetFolderId || 'root');
    res.json({ ok: true, item: driveItemSummary(item) });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/api/drive/trash', express.json(), async (req, res) => {
  if (!req.isAuthenticated()) return res.status(401).json({ ok: false, error: 'login_required' });
  try {
    const drive = driveService.getDriveService(req.user);
    const item = await driveService.trashFile(drive, req.body && req.body.fileId);
    res.json({ ok: true, item });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// 上传本地 XLSX 到当前 Google Drive 文件夹
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
    const uploaded = await driveService.uploadFile(drive, String(req.query.folderId || 'root'), originalName, req.body);
    const now = new Date().toISOString();
    const id = 'drive:' + uploaded.id;
    const wbInfo = {
      id,
      name: uploaded.name || originalName,
      folder: '',
      filePath: null,
      uploadedAt: uploaded.modifiedTime || now,
      updatedAt: now,
      state,
      _driveFileId: uploaded.id,
      _driveMimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      _driveModifiedTime: uploaded.modifiedTime || now,
      _sourceBuffer: Buffer.from(req.body),
      _user: req.user,
      driveSync: { status: 'synced', time: now, error: '' },
    };

    workbooks.set(wbInfo.id, wbInfo);
    broadcastList();
    res.json({
      ok: true,
      workbook: workbookSummary(wbInfo),
      driveFile: driveItemSummary(uploaded),
      state,
    });
  } catch (err) {
    console.error('Drive upload error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// 从 Google Drive 打开文件并加载到内存
app.post('/api/drive/open/:fileId', async (req, res) => {
  if (!req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  try {
    const drive = driveService.getDriveService(req.user);
    const workbookId = 'drive:' + req.params.fileId;
    const existing = workbooks.get(workbookId);
    const fileInfo = await driveService.getFileInfo(drive, req.params.fileId);

    if (existing && existing._driveModifiedTime === fileInfo.modifiedTime) {
      if (!existing._user) existing._user = req.user;
      return res.json({ ok: true, workbook: workbookSummary(existing), state: existing.state });
    }

    if (fileInfo.mimeType !== 'application/vnd.google-apps.spreadsheet' && !driveItemSummary(fileInfo).isSpreadsheet) {
      throw new Error('not_spreadsheet');
    }

    const sourceBuffer = await driveService.downloadFile(drive, req.params.fileId);
    const state = parseWorkbookBuffer(sourceBuffer);
    if (!state.sheetNames.length) throw new Error('no usable sheets');

    workbooks.set(workbookId, {
      id: workbookId,
      name: fileInfo.name || decodeURIComponent(req.query.name || 'GoogleDrive文件'),
      folder: '',
      filePath: null,
      uploadedAt: fileInfo.modifiedTime || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      state,
      _driveFileId: req.params.fileId,
      _driveMimeType: fileInfo.mimeType,
      _driveModifiedTime: fileInfo.modifiedTime,
      _sourceBuffer: fileInfo.mimeType === 'application/vnd.google-apps.spreadsheet' ? null : sourceBuffer,
      _user: req.user,
      driveSync: { status: 'synced', time: new Date().toISOString(), error: '' },
    });
    res.json({ ok: true, workbook: workbookSummary(workbooks.get(workbookId)), state });
  } catch (e) {
    console.error('Drive open error:', e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

loadAllWorkbooks();
server.listen(PORT, () => console.log('Server: http://localhost:' + PORT));
