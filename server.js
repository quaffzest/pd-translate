const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const LEGACY_FILE = path.join(DATA_DIR, '第01话_协作主表.xlsx');
const WORKBOOK_DIR = path.join(DATA_DIR, 'workbooks');

const workbooks = new Map();
const clients = new Map();
const saveTimers = new Map();

function ensureSeedWorkbook() {
  fs.mkdirSync(WORKBOOK_DIR, { recursive: true });
  const existing = fs.readdirSync(WORKBOOK_DIR).filter(name => /\.xlsx?$/i.test(name));
  if (!existing.length && fs.existsSync(LEGACY_FILE)) {
    fs.copyFileSync(LEGACY_FILE, path.join(WORKBOOK_DIR, 'default.xlsx'));
  }
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
  XLSX.writeFile(wb, wbInfo.filePath);
  wbInfo.updatedAt = new Date().toISOString();
}

function scheduleSave(id) {
  clearTimeout(saveTimers.get(id));
  saveTimers.set(id, setTimeout(() => {
    const wbInfo = workbooks.get(id);
    if (wbInfo) writeWorkbook(wbInfo);
  }, 3000));
}

function loadAllWorkbooks() {
  ensureSeedWorkbook();
  workbooks.clear();
  for (const file of fs.readdirSync(WORKBOOK_DIR).filter(name => /\.xlsx?$/i.test(name))) {
    const filePath = path.join(WORKBOOK_DIR, file);
    const stat = fs.statSync(filePath);
    const id = file.replace(/\.[^.]+$/, '');
    const state = parseWorkbookFile(filePath);
    workbooks.set(id, {
      id,
      name: file === 'default.xlsx' ? '第01话_协作主表.xlsx' : file,
      filePath,
      uploadedAt: stat.birthtime.toISOString(),
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

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: '/ws' });

wss.on('connection', ws => {
  const cid = 'user_' + Date.now().toString(36);
  clients.set(ws, { name: cid, workbookId: '' });
  ws.send(JSON.stringify({ type: 'workbookList', workbooks: listWorkbooks(), clientId: cid }));
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
        const info = wbInfo.state.sheets[msg.sheet];
        if (info && info.rows[msg.row]) info.rows[msg.row].content[msg.col] = String(msg.value || '');
        scheduleSave(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
      } else if (msg.type === 'merge') {
        const info = wbInfo.state.sheets[msg.sheet];
        if (info && info.rows[msg.row]) info.rows[msg.row].merge = !!msg.value;
        scheduleSave(wbInfo.id);
        broadcast(msg, ws, wbInfo.id);
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

app.use(express.static(path.join(__dirname, 'public')));
app.get('/api/status', (req, res) => res.json({ clients: clients.size, wsPath: '/ws' }));
app.get('/api/workbooks', (req, res) => res.json({ workbooks: listWorkbooks() }));
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
    const id = makeId(originalName);
    const fileName = id + '.xlsx';
    const filePath = path.join(WORKBOOK_DIR, fileName);
    const state = parseWorkbookBuffer(req.body);
    if (!state.sheetNames.length) throw new Error('no usable sheets');
    fs.writeFileSync(filePath, req.body);
    const now = new Date().toISOString();
    const wbInfo = { id, name: originalName, filePath, uploadedAt: now, updatedAt: now, state };
    workbooks.set(id, wbInfo);
    broadcastList();
    res.json({ ok: true, workbook: workbookSummary(wbInfo), state });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/edit/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  const msg = req.body || {};
  const info = wbInfo && wbInfo.state.sheets[msg.sheet];
  if (!info || !info.rows[msg.row]) return res.status(404).json({ ok: false, error: 'row_not_found' });
  info.rows[msg.row].content[msg.col] = String(msg.value || '');
  scheduleSave(wbInfo.id);
  broadcast({ type: 'edit', workbookId: wbInfo.id, sheet: msg.sheet, row: msg.row, col: msg.col, value: String(msg.value || '') }, null, wbInfo.id);
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
app.post('/api/save/:id', express.json(), (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  writeWorkbook(wbInfo);
  broadcastList();
  res.json({ ok: true, time: new Date().toLocaleTimeString() });
});

loadAllWorkbooks();
server.listen(PORT, () => console.log('Server: http://localhost:' + PORT));
