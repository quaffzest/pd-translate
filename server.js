const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const DATA_FILE = path.join(__dirname, 'data', '第01话_协作主表.xlsx');
const PORT = process.env.PORT || 3000;

let sheetsData = {};
let sheetNames = [];
let currentSheet = '';
let saveTimer = null;

function parseWorkbook(wb) {
  sheetsData = {};
  sheetNames = [];
  for (let i = 0; i < wb.SheetNames.length; i++) {
    const sn = wb.SheetNames[i];
    if (sn === '使用说明·流程' || sn === '说明') continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
    if (rows.length < 2) continue;
    const h = rows[0]; const dr = rows.slice(1);
    const si = Math.max(h.findIndex(x => String(x).trim() === '序'), 0);
    const ki = Math.max(h.findIndex(x => String(x).trim() === '韩文原文'), 0);
    const ti = Math.max(h.findIndex(x => String(x).trim() === '译者'), 0);
    let ch = [];
    for (let ci = ti; ci < h.length; ci++) {
      const hn = String(h[ci]).trim();
      if (hn === '序' || hn === '韩文原文' || hn === '合并') continue;
      ch.push({ name: hn, idx: ci });
    }
    const mi = h.findIndex(x => String(x).trim() === '合并');
    let sheetRows = [];
    for (let j = 0; j < dr.length; j++) {
      const r = dr[j];
      let content = [];
      for (let k = 0; k < ch.length; k++) content.push(String(r[ch[k].idx] || '').trim());
      const mergeValue = mi >= 0 ? String(r[mi] || '').trim().toUpperCase() : '';
      sheetRows.push({ seq: String(r[si] || '').trim(), kr: String(r[ki] || '').trim(), content: content, merge: mergeValue === 'Y' || mergeValue === 'TRUE' || mergeValue === '1' });
    }
    if (sheetRows.length > 0) { sheetsData[sn] = { headers: ch, rows: sheetRows }; sheetNames.push(sn); }
  }
  if (sheetNames.length > 0) currentSheet = sheetNames[0];
  console.log('Loaded: ' + sheetNames.length + ' sheets, ' + Object.values(sheetsData).reduce((a, s) => a + s.rows.length, 0) + ' rows');
}

function loadXlsx() {
  parseWorkbook(XLSX.readFile(DATA_FILE));
}

function saveXlsx() {
  const wb = XLSX.utils.book_new();
  for (const sn of sheetNames) {
    const info = sheetsData[sn];
    const hdr = ['序', '韩文原文'];
    for (const c of info.headers) hdr.push(c.name);
    hdr.push('终稿', '合并');
    const a = [hdr];
    for (const row of info.rows) {
      const r = [row.seq, row.kr];
      for (let ci = 0; ci < info.headers.length; ci++) r.push(row.content[ci] || '');
      let ft = '';
      for (let i = row.content.length - 1; i >= 0; i--) { if (row.content[i].trim()) { ft = row.content[i].trim(); break; } }
      r.push(ft, row.merge ? 'Y' : '');
      a.push(r);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(a), sn);
  }
  XLSX.writeFile(wb, DATA_FILE);
  console.log('Saved xlsx at ' + new Date().toLocaleTimeString());
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveXlsx, 5000);
}

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: "/ws" });
const clients = new Map();

function broadcast(msg, sender) {
  const data = JSON.stringify(msg);
  for (const [ws] of clients) {
    if (ws !== sender && ws.readyState === 1) ws.send(data);
  }
}

wss.on('connection', (ws) => {
  const cid = 'user_' + Date.now().toString(36);
  clients.set(ws, cid);
  console.log('Connect: ' + cid + ' (' + clients.size + ' total)');
  
  ws.send(JSON.stringify({
    type: 'init', state: { sheets: sheetsData, sheetNames: sheetNames, currentSheet: currentSheet }, clientId: cid
  }));
  broadcast({ type: 'userCount', count: clients.size });
  broadcast({ type: 'userList', users: [...clients.values()] });
  
  ws.on('message', (raw) => {
    try {
      const msg = JSON.parse(raw);
      if (msg.type === 'edit') {
        const info = sheetsData[msg.sheet];
        if (info && info.rows[msg.row]) { info.rows[msg.row].content[msg.col] = msg.value; }
        scheduleSave();
        broadcast(msg, ws);
      } else if (msg.type === 'merge') {
        const info = sheetsData[msg.sheet];
        if (info && info.rows[msg.row]) { info.rows[msg.row].merge = msg.value; }
        scheduleSave();
        broadcast(msg, ws);
      } else if (msg.type === 'switchSheet') {
        currentSheet = msg.sheet;
        broadcast(msg, ws);
      } else if (msg.type === 'save') {
        saveXlsx();
        ws.send(JSON.stringify({ type: 'saved', time: new Date().toLocaleTimeString() }));
      } else if (msg.type === 'setName') {
        clients.set(ws, msg.name);
        broadcast({ type: 'userList', users: [...clients.values()] });
      }
    } catch(e) { console.error('Msg err:', e); }
  });
  
  ws.on('close', () => {
    clients.delete(ws);
    console.log('Disconnect (' + clients.size + ' left)');
    broadcast({ type: 'userCount', count: clients.size });
    broadcast({ type: 'userList', users: [...clients.values()] });
  });
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('/api/status', (req, res) => { res.json({ clients: clients.size, wsPath: '/ws' }); });
app.get('/api/download', (req, res) => {
  res.download(DATA_FILE, '协作主表.xlsx');
});
app.post('/api/data', express.json(), (req, res) => {
  res.json({ state: { sheets: sheetsData, sheetNames: sheetNames, currentSheet: currentSheet } });
});
app.post('/api/upload', express.raw({ type: '*/*', limit: '50mb' }), (req, res) => {
  if (!req.body || !req.body.length) return res.status(400).json({ ok: false, error: 'empty_file' });
  const oldState = { sheets: sheetsData, sheetNames: sheetNames, currentSheet: currentSheet };
  try {
    const wb = XLSX.read(req.body, { type: 'buffer' });
    parseWorkbook(wb);
    if (!sheetNames.length) throw new Error('no usable sheets');
    const backup = DATA_FILE + '.bak';
    if (fs.existsSync(DATA_FILE)) fs.copyFileSync(DATA_FILE, backup);
    fs.writeFileSync(DATA_FILE, req.body);
    const state = { sheets: sheetsData, sheetNames: sheetNames, currentSheet: currentSheet };
    broadcast({ type: 'init', state });
    res.json({ ok: true, state });
  } catch (e) {
    sheetsData = oldState.sheets;
    sheetNames = oldState.sheetNames;
    currentSheet = oldState.currentSheet;
    res.status(400).json({ ok: false, error: e.message });
  }
});
app.post('/api/edit', express.json(), (req, res) => {
  const msg = req.body || {};
  const info = sheetsData[msg.sheet];
  if (!info || !info.rows[msg.row]) return res.status(404).json({ ok: false, error: 'row_not_found' });
  info.rows[msg.row].content[msg.col] = String(msg.value || '');
  scheduleSave();
  broadcast({ type: 'edit', sheet: msg.sheet, row: msg.row, col: msg.col, value: String(msg.value || '') });
  res.json({ ok: true });
});
app.post('/api/merge', express.json(), (req, res) => {
  const msg = req.body || {};
  const info = sheetsData[msg.sheet];
  if (!info || !info.rows[msg.row]) return res.status(404).json({ ok: false, error: 'row_not_found' });
  info.rows[msg.row].merge = !!msg.value;
  scheduleSave();
  broadcast({ type: 'merge', sheet: msg.sheet, row: msg.row, value: !!msg.value });
  res.json({ ok: true });
});
app.post('/api/save', express.json(), (req, res) => {
  saveXlsx();
  res.json({ ok: true, time: new Date().toLocaleTimeString() });
});

loadXlsx();
server.listen(PORT, () => console.log('Server: http://localhost:' + PORT));
