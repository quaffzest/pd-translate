const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const XLSX = require('xlsx-js-style');
const {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  ShadingType,
  TextRun,
  UnderlineType,
} = require('docx');
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
const WORKSPACE_ALLOWED_EMAILS = String(process.env.WORKSPACE_ALLOWED_EMAILS || '')
  .split(',')
  .map(email => email.trim().toLowerCase())
  .filter(Boolean);

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

function stripHashColor(value) {
  const text = String(value || '').replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(text) ? text : '';
}

function modelStyleFromXlsx(style) {
  const out = {};
  if (!style) return out;
  const font = style.font || {};
  if (font.bold) out.bold = true;
  if (font.italic) out.italic = true;
  if (font.underline) out.underline = true;
  if (font.strike) out.strike = true;
  if (font.sz) out.fontSize = String(font.sz);
  if (font.color && font.color.rgb) out.fontColor = '#' + String(font.color.rgb).slice(-6);
  const fill = style.fill && (style.fill.fgColor || style.fill.bgColor);
  if (fill && fill.rgb && String(fill.rgb).slice(-6) !== '000000') out.fillColor = '#' + String(fill.rgb).slice(-6);
  if (style.alignment && style.alignment.horizontal) out.align = String(style.alignment.horizontal).toLowerCase();
  if (style.alignment && style.alignment.wrapText) out.wrap = true;
  return out;
}

function xlsxStyleFromModel(style) {
  const input = style || {};
  const out = {};
  if (input.bold || input.italic || input.underline || input.strike || input.fontSize || input.fontColor) {
    out.font = {};
    if (input.bold) out.font.bold = true;
    if (input.italic) out.font.italic = true;
    if (input.underline) out.font.underline = true;
    if (input.strike) out.font.strike = true;
    if (input.fontSize) out.font.sz = Number(input.fontSize);
    const color = stripHashColor(input.fontColor);
    if (color) out.font.color = { rgb: color };
  }
  const fill = stripHashColor(input.fillColor);
  if (fill) {
    out.fill = { patternType: 'solid', fgColor: { rgb: fill } };
  }
  if (input.align || input.wrap === true || input.wrap === false) {
    out.alignment = {};
    if (input.align) out.alignment.horizontal = input.align;
    if (input.wrap === true) out.alignment.wrapText = true;
    if (input.wrap === false) out.alignment.wrapText = false;
  }
  return out;
}

function mergeCellStyle(existing, next) {
  if (!existing) return next;
  return {
    ...existing,
    ...next,
    font: { ...(existing.font || {}), ...(next.font || {}) },
    fill: { ...(existing.fill || {}), ...(next.fill || {}) },
    alignment: { ...(existing.alignment || {}), ...(next.alignment || {}) },
  };
}

function setSheetCell(ws, rowIndex, colIndex, value, style) {
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  if (!ws[addr]) ws[addr] = { t: 's', v: '' };
  XLSX.utils.sheet_add_aoa(ws, [[value]], { origin: addr });
  const modelStyle = xlsxStyleFromModel(style);
  if (Object.keys(modelStyle).length) ws[addr].s = mergeCellStyle(ws[addr].s, modelStyle);
}

function finalContentIndex(row) {
  const content = row && row.content || [];
  for (let i = content.length - 1; i >= 0; i--) {
    if (String(content[i] || '').trim()) return i;
  }
  return -1;
}

function finalTextAndStyle(row) {
  const idx = finalContentIndex(row);
  if (idx < 0) return { text: '', style: {} };
  const style = row.styles && row.styles.content && row.styles.content[idx] || {};
  return { text: String(row.content[idx] || '').trim(), style };
}

function buildDocumentModel(state) {
  const sheetName = state.currentSheet || (state.sheetNames || [])[0] || '';
  const sheet = state.sheets && state.sheets[sheetName];
  const paragraphs = [];
  if (!sheet) return { sheetName, paragraphs, page: defaultDocumentPage() };

  for (let i = 0; i < sheet.rows.length; i++) {
    const row = sheet.rows[i];
    const current = finalTextAndStyle(row);
    if (!current.text) continue;
    const paragraph = { rows: [i], runs: [{ text: current.text, style: current.style || {} }] };
    let next = i + 1;
    while (next < sheet.rows.length && sheet.rows[next].merge) {
      const merged = finalTextAndStyle(sheet.rows[next]);
      if (merged.text) {
        paragraph.rows.push(next);
        paragraph.runs.push({ text: merged.text, style: merged.style || {} });
      }
      next++;
    }
    paragraphs.push(paragraph);
  }

  return { sheetName, paragraphs, page: defaultDocumentPage() };
}

function defaultDocumentPage() {
  return {
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    paragraph: { line: 420, before: 0, after: 180 },
    font: 'SimSun',
    fontSize: 24,
  };
}

function parseWorkbookBuffer(buffer) {
  return parseWorkbook(XLSX.read(buffer, { type: 'buffer', cellStyles: true }));
}

function parseWorkbookFile(filePath) {
  return parseWorkbook(XLSX.readFile(filePath, { cellStyles: true }));
}

function parseWorkbook(wb) {
  const sheets = {};
  const sheetNames = [];
  let currentSheet = '';

  for (const sn of wb.SheetNames) {
    if (sn === '使用说明·流程' || sn === '说明') continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
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

    const sheetRows = dr.map((r, rowOffset) => {
      const excelRow = rowOffset + 1;
      const content = headers.map(c => String(r[c.idx] || '').trim());
      const mergeValue = mi >= 0 ? String(r[mi] || '').trim().toUpperCase() : '';
      return {
        seq: String(r[si] || '').trim(),
        kr: String(r[ki] || '').trim(),
        content,
        styles: {
          seq: modelStyleFromXlsx((ws[XLSX.utils.encode_cell({ r: excelRow, c: si })] || {}).s),
          kr: modelStyleFromXlsx((ws[XLSX.utils.encode_cell({ r: excelRow, c: ki })] || {}).s),
          content: headers.map(c => modelStyleFromXlsx((ws[XLSX.utils.encode_cell({ r: excelRow, c: c.idx })] || {}).s)),
        },
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
  const wb = wbInfo._sourceBuffer ? XLSX.read(wbInfo._sourceBuffer, { type: 'buffer', cellStyles: true }) : XLSX.utils.book_new();
  for (const sn of wbInfo.state.sheetNames) {
    const info = wbInfo.state.sheets[sn];
    let ws = wb.Sheets[sn];
    if (!ws) {
      const header = ['序', '韩文原文', ...info.headers.map(c => c.name), '终稿', '合并'];
      ws = XLSX.utils.aoa_to_sheet([header]);
      XLSX.utils.book_append_sheet(wb, ws, sn);
    }
    info.rows.forEach((row, rowIndex) => {
      const excelRow = rowIndex + 1;
      const styles = row.styles || {};
      setSheetCell(ws, excelRow, info.seqIdx || 0, row.seq || '', styles.seq || {});
      setSheetCell(ws, excelRow, info.krIdx || 1, row.kr || '', styles.kr || {});
      info.headers.forEach((header, contentIndex) => {
        const value = (row.content || [])[contentIndex] || '';
        const style = styles.content && styles.content[contentIndex] || {};
        setSheetCell(ws, excelRow, header.idx, value, style);
      });
      if (Number.isInteger(info.mergeIdx) && info.mergeIdx >= 0) {
        setSheetCell(ws, excelRow, info.mergeIdx, row.merge ? 'Y' : '', {});
      }
    });
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function docxAlignment(style) {
  const align = String(style && style.align || '').toLowerCase();
  if (align === 'center') return AlignmentType.CENTER;
  if (align === 'right') return AlignmentType.RIGHT;
  if (align === 'both' || align === 'justify') return AlignmentType.JUSTIFIED;
  return AlignmentType.LEFT;
}

function docxRunFromModel(run, defaults) {
  const style = run.style || {};
  const options = {
    text: run.text || '',
    bold: !!style.bold,
    italics: !!style.italic,
    strike: !!style.strike,
    size: Number(style.fontSize || defaults.fontSize) * 2,
    font: defaults.font,
  };
  if (style.underline) options.underline = { type: UnderlineType.SINGLE };
  const color = stripHashColor(style.fontColor);
  if (color) options.color = color;
  const fill = stripHashColor(style.fillColor);
  if (fill) {
    options.shading = { type: ShadingType.CLEAR, color: 'auto', fill };
  }
  return new TextRun(options);
}

function documentModelToDocx(model, title) {
  const page = model.page || defaultDocumentPage();
  const defaults = { font: page.font || 'SimSun', fontSize: page.fontSize || 24 };
  const children = [];

  if (title) {
    children.push(new Paragraph({
      children: [new TextRun({ text: title.replace(/\.[^.]+$/, ''), bold: true, size: 28, font: defaults.font })],
      spacing: { after: 240 },
    }));
  }

  for (const paragraph of model.paragraphs || []) {
    const firstStyle = paragraph.runs && paragraph.runs[0] && paragraph.runs[0].style || {};
    children.push(new Paragraph({
      children: (paragraph.runs || []).map((run, index) => docxRunFromModel({
        ...run,
        text: index > 0 ? String(run.text || '') : run.text,
      }, defaults)),
      alignment: docxAlignment(firstStyle),
      spacing: page.paragraph || { line: 420, before: 0, after: 180 },
      keepLines: !!firstStyle.keepLines,
      keepNext: !!firstStyle.keepNext,
    }));
  }

  if (!children.length) {
    children.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
  }

  return new Document({
    sections: [{
      properties: { page: { margin: page.margin || defaultDocumentPage().margin } },
      children,
    }],
  });
}

async function workbookToDocxBuffer(wbInfo) {
  const model = buildDocumentModel(wbInfo.state);
  const doc = documentModelToDocx(model, wbInfo.name);
  return Packer.toBuffer(doc);
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
    wbInfo._driveSignature = driveSignature(updatedInfo);
  } else {
    const buffer = workbookToBuffer(wbInfo);
    const updatedInfo = await driveService.updateFileContent(drive, wbInfo._driveFileId, buffer);
    wbInfo._driveModifiedTime = updatedInfo && updatedInfo.modifiedTime || wbInfo._driveModifiedTime;
    wbInfo._driveSignature = driveSignature(updatedInfo) || wbInfo._driveSignature;
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
    const remoteSignature = driveSignature(fileInfo);
    if (remoteSignature && remoteSignature === wbInfo._driveSignature) return false;
    if (!remoteSignature && fileInfo.modifiedTime === wbInfo._driveModifiedTime) return false;

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
    wbInfo._driveSignature = remoteSignature;
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

function driveSignature(file) {
  if (!file) return '';
  return [
    file.modifiedTime || '',
    file.version || '',
    file.headRevisionId || '',
    file.md5Checksum || '',
  ].join('|');
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

const sessionMiddleware = session({
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
});

app.use(sessionMiddleware);

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

function isAllowedUser(user) {
  if (!WORKSPACE_ALLOWED_EMAILS.length) return true;
  return WORKSPACE_ALLOWED_EMAILS.includes(String(user && user.email || '').toLowerCase());
}

function requireAuth(req, res, next) {
  if (!req.isAuthenticated || !req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: 'login_required' });
  }
  if (!isAllowedUser(req.user)) {
    return res.status(403).json({ ok: false, error: 'forbidden' });
  }
  next();
}

function requirePageAuth(req, res, next) {
  if (!req.isAuthenticated || !req.isAuthenticated()) {
    return res.redirect('/auth/google?returnTo=' + encodeURIComponent(req.originalUrl || '/workbench'));
  }
  if (!isAllowedUser(req.user)) {
    return res.status(403).send('This Google account is not allowed to access this workbench.');
  }
  next();
}

wss.on('connection', (ws, req) => {
  sessionMiddleware(req, {}, () => {
    const user = req.session && req.session.passport && req.session.passport.user;
    if (!user || !isAllowedUser(user)) {
      ws.close(1008, 'login_required');
      return;
    }
    attachWorkspaceSocket(ws, user);
  });
});

function attachWorkspaceSocket(ws, user) {
  const cid = 'user_' + Date.now().toString(36);
  clients.set(ws, { name: user.displayName || user.email || cid, workbookId: '', user });
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
        if (wbInfo._driveFileId && client.user) wbInfo._user = client.user;
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
}

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
      allowed: isAllowedUser(req.user),
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

app.use('/api', (req, res, next) => {
  if (req.path === '/user') return next();
  return requireAuth(req, res, next);
});

app.use((req, res, next) => {
  if (req.path === '/' || req.path.endsWith('.html') || req.path.startsWith('/api/drive/')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

app.get(['/', '/latest', '/drive', '/workbench', '/index.html'], requirePageAuth, (req, res) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/drive-tools', requirePageAuth, (req, res) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  res.sendFile(path.join(__dirname, 'public', 'drive-tools.html'));
});

app.use((req, res, next) => {
  if (req.path.endsWith('.html')) return requirePageAuth(req, res, next);
  next();
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
app.get('/api/document/:id', (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  res.json({ ok: true, document: buildDocumentModel(wbInfo.state) });
});
app.get('/api/export-docx/:id', async (req, res) => {
  const wbInfo = workbooks.get(req.params.id);
  if (!wbInfo) return res.status(404).json({ ok: false, error: 'workbook_not_found' });
  try {
    const buffer = await workbookToDocxBuffer(wbInfo);
    const name = safeName((wbInfo.name || 'final').replace(/\.[^.]+$/, '') + '.docx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(name)}`);
    res.send(buffer);
  } catch (e) {
    console.error('DOCX export failed:', e);
    res.status(500).json({ ok: false, error: e.message });
  }
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
    if (wbInfo._driveFileId && req.isAuthenticated && req.isAuthenticated()) {
      wbInfo._user = req.user;
    }
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
    version: file.version,
    headRevisionId: file.headRevisionId,
    md5Checksum: file.md5Checksum,
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
      _driveSignature: driveSignature(uploaded),
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
    const remoteSignature = driveSignature(fileInfo);

    if (existing && ((existing._driveSignature && existing._driveSignature === remoteSignature) || (!remoteSignature && existing._driveModifiedTime === fileInfo.modifiedTime))) {
      existing._user = req.user;
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
      _driveSignature: remoteSignature,
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
