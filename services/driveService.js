// services/driveService.js
const { google } = require('googleapis');
const { Readable } = require('stream');
const googleConfig = require('../config/google');

/**
 * 从用户 session 创建已认证的 Drive 客户端
 */
function getDriveService(user) {
  if (!user || !user.accessToken) {
    throw new Error('User not authenticated with Google');
  }
  const oauth2Client = new google.auth.OAuth2(
    googleConfig.clientID,
    googleConfig.clientSecret,
    googleConfig.callbackURL
  );
  oauth2Client.setCredentials({
    access_token: user.accessToken,
    refresh_token: user.refreshToken,
  });
  return google.drive({ version: 'v3', auth: oauth2Client });
}

function getSheetsService(user) {
  if (!user || !user.accessToken) {
    throw new Error('User not authenticated with Google');
  }
  const oauth2Client = new google.auth.OAuth2(
    googleConfig.clientID,
    googleConfig.clientSecret,
    googleConfig.callbackURL
  );
  oauth2Client.setCredentials({
    access_token: user.accessToken,
    refresh_token: user.refreshToken,
  });
  return google.sheets({ version: 'v4', auth: oauth2Client });
}

function isSpreadsheetFile(file) {
  const name = String(file && file.name || '').toLowerCase();
  return file && (
    file.mimeType === 'application/vnd.google-apps.spreadsheet' ||
    file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.mimeType === 'application/vnd.ms-excel' ||
    name.endsWith('.xlsx') ||
    name.endsWith('.xls')
  );
}

/**
 * 查找或创建 pd-translate 根文件夹（在用户自己的 Drive 中）
 */
async function findOrCreateRootFolder(drive) {
  const res = await drive.files.list({
    q: `name='${googleConfig.driveRootFolder}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
    pageSize: 5,
  });
  if (res.data.files.length > 0) {
    return res.data.files[0].id;
  }
  const folder = await drive.files.create({
    requestBody: {
      name: googleConfig.driveRootFolder,
      mimeType: 'application/vnd.google-apps.folder',
    },
    fields: 'id',
  });
  return folder.data.id;
}

/**
 * 列出用户 Drive 中可用的 XLSX 文件。
 * 包含整个 My Drive、pd-translate 文件夹，以及别人共享给我的文件。
 */
async function listDriveFiles(drive, folderId) {
  const spreadsheetQuery = "(mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel' or mimeType='application/vnd.google-apps.spreadsheet' or name contains '.xlsx' or name contains '.xls') and trashed=false";
  const listOptions = {
    fields: 'files(id, name, mimeType, modifiedTime, createdTime, size, owners, shared)',
    orderBy: 'modifiedTime desc',
    pageSize: 100,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  };

  // 1. 搜索整个 Drive 中的表格文件
  const allSheets = await drive.files.list({
    ...listOptions,
    q: spreadsheetQuery,
    corpora: 'allDrives',
  });

  // 2. 列出自己的 pd-translate 文件夹中的文件，包含文件夹本身用于展示
  const appFolderFiles = await drive.files.list({
    ...listOptions,
    q: `'${folderId}' in parents and trashed=false`,
  });

  // 3. 搜索共享给我的 XLSX 文件
  let sharedFiles = [];
  try {
    const shared = await drive.files.list({
      ...listOptions,
      q: `sharedWithMe=true and ${spreadsheetQuery}`,
    });
    sharedFiles = shared.data.files || [];
  } catch (e) {
    // sharedWithMe 可能在某些配置下失败，吞掉错误继续
    console.warn('Shared files query failed:', e.message);
  }

  // 去重（同 ID 不要出现两次）
  const seen = new Set();
  return [...(allSheets.data.files || []), ...(appFolderFiles.data.files || []), ...sharedFiles]
    .filter(file => {
      if (seen.has(file.id)) return false;
      seen.add(file.id);
      return true;
    });
}

async function listDriveItems(drive, folderId) {
  const parentId = folderId || 'root';
  const res = await drive.files.list({
    q: `'${parentId}' in parents and trashed=false`,
    fields: 'files(id, name, mimeType, modifiedTime, createdTime, size, owners, shared, parents, webViewLink)',
    orderBy: 'folder,name',
    pageSize: 200,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  return (res.data.files || []).map(file => ({
    ...file,
    isFolder: file.mimeType === 'application/vnd.google-apps.folder',
    isSpreadsheet: isSpreadsheetFile(file),
  }));
}

/**
 * 在指定文件夹下查找或创建子文件夹
 */
async function findOrCreateSubFolder(drive, parentId, folderName) {
  const res = await drive.files.list({
    q: `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`,
    fields: 'files(id, name)',
  });
  if (res.data.files.length > 0) return res.data.files[0].id;
  const folder = await drive.files.create({
    requestBody: {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId],
    },
    fields: 'id',
  });
  return folder.data.id;
}

/**
 * 上传文件到 Google Drive
 */
async function uploadFile(drive, parentId, fileName, buffer) {
  const res = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [parentId || 'root'],
    },
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(buffer),
    },
    fields: 'id, name, modifiedTime, size, webViewLink',
  });
  return res.data;
}

async function createFolder(drive, parentId, name) {
  const res = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId || 'root'],
    },
    fields: 'id, name, mimeType, modifiedTime, webViewLink',
    supportsAllDrives: true,
  });
  return res.data;
}

async function renameFile(drive, fileId, name) {
  const res = await drive.files.update({
    fileId,
    supportsAllDrives: true,
    requestBody: { name },
    fields: 'id, name, mimeType, modifiedTime, webViewLink',
  });
  return res.data;
}

async function moveFile(drive, fileId, targetFolderId) {
  const info = await drive.files.get({
    fileId,
    supportsAllDrives: true,
    fields: 'parents',
  });
  const previousParents = (info.data.parents || []).join(',');
  const res = await drive.files.update({
    fileId,
    addParents: targetFolderId || 'root',
    removeParents: previousParents || undefined,
    supportsAllDrives: true,
    fields: 'id, name, mimeType, modifiedTime, parents, webViewLink',
  });
  return res.data;
}

async function trashFile(drive, fileId) {
  const res = await drive.files.update({
    fileId,
    supportsAllDrives: true,
    requestBody: { trashed: true },
    fields: 'id, name, trashed',
  });
  return res.data;
}

/**
 * 更新文件内容
 */
async function updateFileContent(drive, fileId, buffer) {
  const res = await drive.files.update({
    fileId,
    supportsAllDrives: true,
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(buffer),
    },
    fields: 'id, modifiedTime',
  });
  return res.data;
}

/**
 * 下载文件内容（返回 Buffer）
 */
async function downloadFile(drive, fileId) {
  const meta = await getFileInfo(drive, fileId);
  if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
    const exported = await drive.files.export(
      {
        fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      { responseType: 'arraybuffer' }
    );
    return Buffer.from(exported.data);
  }

  const res = await drive.files.get(
    { fileId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(res.data);
}

/**
 * 获取文件元数据
 */
async function getFileInfo(drive, fileId) {
  const res = await drive.files.get({
    fileId,
    supportsAllDrives: true,
    fields: 'id, name, mimeType, modifiedTime, size, owners, parents, webViewLink',
  });
  return res.data;
}

function quoteSheetName(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

function columnName(index) {
  let n = Number(index) + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function hexToColor(hex) {
  const value = String(hex || '').replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(value)) return null;
  return {
    red: parseInt(value.slice(0, 2), 16) / 255,
    green: parseInt(value.slice(2, 4), 16) / 255,
    blue: parseInt(value.slice(4, 6), 16) / 255,
  };
}

function styleToUserEnteredFormat(style) {
  const format = {};
  if (style.bold || style.italic || style.underline || style.strike || style.fontSize || style.fontColor) {
    format.textFormat = {};
    if (style.bold) format.textFormat.bold = true;
    if (style.italic) format.textFormat.italic = true;
    if (style.underline) format.textFormat.underline = true;
    if (style.strike) format.textFormat.strikethrough = true;
    if (style.fontSize) format.textFormat.fontSize = Number(style.fontSize);
    const fg = hexToColor(style.fontColor);
    if (fg) format.textFormat.foregroundColor = fg;
  }
  const bg = hexToColor(style.fillColor);
  if (bg) format.backgroundColor = bg;
  if (style.align) format.horizontalAlignment = String(style.align).toUpperCase();
  return Object.keys(format).length ? format : null;
}

async function updateGoogleSheetValues(sheets, spreadsheetId, state) {
  const data = [];
  const requests = [];
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(sheetId,title))',
  });
  const sheetIds = new Map((meta.data.sheets || []).map(sheet => [sheet.properties.title, sheet.properties.sheetId]));
  for (const sheetName of state.sheetNames || []) {
    const info = state.sheets[sheetName];
    if (!info) continue;
    const sheetId = sheetIds.get(sheetName);
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
      const colName = columnName(col.idx);
      data.push({
        range: `${quoteSheetName(sheetName)}!${colName}2:${colName}${col.values.length + 1}`,
        values: col.values.map(value => [value]),
      });
    }
    if (Number.isInteger(sheetId)) {
      info.rows.forEach((row, rowIndex) => {
        const styleEntries = [
          { idx: info.seqIdx || 0, style: row.styles && row.styles.seq },
          { idx: info.krIdx || 1, style: row.styles && row.styles.kr },
          ...info.headers.map((header, contentIndex) => ({
            idx: header.idx,
            style: row.styles && row.styles.content && row.styles.content[contentIndex],
          })),
        ];
        styleEntries.forEach(entry => {
          const format = styleToUserEnteredFormat(entry.style || {});
          if (!format || !Number.isInteger(entry.idx)) return;
          requests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: rowIndex + 1,
                endRowIndex: rowIndex + 2,
                startColumnIndex: entry.idx,
                endColumnIndex: entry.idx + 1,
              },
              cell: { userEnteredFormat: format },
              fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
            },
          });
        });
      });
    }
  }
  const result = {};
  if (data.length) {
    const res = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data,
      },
    });
    result.values = res.data;
  }
  if (requests.length) {
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
    result.styles = res.data;
  }
  return result;
}

/**
 * 删除文件
 */
async function deleteFile(drive, fileId) {
  await drive.files.delete({ fileId });
}

module.exports = {
  getDriveService,
  getSheetsService,
  findOrCreateRootFolder,
  listDriveFiles,
  listDriveItems,
  findOrCreateSubFolder,
  uploadFile,
  createFolder,
  renameFile,
  moveFile,
  trashFile,
  updateFileContent,
  updateGoogleSheetValues,
  downloadFile,
  getFileInfo,
  deleteFile,
};
