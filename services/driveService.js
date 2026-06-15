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
  const xlsxQuery = `(mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or name contains '.xlsx' or name contains '.xls') and trashed=false`;

  // 1. 搜索整个 Drive 中的 XLSX 文件
  const allSheets = await drive.files.list({
    q: xlsxQuery,
    fields: 'files(id, name, mimeType, modifiedTime, createdTime, size, owners, shared)',
    orderBy: 'modifiedTime desc',
    pageSize: 100,
  });

  // 2. 列出自己的 pd-translate 文件夹中的文件，包含文件夹本身用于展示
  const appFolderFiles = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false`,
    fields: 'files(id, name, mimeType, modifiedTime, createdTime, size, owners, shared)',
    orderBy: 'modifiedTime desc',
    pageSize: 100,
  });

  // 3. 搜索共享给我的 XLSX 文件
  let sharedFiles = [];
  try {
    const shared = await drive.files.list({
      q: `sharedWithMe=true and ${xlsxQuery}`,
      fields: 'files(id, name, mimeType, modifiedTime, createdTime, size, owners, shared)',
      orderBy: 'modifiedTime desc',
      pageSize: 100,
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
      parents: [parentId],
    },
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(buffer),
    },
    fields: 'id, name, modifiedTime, size',
  });
  return res.data;
}

/**
 * 更新文件内容
 */
async function updateFileContent(drive, fileId, buffer) {
  const res = await drive.files.update({
    fileId,
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
  const res = await drive.files.get(
    { fileId, alt: 'media' },
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
    fields: 'id, name, mimeType, modifiedTime, size, owners',
  });
  return res.data;
}

/**
 * 删除文件
 */
async function deleteFile(drive, fileId) {
  await drive.files.delete({ fileId });
}

module.exports = {
  getDriveService,
  findOrCreateRootFolder,
  listDriveFiles,
  findOrCreateSubFolder,
  uploadFile,
  updateFileContent,
  downloadFile,
  getFileInfo,
  deleteFile,
};
