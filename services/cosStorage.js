const {
  DeleteObjectCommand,
  GetObjectCommand,
  ListObjectsV2Command,
  PutObjectCommand,
  S3Client,
} = require('@aws-sdk/client-s3');
const cosConfig = require('../config/tencentCos');

function isConfigured() {
  return !!(cosConfig.secretId && cosConfig.secretKey && cosConfig.bucket && cosConfig.region);
}

function client() {
  if (!isConfigured()) return null;
  return new S3Client({
    region: cosConfig.region,
    endpoint: `https://cos.${cosConfig.region}.myqcloud.com`,
    credentials: {
      accessKeyId: cosConfig.secretId,
      secretAccessKey: cosConfig.secretKey,
    },
  });
}

function normalizeKey(key) {
  return String(key || '').replace(/\\/g, '/').replace(/^\/+/, '');
}

function key(relativeKey) {
  const clean = normalizeKey(relativeKey);
  return cosConfig.prefix ? `${cosConfig.prefix}/${clean}` : clean;
}

function unkey(fullKey) {
  const prefix = cosConfig.prefix ? `${cosConfig.prefix}/` : '';
  return String(fullKey || '').startsWith(prefix) ? String(fullKey).slice(prefix.length) : String(fullKey || '');
}

async function streamToBuffer(stream) {
  if (Buffer.isBuffer(stream)) return stream;
  const chunks = [];
  for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  return Buffer.concat(chunks);
}

async function putObject(relativeKey, body, contentType) {
  const s3 = client();
  if (!s3) return null;
  return s3.send(new PutObjectCommand({
    Bucket: cosConfig.bucket,
    Key: key(relativeKey),
    Body: body,
    ContentType: contentType || 'application/octet-stream',
  }));
}

async function getObject(relativeKey) {
  const s3 = client();
  if (!s3) return null;
  const data = await s3.send(new GetObjectCommand({
    Bucket: cosConfig.bucket,
    Key: key(relativeKey),
  }));
  return streamToBuffer(data.Body);
}

async function deleteObject(relativeKey) {
  const s3 = client();
  if (!s3) return null;
  return s3.send(new DeleteObjectCommand({
    Bucket: cosConfig.bucket,
    Key: key(relativeKey),
  }));
}

async function listObjects(relativePrefix) {
  const s3 = client();
  if (!s3) return [];
  const prefix = key(normalizeKey(relativePrefix || ''));
  const fullPrefix = prefix && !prefix.endsWith('/') ? `${prefix}/` : prefix;
  const all = [];
  let token;
  do {
    const data = await s3.send(new ListObjectsV2Command({
      Bucket: cosConfig.bucket,
      Prefix: fullPrefix,
      ContinuationToken: token,
      MaxKeys: 1000,
    }));
    all.push(...(data.Contents || []).map(item => ({
      key: unkey(item.Key),
      size: Number(item.Size || 0),
      lastModified: item.LastModified,
      etag: item.ETag,
    })));
    token = data.NextContinuationToken;
  } while (token);
  return all;
}

async function deletePrefix(relativePrefix) {
  if (!isConfigured()) return null;
  const objects = await listObjects(relativePrefix);
  await Promise.all(objects.map(item => deleteObject(item.key)));
  return objects.length;
}

module.exports = {
  isConfigured,
  putObject,
  getObject,
  deleteObject,
  deletePrefix,
  listObjects,
};
