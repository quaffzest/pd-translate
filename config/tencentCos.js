module.exports = {
  secretId: process.env.TENCENT_SECRET_ID,
  secretKey: process.env.TENCENT_SECRET_KEY,
  bucket: process.env.TENCENT_COS_BUCKET,
  region: process.env.TENCENT_COS_REGION,
  prefix: String(process.env.TENCENT_COS_PREFIX || 'pd-translate').replace(/^\/+|\/+$/g, ''),
};
