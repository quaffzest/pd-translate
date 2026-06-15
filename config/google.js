// config/google.js
module.exports = {
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: (process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000') + '/auth/google/callback',
  // Google Drive 中存储文件的根文件夹名
  driveRootFolder: 'pd-translate',
};
