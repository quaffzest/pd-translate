// config/google.js
module.exports = {
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: process.env.GOOGLE_CALLBACK_URL || ((process.env.RENDER_EXTERNAL_URL || `http://localhost:${process.env.PORT || 3000}`) + '/auth/google/callback'),
  // Google Drive 中存储文件的根文件夹名
  driveRootFolder: 'pd-translate',
};
