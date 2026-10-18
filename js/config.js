// Non-sensitive application configuration（此檔公開，請勿放 API 金鑰）
// Google Books 金鑰改存在 Apps Script 的「指令碼屬性」，見 apps-script/Code.gs
// Stage 6–8: 部署 apps-script/Code.gs 後，將 Google Apps Script Web App URL 填入 APPS_SCRIPT_URL
const CONFIG = Object.freeze({
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbw95crT43Q4UO4HWp22CZ9LaiVeu4cTa0ZzIUh7-jRDBghLwB5jt2ee5ZvbeP3XY6qm/exec',
  READERS: ['妍皮', '刺球', 'test'],
  OFFLINE_QUEUE_KEY: 'readingOfflineQueue',
});
