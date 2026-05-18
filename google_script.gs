/**
 * 🛡️ 兄弟聯盟 ONLINE — Google Apps Script 雲端後台 (能量寶石版)
 *
 * ──────────────────────────────────────────────────
 * 部署步驟：
 * 1. 貼入此程式碼，修改 CONFIG 區的 FATHER_EMAIL 為你的 Gmail 信箱。
 * 2. 修改 LINE_CONFIG 填入你的 Channel Access Token 與 User ID。
 * 3. 點擊「部署」→「新增部署」→ 類型選「Web 應用程式」。
 * - 執行身份：選「我」
 * - 存取權：選「所有人」
 * 4. 將產生的 URL 貼回 HTML 中的上傳函式內。
 * ──────────────────────────────────────────────────
 */

// ══ 父親設定區（請修改這裡）════════════════════════
var CONFIG = {
  FATHER_EMAIL: 'dad@example.com',     // 爸爸的 Gmail，戰報寄到這裡
  SHEET_NAME:   '能量寶石戰報'           // 試算表分頁名稱
};

// ══ LINE Messaging API 設定（請填入你的憑證）══════
var LINE_CONFIG = {
  CHANNEL_ACCESS_TOKEN: 'YOUR_CHANNEL_ACCESS_TOKEN_HERE',  
  USER_ID:              'YOUR_LINE_USER_ID_HERE',           
};

// ══ 接收 POST 請求 ════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    sendEmailToFather(data);
    sendLinePush(data);        
    return respond({ status: 'ok', message: '同步成功！' });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

function doGet() {
  return respond({ status: 'ok', message: '兄弟聯盟後台運作正常 🛡️' });
}

// ══ 寫入 Google 試算表════════════════════════════════
function saveToSheet(d) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    var headers = [
      '日期', '團隊連鎖', 
      '哥哥-書名', '哥哥-分鐘', '哥哥-運動項目', '哥哥-步數', '哥哥-感恩與突破', '哥哥-重點修煉', '哥哥獲得寶石', '哥哥任務數',
      '弟弟-書名', '弟弟-分鐘', '弟弟-運動項目', '弟弟-步數', '弟弟-感恩與突破', '弟弟-重點修煉', '弟弟獲得寶石', '弟弟任務數',
      '兄弟連攜加成', '今日總寶石'
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#1e293b')
         .setFontColor('#10b981');
  }

  var p1 = d.p1 || {};
  var p2 = d.p2 || {};

  sheet.appendRow([
    d.date        || new Date().toLocaleDateString('zh-TW'),
    d.teamStreak  || 1,

    p1.book       || '',
    p1.mins       || '',
    p1.sport      || '',
    p1.steps      || '',
    p1.gratitude  || '',
    p1.challenge  || '',
    p1.gems       || 0,
    p1.tasksDone  || 0,

    p2.book       || '',
    p2.mins       || '',
    p2.sport      || '',
    p2.steps      || '',
    p2.gratitude  || '',
    p2.challenge  || '',
    p2.gems       || 0,
    p2.tasksDone  || 0,

    d.allianceBonus ? '🔥 +20×2 💎' : '—',
    d.totalGems   || 0
  ]);
}

// ══ 寄送 Email 戰報給爸爸 ══════════════════════════
function sendEmailToFather(d) {
  var to = d.fatherEmail || CONFIG.FATHER_EMAIL;
  if (!to || to === 'dad@example.com') return; 

  var p1 = d.p1 || {};
  var p2 = d.p2 || {};

  var subject = '🛡️ 兄弟聯盟：今日成長戰報' + (d.allianceBonus ? '【🔥 連攜加成解鎖！】' : '');
  var body = buildHtmlEmail(d, p1, p2);

  MailApp.sendEmail({
    to:       to,
    subject:  subject,
    htmlBody: body,
  });
}

// ══ HTML Email 範本═══════════════════════════════════
function buildHtmlEmail(d, p1, p2) {
  var bonusBanner = d.allianceBonus
    ? '<div style="background:#10b981;color:#fff;padding:12px 20px;border-radius:10px;font-weight:bold;font-size:1.05em;margin-bottom:20px;text-align:center">🔥 兄弟連攜加成解鎖！雙方各獲得 +20 能量寶石 💎</div>'
    : '';

  function txt(s, fb) { return (s && s.trim()) ? s : '<span style="color:#94a3b8">（未填寫）</span>'; }

  return [
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>',
    '<body style="font-family:sans-serif;background:#f8fafc;margin:0;padding:20px">',
    '<div style="max-width:600px;margin:auto;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.05);border:1px solid #e2e8f0">',

    // Header
    '<div style="background:#ffffff;padding:28px;text-align:center;border-bottom:1px solid #e2e8f0">',
    '<h1 style="color:#3b82f6;margin:0;font-size:1.6em;">🛡️ 兄弟聯盟 ONLINE</h1>',
    '<p style="color:#64748b;margin:8px 0 0;font-size:0.9em">📅 ' + (d.date||'') + ' · 🔥 連續登入 ' + (d.teamStreak||1) + ' 天</p>',
    '</div>',

    // Body
    '<div style="padding:24px">',
    bonusBanner,

    // 哥哥 (策略弓箭手)
    '<div style="background:#eff6ff;border-radius:12px;padding:18px;margin-bottom:18px;border-top:4px solid #3b82f6">',
    '<h2 style="color:#3b82f6;margin:0 0 12px;font-size:1.1em">🧙‍♂️ 哥哥 (策略弓箭手) · ' + (p1.gems||0) + ' 💎 · ' + (p1.tasksDone||0) + '/7 任務</h2>',
    '<ul style="margin:0;padding-left:20px;line-height:1.8;color:#334155;font-size:0.95em">',
    '<li>📖 閱讀：' + txt(p1.book) + ' (' + txt(p1.mins) + ' 分鐘)</li>',
    '<li>🏃 運動：' + txt(p1.sport) + ' (' + txt(p1.steps) + ' 步)</li>',
    '<li>✏️ 感恩：' + txt(p1.gratitude) + '</li>',
    '<li>⚔️ 修煉：' + txt(p1.challenge) + '</li>',
    '</ul>',
    '</div>',

    // 弟弟 (閃電劍客)
    '<div style="background:#fff7ed;border-radius:12px;padding:18px;margin-bottom:18px;border-top:4px solid #f97316">',
    '<h2 style="color:#f97316;margin:0 0 12px;font-size:1.1em">⚡ 弟弟 (閃電劍客) · ' + (p2.gems||0) + ' 💎 · ' + (p2.tasksDone||0) + '/7 任務</h2>',
    '<ul style="margin:0;padding-left:20px;line-height:1.8;color:#334155;font-size:0.95em">',
    '<li>📖 閱讀：' + txt(p2.book) + ' (' + txt(p2.mins) + ' 分鐘)</li>',
    '<li>🏃 運動：' + txt(p2.sport) + ' (' + txt(p2.steps) + ' 步)</li>',
    '<li>✏️ 感恩：' + txt(p2.gratitude) + '</li>',
    '<li>⚔️ 修煉：' + txt(p2.challenge) + '</li>',
    '</ul>',
    '</div>',

    // 總覽
    '<div style="background:#f1f5f9;border-radius:10px;padding:16px;text-align:center">',
    '<p style="margin:0;font-size:1.1em;font-weight:bold;color:#1e293b">',
    '今日總計獲得：<span style="color:#10b981">' + (d.totalGems||0) + ' 💎 能量寶石</span></p>',
    '</div>',

    '</div></div></body></html>',
  ].join('');
}

// ══ LINE Messaging API 推播 ═══════════════════════
function sendLinePush(d) {
  var token  = LINE_CONFIG.CHANNEL_ACCESS_TOKEN;
  var userId = LINE_CONFIG.USER_ID;
  if (!token || token.indexOf('YOUR_') === 0) return;
  if (!userId || userId.indexOf('YOUR_') === 0) return;

  var text = buildLineMessage(d);

  var options = {
    method:      'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      to:       userId,
      messages: [{ type: 'text', text: text }],
    }),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
}

function buildLineMessage(d) {
  var p1 = d.p1 || {};
  var p2 = d.p2 || {};
  function txt(s, fb) { return (s && s.trim()) ? s.trim() : (fb || '—'); }

  var bonusLine = d.allianceBonus ? '\n🔥 連攜大絕啟動！雙方各 +20 💎！' : '';

  var lines = [
    '╔═══════════════════════╗',
    '  🛡️ 兄弟聯盟 ONLINE',
    '  今日成長戰報',
    '╚═══════════════════════╝',
    '📅 ' + (d.date||'') + '　🔥 連鎖 ' + (d.teamStreak||1) + ' 天',
    '',
    '━━━━ 🧙‍♂️ 哥哥 (' + (p1.tasksDone||0) + '/7) ━━━━',
    '📖 書籍：' + txt(p1.book, '未填寫'),
    '✏️ 感恩：' + txt(p1.gratitude, '未填寫'),
    '⚔️ 修煉：' + txt(p1.challenge, '未填寫'),
    '💎 獲得寶石：' + (p1.gems||0) + ' 💎',
    '',
    '━━━━ ⚡ 弟弟 (' + (p2.tasksDone||0) + '/7) ━━━━',
    '📖 書籍：' + txt(p2.book, '未填寫'),
    '✏️ 感恩：' + txt(p2.gratitude, '未填寫'),
    '⚔️ 修煉：' + txt(p2.challenge, '未填寫'),
    '💎 獲得寶石：' + (p2.gems||0) + ' 💎',
    '',
    '━━━━━━ 聯盟總結 ━━━━━━',
    bonusLine,
    '💎 總共獲得：' + (d.totalGems||0) + ' 能量寶石',
    '',
    '💙 爸爸快去準備特權兌換吧！',
  ];
  return lines.join('\n');
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}