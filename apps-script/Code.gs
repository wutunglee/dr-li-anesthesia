/**
 * 黎乃良醫師網站｜預約洽詢表單
 * 網站表單送出 → 存進這份試算表 → 寄通知信給黎醫師 →（可選）寄確認信給填表人
 *
 * 設定只需要改下面三行。
 */
const NOTIFY_EMAIL = 'reneelnl@gmail.com'; // 收通知信的信箱
const AUTO_REPLY   = true;                  // 是否寄確認信給填表人（true 寄／false 不寄）
const REPLY_DAYS   = 3;                     // 確認信裡寫「幾個工作天內回覆」

const SHEET_NAME = '表單回覆';
const HEADERS = ['時間', '類型', '姓名', 'Email', '電話', '診所／單位', '地點',
                 '手術或處置類型', '希望時段／頻率', '預計手術日期', '想詢問的內容'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // 防機器人：網頁上看不到的欄位被填了，就當作沒收到
  if (p.website) return json_({ ok: true });

  const type = p.type === 'clinic' ? '診所合作' : '個人諮詢';
  if (!p.name || !p.email || p.consent !== 'yes') return json_({ ok: false, error: 'missing' });
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(p.email)) return json_({ ok: false, error: 'email' });

  const values = [type, p.name, p.email, p.phone, p.org, p.location,
                  p.procedure, p.schedule, p.date, p.message].map(clean_);

  // 1. 存進試算表
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date()].concat(values));
  } finally {
    lock.releaseLock();
  }

  // 2. 寄通知信給黎醫師（直接按「回覆」就會回給填表人）
  const lines = HEADERS.slice(1)
    .map((h, i) => h + '：' + (values[i] || '—'))
    .join('\n');
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: p.email,
    subject: '【網站' + type + '】' + p.name,
    body: lines + '\n\n——\n直接回覆這封信，就會寄給對方。'
  });

  // 3. 寄確認信給填表人
  if (AUTO_REPLY) {
    MailApp.sendEmail({
      to: p.email,
      name: '黎乃良醫師',
      replyTo: NOTIFY_EMAIL,
      subject: '已收到您的' + type + '｜黎乃良醫師',
      body: p.name + ' 您好：\n\n' +
            '已收到您透過網站送出的' + type + '，我會在 ' + REPLY_DAYS + ' 個工作天內回覆您。\n\n' +
            '本信由系統自動寄出。網站諮詢僅提供一般衛教說明與預約安排，不能取代實際的術前評估與診療。\n\n' +
            '黎乃良 麻醉專科醫師'
    });
  }

  return json_({ ok: true });
}

// 用瀏覽器打開網頁應用程式網址時顯示，確認有部署成功
function doGet() {
  return ContentService.createTextOutput('黎醫師網站表單：運作中');
}

// 安裝後可手動執行一次，確認收得到信
function testEmail() {
  MailApp.sendEmail(NOTIFY_EMAIL, '【測試】黎醫師網站表單', '收到這封信，代表通知信設定成功。');
}

function clean_(v) {
  v = String(v == null ? '' : v).trim().slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // 避免被試算表當成公式
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
