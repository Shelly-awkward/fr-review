/**
 * fr-review 寄送端點（Google Apps Script）
 *
 * 網頁產出 Excel／Word 後把檔案 POST 到這裡，本程式驗密碼、用你自己的 Gmail 寄出。
 * 密碼與收件人都存在「指令碼屬性」裡，不會出現在網頁原始碼——這是靜態網頁做不到、
 * 而必須由伺服器端把關的部分。
 *
 * ── 部署步驟（做一次，約五分鐘）──────────────────────────
 * 1. 開 https://script.google.com → 新增專案，把本檔內容整份貼上。
 * 2. 左側「專案設定」→ 指令碼屬性 → 新增兩筆：
 *      PASSWORD   你要發給同事的密碼
 *      RECIPIENT  收件 Gmail（多個以逗號分隔）
 * 3. 右上「部署」→ 新增部署作業 → 類型選「網頁應用程式」
 *      執行身分：我
 *      具有存取權的使用者：<b>任何人</b>（同事沒有 Google 帳號也能用）
 * 4. 複製部署後的網址（https://script.google.com/macros/s/…/exec），
 *    填進網頁的「寄送設定」欄位（或寫進 index.html 的 MAILER_URL 常數）。
 *
 * ── 注意 ────────────────────────────────────────────
 * ・Gmail 每日寄信量有配額（一般帳號約 100 封／日），內部使用綽綽有餘。
 * ・「具有存取權：任何人」意謂知道網址的人都能呼叫，所以密碼是唯一的門——
 *   請透過內部管道發送，不要寫在公開網頁上。
 * ・本程式只寄信，不儲存任何檔案；附件用完即拋。
 */

function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty('PASSWORD');
  const recipient = props.getProperty('RECIPIENT');

  try {
    if (!password || !recipient) {
      return json({ ok: false, error: '端點尚未設定 PASSWORD／RECIPIENT 指令碼屬性' });
    }
    const req = JSON.parse(e.postData.contents);
    if (req.password !== password) {
      return json({ ok: false, error: '密碼錯誤' });
    }
    if (!req.files || !req.files.length) {
      return json({ ok: false, error: '沒有附件' });
    }

    const attachments = req.files.map(function (f) {
      return Utilities.newBlob(
        Utilities.base64Decode(f.data),
        f.mime || 'application/octet-stream',
        f.name);
    });

    const subject = '【財報實審】' + (req.subject || '產出文件');
    const body =
      '本信由財報實審文件產生器自動寄出。\n\n' +
      '公司：' + (req.company || '－') + '\n' +
      '年度：' + (req.year || '－') + '\n' +
      '產出時間：' + new Date().toLocaleString('zh-TW') + '\n' +
      '操作人備註：' + (req.note || '－') + '\n\n' +
      '附件為初稿：Word 管區意見中標示「擬行前查證」者須於行前補實，\n' +
      '財報頁碼等 XBRL 未提供之資訊請自行補註。判斷責任在檢查員。';

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
      attachments: attachments,
    });
    return json({ ok: true, sent: attachments.length, to: recipient });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, service: 'fr-review mailer', hint: '請以 POST 呼叫' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
