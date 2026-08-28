/**
 * GAO卓球場 無料予約システム backend.gs
 * Google Apps Script + Google Sheets
 *
 * Sheet name: Bookings
 * Columns:
 * A createdAt | B date | C hour | D kind | E name | F email | G member
 * H people | I note | J status
 */
const SHEET_NAME = 'Bookings';
const MAX_TABLES = 1;
const CALENDAR_ID = ''; // 任意：Google Calendarへ入れる場合はカレンダーIDを設定

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  if (action === 'availability') {
    const date = e.parameter.date;
    const hour = Number(e.parameter.hour);
    return json_({ ok:true, count: getBookingCount_(date, hour) });
  }
  return json_({ ok:true, service:'GAO booking API' });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    if (p.action !== 'book') return json_({ok:false,message:'Invalid action'});
    validate_(p);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const count = getBookingCount_(p.date, Number(p.hour));
      if (count >= MAX_TABLES) return json_({ok:false,message:'この時間は満席です。'});

      const hour = Number(p.hour);
      if (p.kind === 'table' && (hour < 10 || hour >= 17)) {
        return json_({ok:false,message:'台貸しは原則10:00〜17:00です。'});
      }

      const sh = getSheet_();
      sh.appendRow([
        new Date(), p.date, hour, p.kind || 'table', p.name, p.email,
        p.member || '', Number(p.people || 1), p.note || '', 'confirmed'
      ]);

      if (CALENDAR_ID) createCalendarEvent_(p);
      sendConfirmation_(p);
      return json_({ok:true});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ok:false,message:String(err.message || err)});
  }
}

function getBookingCount_(date, hour) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const values = sh.getRange(2,1,last-1,10).getValues();
  return values.filter(r =>
    String(r[1]) === String(date) &&
    Number(r[2]) === Number(hour) &&
    String(r[3]) === 'table' &&
    String(r[9]).toLowerCase() !== 'cancelled'
  ).length;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['createdAt','date','hour','kind','name','email','member','people','note','status']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function validate_(p) {
  if (!p.date) throw new Error('日付を選択してください。');
  if (p.hour === undefined || p.hour === null) throw new Error('時間を選択してください。');
  if (!p.name) throw new Error('お名前を入力してください。');
  if (!p.email || !String(p.email).includes('@')) throw new Error('メールアドレスを確認してください。');
}

function isTomorrowOrLater_(dateStr) {
  const now = new Date();
  const d = new Date(dateStr + 'T00:00:00');
  const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate()+1);
  tomorrow.setHours(0,0,0,0);
  return d >= tomorrow;
}

function createCalendarEvent_(p) {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) return;
  const start = new Date(p.date + 'T' + String(p.hour).padStart(2,'0') + ':00:00');
  const end = new Date(start.getTime() + 60*60*1000);
  cal.createEvent(
    `GAO予約：${p.name}（${p.kind === 'lesson' ? '個人レッスン' : '台貸し'}）`,
    start, end,
    {description:`会員区分: ${p.member}\n人数: ${p.people}\nEmail: ${p.email}\n備考: ${p.note || ''}`}
  );
}

function sendConfirmation_(p) {
  const subject = `【GAO卓球場】予約確認 ${p.date} ${String(p.hour).padStart(2,'0')}:00`;
  const body =
`${p.name} 様

GAO卓球場のご予約ありがとうございます。

日付：${p.date}
時間：${String(p.hour).padStart(2,'0')}:00
内容：${p.kind === 'lesson' ? '個人レッスン' : '台貸し'}
人数：${p.people}
区分：${p.member}

〒194-0045
東京都町田市南成瀬1丁目8-12 中山ビル3階301
TEL：070-2178-6868

GAO卓球場`;
  MailApp.sendEmail(p.email, subject, body);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
