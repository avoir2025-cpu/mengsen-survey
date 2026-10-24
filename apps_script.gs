/**
 * 萌森小水手補給站｜問卷後端（Google Apps Script）
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
 * 指令碼屬性（選填）：LINE_CHANNEL_ID ＝ LIFF 所在的 LINE Login Channel ID，
 * 設定後會用 ID Token 驗證送出者真的是該 LINE 帳號。
 */
const SHEET_NAME = '名單';
const TZ = 'Asia/Taipei';
const HEADERS = [
  '送出時間', '紀錄ID', 'LINE userId', 'LINE 名稱', '姓名', '手機', '出生年月日',
  '出遊人數', '國內次數/年', '國外次數/年', '年旅遊花費', '下一趟出發',
  '夢想目的地', '每人預算', '總預算', '透過萌森只需儲值', '個資同意', '兌換時間'
];
const COL = { id: 2, uid: 3, redeemedAt: 18 };

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fmt_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); }
function toIso_(s) { return s ? String(s).replace(' ', 'T') + '+08:00' : null; }
function text_(v) { return "'" + String(v == null ? '' : v); } // 保留手機開頭的 0，不讓試算表自動轉格式

function findRow_(sh, col, value) {
  if (!value || sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

// 回給前端的只有試算結果與兌換狀態，不回傳手機與生日
function publicRecord_(sh, row) {
  const v = sh.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0];
  return {
    id: v[1], uid: v[2],
    people: Number(v[7]), dream: v[12],
    perPerson: Number(v[13]), total: Number(v[14]), need: Number(v[15]),
    save: Number(v[14]) - Number(v[15]),
    redeemedAt: toIso_(v[17])
  };
}

function verifyUid_(idToken) {
  const cid = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  if (!cid) return null; // 沒設定就不驗證
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', payload: { id_token: idToken || '', client_id: cid }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE 身分驗證失敗');
  return JSON.parse(res.getContentText()).sub;
}

function doGet(e) {
  const p = e.parameter || {};
  const sh = sheet_();
  if (p.action === 'get') {
    const row = findRow_(sh, COL.uid, p.uid);
    return json_({ ok: true, record: row ? publicRecord_(sh, row) : null });
  }
  if (p.action === 'stats') {
    const n = Math.max(sh.getLastRow() - 1, 0);
    const redeemed = n ? sh.getRange(2, COL.redeemedAt, n, 1).getValues().filter(r => r[0] !== '').length : 0;
    return json_({ ok: true, submitted: n, redeemed: redeemed });
  }
  return json_({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();

    if (body.action === 'submit') {
      const r = body.record || {};
      const sub = verifyUid_(body.idToken);
      if (sub && sub !== r.uid) return json_({ ok: false, error: 'uid mismatch' });

      const exist = findRow_(sh, COL.uid, r.uid);
      if (exist) return json_({ ok: true, existed: true, record: publicRecord_(sh, exist) });

      sh.appendRow([
        fmt_(new Date(r.submittedAt || Date.now())), r.id, r.uid, r.displayName,
        r.name, text_(r.phone), text_(r.birthday),
        r.people, r.tripsDomestic, r.tripsOverseas, r.spend, r.nextTrip,
        r.dream, r.perPerson, r.total, r.need, r.consent ? '同意' : '',
        r.redeemedAt ? text_(fmt_(new Date(r.redeemedAt))) : ''
      ]);
      return json_({ ok: true, record: publicRecord_(sh, sh.getLastRow()) });
    }

    if (body.action === 'redeem') {
      const row = findRow_(sh, COL.id, body.id);
      if (!row) return json_({ ok: false, error: 'not found' });
      const cell = sh.getRange(row, COL.redeemedAt);
      if (!cell.getDisplayValue()) cell.setValue(text_(fmt_(new Date(body.redeemedAt || Date.now()))));
      return json_({ ok: true, redeemedAt: toIso_(cell.getDisplayValue()) });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
