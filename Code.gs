// 怪我管理システム - 共有スプレッドシート用バックエンド
//
// セットアップ:
// 1. Google スプレッドシートを新規作成する
// 2. 「拡張機能」>「Apps Script」を開き、デフォルトのコードを全て削除してこのファイルの内容を貼り付ける
// 3. 「デプロイ」>「新しいデプロイ」>種類「ウェブアプリ」を選択
//    - 実行するユーザー: 自分
//    - アクセスできるユーザー: 全員
//      (「全員」にしないとアプリからログインなしで読み書きできません。
//       このURLを知っていれば誰でもデータを読み書きできるため、URLは第三者に共有しないでください)
// 4. デプロイ後に表示される「ウェブアプリ URL」をコピーし、
//    injury-tracker アプリの「共有設定」欄に貼り付けて「接続」する
// 5. コードを更新した場合は「デプロイを管理」から新バージョンを配置し直す

const SHEET_NAME = '怪我記録';
const HEADERS = ['id', '名前', '診断名', '受傷日', '区分', 'コンタクト', '復帰日'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function formatDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '');
}

function readAll_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const rows = values.slice(1);
  return rows
    .filter(row => row[0] !== '')
    .map(row => {
      const rec = {};
      header.forEach((key, i) => { rec[key] = row[i]; });
      const returnDate = formatDate_(rec['復帰日']);
      return {
        id: String(rec['id']),
        name: rec['名前'],
        diagnosis: rec['診断名'],
        injuryDate: formatDate_(rec['受傷日']),
        category: rec['区分'],
        contact: rec['コンタクト'],
        returnDate: returnDate === '' ? null : returnDate,
      };
    });
}

function findRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-indexed row number
  }
  return -1;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return jsonOutput_(readAll_());
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const record = payload.record || {};
    const sheet = getSheet_();

    if (action === 'add') {
      const id = Utilities.getUuid();
      sheet.appendRow([
        id,
        record.name || '',
        record.diagnosis || '',
        record.injuryDate || '',
        record.category || '',
        record.contact || '',
        record.returnDate === null || record.returnDate === undefined ? '' : record.returnDate,
      ]);
    } else if (action === 'update') {
      const rowNum = findRowById_(sheet, record.id);
      if (rowNum !== -1) {
        const colMap = { name: 2, diagnosis: 3, injuryDate: 4, category: 5, contact: 6, returnDate: 7 };
        Object.keys(colMap).forEach(key => {
          if (Object.prototype.hasOwnProperty.call(record, key)) {
            let v = record[key];
            if (key === 'returnDate' && (v === null || v === undefined)) v = '';
            sheet.getRange(rowNum, colMap[key]).setValue(v);
          }
        });
      }
    } else if (action === 'delete') {
      const rowNum = findRowById_(sheet, record.id);
      if (rowNum !== -1) sheet.deleteRow(rowNum);
    }

    return jsonOutput_(readAll_());
  } finally {
    lock.releaseLock();
  }
}
