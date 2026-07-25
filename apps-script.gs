// あそびばマップ フェーズ2 - Google Apps Script（記録の同期用）
// 使い方: 新しい Google スプレッドシートを開き、拡張機能 > Apps Script に
// このコードを全部貼り付けて保存 → デプロイ > 新しいデプロイ > 種類「ウェブアプリ」
//   実行するユーザー: 自分 / アクセスできるユーザー: 全員 → デプロイ
// で発行される /exec の URL を Claude に渡してください。
//
// このスクリプトは「records」シートに、あそびばマップの記録
// （行った/行きたい/候補・メモ）を1行1スポットで保存します。
// カタログ（遊び場マスタ）はサイト側が持つので、ここには記録だけが入ります。

const SHEET_NAME = 'records';
const WRITE_TOKEN = 'ufSiDgxVTDE9D5_YR-uiaUXYVP0rs2TT';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;
  try {
    if (action === 'getRecords') {
      result = getRecords();
    } else if (action === 'setRecord') {
      if (e.parameter.token !== WRITE_TOKEN) return forbidden();
      result = setRecord(e.parameter);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['spotId', 'status', 'note', 'visitedAt', 'updatedAt']);
    // 日時・IDは文字列として保持（ISO文字列が日付に自動変換されるのを防ぐ）
    sh.getRange('A:E').setNumberFormat('@');
  }
  return sh;
}

function getRecords() {
  const sh = getSheet();
  const last = sh.getLastRow();
  const out = {};
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, 5).getValues();
    rows.forEach((r) => {
      const id = String(r[0] || '').trim();
      if (!id) return;
      out[id] = {
        status: String(r[1] || 'koho'),
        note: String(r[2] || ''),
        visitedAt: r[3] ? String(r[3]) : null,
        updatedAt: r[4] ? String(r[4]) : null,
      };
    });
  }
  return { records: out };
}

function setRecord(p) {
  const id = String(p.id || '').trim();
  if (!id) return { error: 'id required' };
  const sh = getSheet();
  const row = [
    id,
    String(p.status || 'koho'),
    String(p.note || ''),
    String(p.visitedAt || ''),
    String(p.updatedAt || new Date().toISOString()),
  ];
  const last = sh.getLastRow();
  let rowIdx = -1;
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) { rowIdx = i + 2; break; }
    }
  }
  if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, 5).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return { success: true };
}

function forbidden() {
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Forbidden' }))
    .setMimeType(ContentService.MimeType.JSON);
}
