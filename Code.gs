/**
 * One-Forty 夜市撈魚 中文小遊戲 — 後端記錄 API
 *
 * 用途：
 *  - doPost：遊戲結束時，接收一筆遊玩紀錄，寫入「Log」分頁
 *  - doGet：
 *      ?action=leaderboard  → 回傳分數前 10 名（JSON）
 *      不帶 action           → 回傳簡易存活確認（方便測試網址有沒有部署成功）
 *
 * 部署方式請見附上的「部署說明.md」。
 */

const SHEET_LOG_NAME = 'Log';
const SHEET_STATS_NAME = 'Stats';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateLogSheet_();

    sheet.appendRow([
      new Date(),                 // A 時間戳記
      data.deviceId || '',        // B 裝置代碼（匿名）
      data.lang || '',            // C 語言（id/ph/vn）
      data.score || 0,            // D 分數
      data.correctCount || 0,     // E 答對題數
      data.totalAnswered || 0,    // F 已作答題數
      data.durationSec || 0,      // G 花費秒數
      data.result || ''           // H 結果（finished=玩完10題 / lost=生命值歸零）
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'leaderboard') {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: getLeaderboard_() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'One-Forty game API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLeaderboard_() {
  const sheet = getOrCreateLogSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const rows = values.slice(1).map(r => ({
    deviceId: String(r[1] || '').slice(-6), // 只取後 6 碼，避免整串代碼太長
    lang: r[2],
    score: Number(r[3]) || 0,
    durationSec: Number(r[6]) || 0,
    timestamp: r[0]
  }));

  rows.sort((a, b) => b.score - a.score || a.durationSec - b.durationSec);
  return rows.slice(0, 10);
}

function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LOG_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG_NAME);
    sheet.appendRow(['Timestamp', 'DeviceID', 'Lang', 'Score', 'CorrectCount', 'TotalAnswered', 'DurationSec', 'Result']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 選用：在試算表上手動執行一次，會建立/更新「Stats」分頁，
 * 用 QUERY 公式自動算出「獨立玩家數」與「每個裝置玩了幾次」。
 * 之後 Log 分頁有新資料，Stats 分頁的數字會自動更新，不需要重跑。
 */
function setupStatsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let stats = ss.getSheetByName(SHEET_STATS_NAME);
  if (!stats) {
    stats = ss.insertSheet(SHEET_STATS_NAME);
  }
  stats.clear();

  stats.getRange('A1').setValue('總遊玩次數');
  stats.getRange('B1').setFormula(`=COUNTA(${SHEET_LOG_NAME}!B2:B)`);

  stats.getRange('A2').setValue('獨立玩家數（不同裝置）');
  stats.getRange('B2').setFormula(`=SUMPRODUCT((${SHEET_LOG_NAME}!B2:B<>"")/COUNTIF(${SHEET_LOG_NAME}!B2:B,${SHEET_LOG_NAME}!B2:B&""))`);

  stats.getRange('A4').setValue('每個裝置的遊玩次數與最高分');
  stats.getRange('A5').setValue('裝置代碼');
  stats.getRange('B5').setValue('遊玩次數');
  stats.getRange('C5').setValue('最高分');
  stats.getRange('D5').setValue('平均花費秒數');
  stats.getRange('A6').setFormula(
    `=IFERROR(QUERY(${SHEET_LOG_NAME}!B2:G, "select B, count(B), max(D), avg(G) where B is not null group by B order by count(B) desc label count(B) '', max(D) '', avg(G) ''"), "尚無資料")`
  );

  stats.autoResizeColumns(1, 4);
}
