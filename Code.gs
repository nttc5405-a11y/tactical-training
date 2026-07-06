/**
 * 戰術體能訓練 App － Google Apps Script 後端
 * 功能：接收 index.html 送來的訓練紀錄，寫入本 Google Sheet 的「訓練紀錄」工作表。
 *
 * 使用方式請見同資料夾的「設定說明.md」，這裡不需要修改任何程式碼，
 * 只要把整份程式碼貼到 Apps Script 編輯器並部署成「網頁應用程式」即可。
 */

const SHEET_NAME = '訓練紀錄';
const HEADERS = [
  '上傳時間', '訓練日期', '訓練時間', '單位', '姓名',
  '週次', '訓練', '完成率%', '完成組數', '動作明細', '紀錄ID'
];

/**
 * 接收 index.html 用 fetch POST 送來的資料。
 * body 格式（JSON 字串）：
 * {
 *   unit: '單位名稱',
 *   name: '姓名',
 *   isoTime: 'ISO 時間字串（上傳當下時間，供排序用）',
 *   record: {
 *     id, week, day, date, time, pct, completion, totalSets,
 *     exercises: [{name, note, type, doneSets, totalSets, weight}, ...]
 *   }
 * }
 */
function doPost(e) {
  // 防止多人同時送出資料時互相覆蓋 / 重複寫入
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000); // 最多等 10 秒

  if (!gotLock) {
    return jsonResponse({ status: 'error', message: '系統忙碌中，請稍後再試（LockService 逾時）' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: '沒有收到資料內容（postData 為空）' });
    }

    const payload = JSON.parse(e.postData.contents);
    const unit = payload.unit || '';
    const name = payload.name || '';
    const record = payload.record || {};

    if (!record.id) {
      return jsonResponse({ status: 'error', message: '資料缺少紀錄ID（record.id）' });
    }

    const sheet = getOrCreateSheet();

    // 防重複：同一「紀錄ID＋姓名」已存在就跳過，不重複寫入，但仍回 ok
    if (isDuplicate(sheet, record.id, name)) {
      return jsonResponse({ status: 'ok', message: '此紀錄已存在，略過重複寫入' });
    }

    const uploadTime = payload.isoTime
      ? new Date(payload.isoTime)
      : new Date();

    const dayLabel = (record.day % 2 === 1) ? '第一練' : '第二練';
    const exerciseDetail = buildExerciseDetail(record.exercises);

    sheet.appendRow([
      uploadTime,
      record.date || '',
      record.time || '',
      unit,
      name,
      record.week || '',
      dayLabel,
      record.completion != null ? record.completion : '',
      record.totalSets != null ? record.totalSets : '',
      exerciseDetail,
      record.id
    ]);

    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 把動作明細陣列組成一行好讀的文字。
 * 例："臥推 60kg 4/4組；深蹲 5/5組"（weight 大於 0 才顯示 kg）
 */
function buildExerciseDetail(exercises) {
  if (!exercises || !exercises.length) return '';
  return exercises.map(function (ex) {
    const weightText = (ex.weight && ex.weight > 0) ? (ex.weight + 'kg ') : '';
    return ex.name + ' ' + weightText + ex.doneSets + '/' + ex.totalSets + '組';
  }).join('；');
}

/**
 * 掃描「紀錄ID」欄（第 11 欄）與「姓名」欄（第 5 欄），
 * 若同時符合就視為重複，回傳 true。
 */
function isDuplicate(sheet, recordId, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // 只有表頭，沒有資料

  const idColumn = sheet.getRange(2, 11, lastRow - 1, 1).getValues(); // 紀錄ID
  const nameColumn = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); // 姓名

  for (let i = 0; i < idColumn.length; i++) {
    const existingId = String(idColumn[i][0]);
    const existingName = String(nameColumn[i][0]);
    if (existingId === String(recordId) && existingName === String(name)) {
      return true;
    }
  }
  return false;
}

/**
 * 取得「訓練紀錄」工作表，不存在就建立並寫入表頭。
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    // 工作表存在但是空的，也要補上表頭
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 回傳 JSON 格式的回應給前端 fetch()。
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 可選：用瀏覽器直接打開網頁應用程式網址時的測試回應，
 * 方便你確認部署是否成功（不影響 App 的上傳功能）。
 */
function doGet(e) {
  return ContentService
    .createTextOutput('戰術體能訓練 App 的 Apps Script 已成功部署，可以開始使用。')
    .setMimeType(ContentService.MimeType.TEXT);
}
