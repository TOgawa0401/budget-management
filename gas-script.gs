/**
 * データ推進部 2026年度予算管理 - Google Apps Script
 * このスクリプトはGoogleスプレッドシートのスクリプトエディタにコピペしてください
 */

const SHEET_NAME = '予算データ';
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * 初期化：スプレッドシートにヘッダーと初期データを設定
 * 初回のみ手動実行してください（スクリプトエディタの▶実行ボタン）
 */
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  }

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // ヘッダー行
  const headers = ['id', 'category', 'name', 'budget26', 'actual', 'forecast', 'responsible', 'executionDate', 'details'];
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 初期データ
  const initialData = [
    [1, '予算化項目', 'LineWorks 契約費用', 2160000, 2160000, 0, '', '', '216万で400アカウントで契約。2026年4月契約'],
    [2, '予算化項目', 'LineWorks 追加100アカウント', 0, 540000, 0, '', '', '100アカウント追加54万'],
    [3, '予算化項目', 'アルバイト募集有料施策（茨城大）', 125000, 8800, 0, '', '', '茨城大に募集広告掲載'],
    [4, '予算化項目', 'アルバイト募集有料施策（静岡大）', 125000, 5500, 0, '', '', '静岡大に募集広告掲載'],
    [5, '予算化項目', 'アルバイト募集有料施策（広島大）', 125000, 3300, 0, '', '', '広島大に募集広告掲載'],
    [6, '予算化項目', 'アルバイト募集有料施策（長崎大）', 125000, 0, 0, '', '', '長崎大に募集広告掲載（二重広告掲載）'],
    [7, '予算化項目', 'AI有料アカウント契約料', 400000, 350000, 0, '', '', 'Claude 8名のプロプランを契約'],
    [8, '予算化項目', 'ispring契約費用', 390000, 0, 390000, '', '', '2025年11月に契約更新。次回更新は2026年11月。更新は要検討'],
    [9, '予算化項目', '野球エキスパート慰労会費用', 200000, 0, 200000, '', '', ''],
    [10, '予算化項目', 'バスケエキスパート慰労会費用', 200000, 0, 200000, '', '', ''],
    [11, '予算化項目', 'FBEXP経費精算システム', 180000, 180000, 0, '', '', ''],
    [12, '予算化項目', 'FBEXP納会費用', 180000, 0, 180000, '', '', ''],
    [13, '予算化項目', '書籍雑誌購入費用', 72000, 50000, 0, '', '', ''],
    [14, '予算化項目', '入力室備品', 50000, 0, 0, '', '', '未執行'],
    [15, '予算化項目', 'サッカー映像保管用HDD購入費', 50000, 0, 50000, '', '', ''],
    [16, '予算化項目', '卓球エキスパート慰労会費用', 50000, 0, 50000, '', '', ''],
    [17, '予算化項目', 'DS Football Twitter(X)アカウント有料プランサブスク料', 9180, 0, 9180, '', '', ''],
    [18, '追加項目', 'PowerAutomate有料契約', 49464, 49464, 0, '', '', '']
  ];

  targetSheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);

  SpreadsheetApp.flush();
  Logger.log('初期化完了');
}

/**
 * GETリクエスト：全データ取得
 */
function doGet(e) {
  const action = e.parameter.action;

  try {
    if (action === 'getAll') {
      return getAllData();
    } else {
      return errorResponse('Invalid action');
    }
  } catch (error) {
    return errorResponse(error.toString());
  }
}

/**
 * POSTリクエスト：データ追加・更新・削除
 */
function doPost(e) {
  const action = e.parameter.action;
  const payload = JSON.parse(e.postData.contents);

  try {
    switch (action) {
      case 'add':
        return addData(payload);
      case 'update':
        return updateData(payload);
      case 'delete':
        return deleteData(payload);
      default:
        return errorResponse('Invalid action');
    }
  } catch (error) {
    return errorResponse(error.toString());
  }
}

/**
 * 全データ取得
 */
function getAllData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return successResponse([]);
  }

  const headers = data[0];
  const rows = data.slice(1);

  const cases = rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  }).filter(obj => obj.id); // idがあるもののみ

  return successResponse(cases);
}

/**
 * データ追加
 */
function addData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  // 新しいIDは最後のID + 1
  const allData = sheet.getDataRange().getValues();
  const maxId = Math.max(...allData.slice(1).map(row => row[0]));
  const newId = maxId + 1;

  const newRow = [
    newId,
    payload.category || '予算化項目',
    payload.name,
    payload.budget26 || 0,
    payload.actual || 0,
    payload.forecast || 0,
    payload.responsible || '',
    payload.executionDate || '',
    payload.details || ''
  ];

  sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
  SpreadsheetApp.flush();

  return successResponse({ id: newId, ...payload });
}

/**
 * データ更新
 */
function updateData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const targetRowIndex = data.findIndex((row, index) => index > 0 && row[0] === payload.id);

  if (targetRowIndex === -1) {
    return errorResponse('Data not found');
  }

  const rowNum = targetRowIndex + 1;

  headers.forEach((header, colIndex) => {
    if (payload.hasOwnProperty(header)) {
      sheet.getRange(rowNum, colIndex + 1).setValue(payload[header]);
    }
  });

  SpreadsheetApp.flush();
  return successResponse(payload);
}

/**
 * データ削除
 */
function deleteData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const targetRowIndex = data.findIndex((row, index) => index > 0 && row[0] === payload.id);

  if (targetRowIndex === -1) {
    return errorResponse('Data not found');
  }

  sheet.deleteRow(targetRowIndex + 1);
  SpreadsheetApp.flush();

  return successResponse({ id: payload.id });
}

/**
 * 成功レスポンス
 */
function successResponse(data) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: true, data: data })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * エラーレスポンス
 */
function errorResponse(error) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, error: error })
  ).setMimeType(ContentService.MimeType.JSON);
}
