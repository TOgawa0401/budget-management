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
  const headers = ['id', 'category', 'name', 'budget26', 'actual', 'forecast', 'responsible', 'executionDate', 'details', 'proposalUrl', 'kintone_record_id', 'kintone_import_timestamp', 'kintone_sync_status'];
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 初期データ
  const initialData = [
    [1, '予算化項目', 'LineWorks 契約費用', 2160000, 2160000, 0, '', '', '216万で400アカウントで契約。2026年4月契約', '', '', '', ''],
    [2, '予算化項目', 'LineWorks 追加100アカウント', 0, 540000, 0, '', '', '100アカウント追加54万', '', '', '', ''],
    [3, '予算化項目', 'アルバイト募集有料施策（茨城大）', 125000, 8800, 0, '', '', '茨城大に募集広告掲載', '', '', '', ''],
    [4, '予算化項目', 'アルバイト募集有料施策（静岡大）', 125000, 5500, 0, '', '', '静岡大に募集広告掲載', '', '', '', ''],
    [5, '予算化項目', 'アルバイト募集有料施策（広島大）', 125000, 3300, 0, '', '', '広島大に募集広告掲載', '', '', '', ''],
    [6, '予算化項目', 'アルバイト募集有料施策（長崎大）', 125000, 0, 0, '', '', '長崎大に募集広告掲載（二重広告掲載）', '', '', '', ''],
    [7, '予算化項目', 'AI有料アカウント契約料', 400000, 350000, 0, '', '', 'Claude 8名のプロプランを契約', '', '', '', ''],
    [8, '予算化項目', 'ispring契約費用', 390000, 0, 390000, '', '', '2025年11月に契約更新。次回更新は2026年11月。更新は要検討', '', '', '', ''],
    [9, '予算化項目', '野球エキスパート慰労会費用', 200000, 0, 200000, '', '', '', '', '', '', ''],
    [10, '予算化項目', 'バスケエキスパート慰労会費用', 200000, 0, 200000, '', '', '', '', '', '', ''],
    [11, '予算化項目', 'FBEXP経費精算システム', 180000, 180000, 0, '', '', '', '', '', '', ''],
    [12, '予算化項目', 'FBEXP納会費用', 180000, 0, 180000, '', '', '', '', '', '', ''],
    [13, '予算化項目', '書籍雑誌購入費用', 72000, 50000, 0, '', '', '', '', '', '', ''],
    [14, '予算化項目', '入力室備品', 50000, 0, 0, '', '', '未執行', '', '', '', ''],
    [15, '予算化項目', 'サッカー映像保管用HDD購入費', 50000, 0, 50000, '', '', '', '', '', '', ''],
    [16, '予算化項目', '卓球エキスパート慰労会費用', 50000, 0, 50000, '', '', '', '', '', '', ''],
    [17, '予算化項目', 'DS Football Twitter(X)アカウント有料プランサブスク料', 9180, 0, 9180, '', '', '', '', '', '', ''],
    [18, '追加項目', 'PowerAutomate有料契約', 49464, 49464, 0, '', '', '', '', '', '', '']
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
      case 'importFromKintone':
        return importFromKintoneAPI();
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
    payload.details || '',
    payload.proposalUrl || '',
    payload.kintone_record_id || '',
    payload.kintone_import_timestamp || '',
    payload.kintone_sync_status || ''
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

// ===== kintone インポート機能 =====

/**
 * kintone 認証情報取得
 * 事前に Google Apps Script の「プロジェクト設定」→「スクリプトプロパティ」に以下を設定してください：
 * KINTONE_SUBDOMAIN: "dtsmx"
 * KINTONE_API_TOKEN: "VRVWhHR87cy4NYcB1RMZ0kEeyWtxTAB38BhPHBub"
 * KINTONE_APP_ID: "375"
 */
function getKintoneCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    subdomain: scriptProperties.getProperty('KINTONE_SUBDOMAIN'),
    token: scriptProperties.getProperty('KINTONE_API_TOKEN'),
    appId: scriptProperties.getProperty('KINTONE_APP_ID'),
  };
}

/**
 * インポートログシート作成（初回のみ）
 */
function createImportLogSheetIfNotExists() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheetName = 'インポートログ';

  let logSheet = spreadsheet.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(logSheetName);
    const headers = ['実行日時', 'ステータス', 'メッセージ', 'インポート件数', 'エラー詳細'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダーのフォーマット
    const headerRange = logSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0a1020').setFontColor('#f1f5f9').setFontWeight('bold');
  }
  return logSheet;
}

/**
 * インポートイベントをログに記録
 */
function logImportEvent(status, message, recordCount, errorDetails) {
  const logSheet = createImportLogSheetIfNotExists();
  const timestamp = new Date().toLocaleString('ja-JP');
  const row = [timestamp, status, message, recordCount, JSON.stringify(errorDetails)];
  logSheet.appendRow(row);
}

/**
 * レート制限対応フェッチ関数
 */
function fetchWithRetry(url, options, maxRetries = 3) {
  let lastError;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        ...options,
        muteHttpExceptions: true,
      });
      const code = response.getResponseCode();

      if (code === 200) {
        return response;
      }

      if (code === 429) {
        // レート制限：指数バックオフで待機
        const waitTime = Math.pow(2, i) * 1000;
        Logger.log(`Rate limit hit, waiting ${waitTime}ms before retry ${i + 1}/${maxRetries}`);
        Utilities.sleep(waitTime);
        continue;
      }

      lastError = `HTTP ${code}: ${response.getContentText()}`;
    } catch (e) {
      lastError = e.message;
    }
  }
  throw new Error(`API failed after ${maxRetries} retries: ${lastError}`);
}

/**
 * kintone からレコード取得
 */
function fetchKintoneRecords() {
  const creds = getKintoneCredentials();
  if (!creds.subdomain || !creds.token || !creds.appId) {
    throw new Error('kintone 認証情報が設定されていません。プロジェクト設定を確認してください。');
  }

  // KQL: 対象部署 + 販管費種別（部署経費） + 申請日の範囲
  // 注：承認状態のフィルタリングは不要（要件から）
  const kql = `対象部署 like "スポーツデータセンター / データ推進部" AND 販管費種別 = "部署経費" AND 申請日時 >= "2026-04-01" AND 申請日時 < "2027-04-01"`;

  const url = `https://${creds.subdomain}.cybozu.com/k/v1/records?app=${creds.appId}&query=${encodeURIComponent(kql)}&limit=100`;
  const options = {
    method: 'GET',
    headers: {
      'X-Cybozu-API-Token': creds.token,
    },
  };

  const response = fetchWithRetry(url, options);
  const json = JSON.parse(response.getContentText());

  if (!json.records) {
    throw new Error('Invalid kintone response: missing records field');
  }

  return json.records;
}

/**
 * kintone レコード検証
 */
function validateKintoneRecord(record) {
  const errors = [];

  // 必須フィールドチェック
  if (!record['件名'] || !record['件名'].value) {
    errors.push('案件名（件名）が空');
  }

  if (!record['申請日時'] || !record['申請日時'].value) {
    errors.push('申請日時が空');
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

/**
 * kintone レコード → 予算管理レコード への変換
 */
function transformKintoneRecord(kintoneRecord) {
  return {
    name: (kintoneRecord['件名'] && kintoneRecord['件名'].value) || '',
    category: '予算化項目', // デフォルト
    actual: (kintoneRecord['申請金額'] && parseFloat(kintoneRecord['申請金額'].value)) || 0,
    forecast: 0,
    executionDate: (kintoneRecord['申請日時'] && kintoneRecord['申請日時'].value) || '',
    responsible: (kintoneRecord['申請者'] && kintoneRecord['申請者'].value && kintoneRecord['申請者'].value.name) || '',
    details: `[kintone より自動インポート] レコードID: ${kintoneRecord.$id.value}`,
    kintone_record_id: kintoneRecord.$id.value,
    kintone_import_timestamp: new Date().toISOString(),
    kintone_sync_status: 'synced',
  };
}

/**
 * kintone レコードが既にインポート済みかチェック
 */
function isDuplicateImport(kintoneRecordId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // kintone_record_id カラムのインデックスを取得
  const kintoneIdIndex = headers.indexOf('kintone_record_id');
  if (kintoneIdIndex === -1) return false;

  // 既存レコードをチェック
  for (let i = 1; i < data.length; i++) {
    if (data[i][kintoneIdIndex] == kintoneRecordId) {
      return true;
    }
  }
  return false;
}

/**
 * 既存のデータから新規レコードのみをフィルタ
 */
function filterDuplicates(kintoneRecords) {
  return kintoneRecords.filter(record => {
    return !isDuplicateImport(record.$id.value);
  });
}

/**
 * メイン: kintone からデータをインポート
 */
function importKintoneExpenses() {
  const lockService = LockService.getScriptLock();

  try {
    // 他の操作と競合しないようロックを取得（30秒待機）
    lockService.waitLock(30000);

    const result = {
      imported: [],
      updated: [],
      skipped: [],
      errors: [],
    };

    Logger.log('[kintone import] Starting import process...');

    // 1. kintone からレコード取得
    const kintoneRecords = fetchKintoneRecords();
    Logger.log(`[kintone import] Fetched ${kintoneRecords.length} records from kintone`);

    // 2. 重複排除
    const newRecords = filterDuplicates(kintoneRecords);
    Logger.log(`[kintone import] After duplicate filter: ${newRecords.length} new records`);

    // 3. バリデーション＆変換
    const validRecords = [];
    for (const record of newRecords) {
      const validation = validateKintoneRecord(record);
      if (validation.isValid) {
        const transformed = transformKintoneRecord(record);
        validRecords.push(transformed);
      } else {
        result.skipped.push({
          recordId: record.$id.value,
          reason: validation.errors.join('; '),
        });
      }
    }
    Logger.log(`[kintone import] Valid records: ${validRecords.length}`);

    // 4. データベースに追加
    for (const record of validRecords) {
      try {
        addData(record);
        result.imported.push(record.name);
      } catch (e) {
        result.errors.push({
          record: record.name,
          error: e.message,
        });
        Logger.log(`[kintone import] Error adding record: ${e.message}`);
      }
    }

    // 5. ログに記録
    const status = result.errors.length === 0 ? 'success' : 'partial';
    const message = `${result.imported.length}件をインポート${result.errors.length > 0 ? `, ${result.errors.length}件エラー` : ''}`;
    logImportEvent(status, message, result.imported.length, result.errors);

    Logger.log(`[kintone import] Import completed: ${message}`);
    return result;

  } catch (error) {
    const errorMsg = `インポート失敗: ${error.message}`;
    Logger.log(`[kintone import] Error: ${errorMsg}`);
    logImportEvent('failed', errorMsg, 0, [error.message]);
    throw error;

  } finally {
    lockService.releaseLock();
  }
}

/**
 * 時間ベーストリガー用：定期実行関数（1時間毎に実行）
 * GAS のトリガー設定で「時間駆動型」→「時間ベースのトリガー」→「1時間ごと」を選択
 */
function scheduledImportFromKintone() {
  try {
    const result = importKintoneExpenses();
    Logger.log(`[Scheduled] Import successful: ${result.imported.length} records imported`);
  } catch (error) {
    Logger.log(`[Scheduled] Import failed: ${error.message}`);
    sendErrorNotificationEmail(error, 'SCHEDULED_IMPORT_ERROR');
  }
}

/**
 * API エンドポイント：UI から呼び出される手動インポート
 */
function importFromKintoneAPI() {
  try {
    const result = importKintoneExpenses();
    return successResponse(result);
  } catch (error) {
    Logger.log(`[API] Import error: ${error.message}`);
    return errorResponse(error.message);
  }
}

/**
 * エラー通知メール送信
 */
function sendErrorNotificationEmail(error, errorType) {
  const adminEmail = 'noreply@example.com'; // TODO: 管理者メールアドレスに変更
  const subject = `【警告】予算管理システム - kintone インポートエラー (${errorType})`;
  const body = `予算管理システムで kintone インポートエラーが発生しました。\n\n` +
               `エラータイプ: ${errorType}\n` +
               `メッセージ: ${error.message}\n` +
               `実行日時: ${new Date().toLocaleString('ja-JP')}\n\n` +
               `詳細はログシート「インポートログ」を確認してください。`;

  try {
    GmailApp.sendEmail(adminEmail, subject, body);
  } catch (e) {
    Logger.log(`[Error] Failed to send notification email: ${e.message}`);
  }
}
