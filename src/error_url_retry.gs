/**
 * エラーURL再実行関連の機能
 * 結果シートでERRORになったURLを再実行・更新する
 */

/**
 * 結果シートからERRORになったURLを特定して再実行
 */
function retryErrorUrls() {
  const startTime = new Date().getTime();
  const config = getConfig();
  Logger.log('=== エラーURL再実行開始 ===');

  try {
    // エラーURLを取得
    const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);

    if (errorUrls.length === 0) {
      Logger.log('再実行対象のエラーURLはありません');
      return { processedCount: 0, successCount: 0, hasMoreUrls: false };
    }

    Logger.log(`エラーURL数: ${errorUrls.length}件`);

    const result = processErrorUrlsWithTimeLimit(startTime, errorUrls);

    if (result.hasMoreUrls) {
      Logger.log(`エラーURL再実行: 時間制限により中断。処理 ${result.processedCount}件, 成功 ${result.successCount}件`);
      Logger.log(`残り ${errorUrls.length - result.processedCount}件のエラーURLがあります`);
    } else {
      Logger.log(`エラーURL再実行完了: 処理 ${result.processedCount}件, 成功 ${result.successCount}件`);
    }

    return result;

  } catch (error) {
    Logger.log(`エラーURL再実行処理エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 実行時間制限を考慮したエラーURL処理
 */
function processErrorUrlsWithTimeLimit(startTime, errorUrls) {
  const config = getConfig();
  const EXECUTION_TIME_LIMIT = 300; // 5分（300秒）
  const SAFETY_MARGIN = 30; // 安全マージン（30秒）

  let processedCount = 0;
  let successCount = 0;

  for (let i = 0; i < errorUrls.length; i++) {
    // 実行時間をチェック
    const currentTime = new Date().getTime();
    const elapsedTime = (currentTime - startTime) / 1000; // 秒

    if (elapsedTime > (EXECUTION_TIME_LIMIT - SAFETY_MARGIN)) {
      Logger.log(`実行時間制限に近づきました。処理を中断します。経過時間: ${Math.round(elapsedTime)}秒`);
      return { processedCount, successCount, hasMoreUrls: true };
    }

    const errorData = errorUrls[i];

    try {
      Logger.log(`エラーURL再実行中 ${i + 1}/${errorUrls.length}: ${errorData.url} (経過時間: ${Math.round(elapsedTime)}秒)`);

      // スクレイピング実行
      const hotelData = scrapeHotelFromUrl(errorData.url);

      // 結果シートの該当行を更新（エラー行を正常データで上書き）
      updateResultSheetRow(config.SPREADSHEET_ID, config.SHEET_NAME, errorData.rowIndex, hotelData);

      // マスタシートの最終クロール日時を更新
      const masterRowIndex = findMasterSheetRowByUrl(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, errorData.url);
      if (masterRowIndex > 0) {
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, masterRowIndex, new Date());
      }

      successCount++;
      processedCount++;

      Logger.log(`エラーURL再実行成功: ${errorData.url}`);

      // リクエスト間隔を空ける
      if (i < errorUrls.length - 1) {
        Utilities.sleep(config.REQUEST_DELAY);
      }

    } catch (error) {
      Logger.log(`エラーURL再実行失敗: ${errorData.url} - ${error.toString()}`);

      // エラー情報を更新（再実行した日時で更新）
      const updatedErrorData = {
        title: 'ERROR (RETRY)',
        score: '',
        total: '',
        address: '',
        tel: '',
        totalRooms: '',
        url: errorData.url,
        error: `${error.toString()} (再実行: ${new Date().toLocaleString()})`
      };

      updateResultSheetRow(config.SPREADSHEET_ID, config.SHEET_NAME, errorData.rowIndex, updatedErrorData);
      processedCount++;
    }
  }

  return { processedCount, successCount, hasMoreUrls: false };
}

/**
 * 結果シートからERRORになったURLを取得
 */
function getErrorUrlsFromResultSheet(spreadsheetId, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const resultSheet = spreadsheet.getSheetByName(sheetName);

    if (!resultSheet) {
      Logger.log(`結果シート "${sheetName}" が見つかりません`);
      return [];
    }

    const lastRow = resultSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('結果シートにデータがありません');
      return [];
    }

    // データを取得（ヘッダー行を除く）
    const data = resultSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const errorUrls = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const title = row[0]; // A列: 施設名
      const url = row[6];   // G列: URL
      const error = row[7]; // H列: エラー

      // タイトルが"ERROR"または"ERROR (RETRY)"で始まり、URLが存在する場合
      if (title && title.toString().startsWith('ERROR') && url && url.toString().trim() !== '') {
        errorUrls.push({
          url: url.toString().trim(),
          rowIndex: i + 2, // スプレッドシートの行番号（1ベース + ヘッダー行）
          error: error
        });
      }
    }

    Logger.log(`結果シートからエラーURL ${errorUrls.length}件を抽出しました`);
    return errorUrls;

  } catch (error) {
    Logger.log(`結果シート読み取りエラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 結果シートの指定行を更新
 */
function updateResultSheetRow(spreadsheetId, sheetName, rowIndex, hotelData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const resultSheet = spreadsheet.getSheetByName(sheetName);

    if (!resultSheet) {
      throw new Error(`結果シート "${sheetName}" が見つかりません`);
    }

    // 行データを準備
    const rowData = [
      hotelData.title || '',
      hotelData.score || '',
      hotelData.total || '',
      hotelData.address || '',
      hotelData.tel || '',
      hotelData.totalRooms || '',
      hotelData.url || '',
      hotelData.error || '',
      new Date().toLocaleString()
    ];

    // 指定行を更新
    resultSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

    // 成功時は背景色を緑、エラー時は赤に設定（該当行のみ）
    const updateRange = resultSheet.getRange(rowIndex, 1, 1, rowData.length);
    const bgColor = hotelData.error ? '#ffcccc' : '#ccffcc';
    updateRange.setBackground(bgColor);

    Logger.log(`結果シート行${rowIndex}を更新しました`);

  } catch (error) {
    Logger.log(`結果シート更新エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * マスタシートから指定URLの行番号を検索
 */
function findMasterSheetRowByUrl(spreadsheetId, masterSheetName, targetUrl) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(masterSheetName);

    if (!masterSheet) {
      Logger.log(`マスタシート "${masterSheetName}" が見つかりません`);
      return -1;
    }

    const lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) {
      return -1;
    }

    // URL列（B列）をチェック
    const urlData = masterSheet.getRange(2, 2, lastRow - 1, 1).getValues();

    for (let i = 0; i < urlData.length; i++) {
      const url = urlData[i][0];
      if (url && url.toString().trim() === targetUrl.trim()) {
        return i + 2; // スプレッドシートの行番号（1ベース + ヘッダー行）
      }
    }

    Logger.log(`マスタシートにURL "${targetUrl}" が見つかりませんでした`);
    return -1;

  } catch (error) {
    Logger.log(`マスタシート検索エラー: ${error.toString()}`);
    return -1;
  }
}

/**
 * 日次エラーURL再実行（トリガー用）
 */
function dailyRetryErrorUrls() {
  Logger.log('=== 日次エラーURL再実行開始 ===');

  try {
    // 進捗管理機能付きエラーURL再実行を使用
    retryErrorUrlsWithProgressTracking();

  } catch (error) {
    Logger.log(`日次エラーURL再実行エラー: ${error.toString()}`);
  }
}

/**
 * エラーURL再実行のステータス確認（統合版）
 */
function showErrorRetryStatusLegacy() {
  const config = getConfig();
  Logger.log('=== エラーURL再実行ステータス（従来版） ===');

  try {
    // エラーURLの件数を確認
    const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);
    Logger.log(`現在のエラーURL数: ${errorUrls.length}件`);

    if (errorUrls.length > 0) {
      Logger.log('エラーURL一覧:');
      errorUrls.slice(0, 5).forEach((errorData, index) => {
        Logger.log(`  ${index + 1}. ${errorData.url} (行: ${errorData.rowIndex})`);
      });

      if (errorUrls.length > 5) {
        Logger.log(`  ... その他 ${errorUrls.length - 5}件`);
      }
    }

    // 日次トリガーの確認
    const triggers = ScriptApp.getProjectTriggers();
    const retryTrigger = triggers.find(t => t.getHandlerFunction() === 'dailyRetryErrorUrls');

    if (retryTrigger) {
      Logger.log('日次エラーURL再実行トリガー: 設定済み');
    } else {
      Logger.log('日次エラーURL再実行トリガー: 未設定');
    }

    // 新しいステータス確認も表示
    Logger.log('\n--- 進捗管理機能付きステータス ---');
    showErrorRetryStatus();

  } catch (error) {
    Logger.log(`ステータス確認エラー: ${error.toString()}`);
  }
}

/**
 * エラーURL再実行トリガー設定
 */
function setupErrorRetryTrigger() {
  try {
    // 既存のエラーURL再実行トリガーを削除
    clearErrorRetryTriggers();

    // 設定を取得
    const config = getConfig();

    // エラーURL再実行用トリガー（朝10時）
    ScriptApp.newTrigger('dailyRetryErrorUrls')
      .timeBased()
      .everyDays(1)
      .atHour(config.DAILY_EXECUTION_HOUR + 1)
      .create();

    Logger.log(`エラーURL再実行トリガーを設定しました: 毎日 ${config.DAILY_EXECUTION_HOUR + 1}:00`);

  } catch (error) {
    Logger.log(`エラーURL再実行トリガー設定エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * エラーURL再実行トリガーをクリア
 */
function clearErrorRetryTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'dailyRetryErrorUrls') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });

    if (deletedCount > 0) {
      Logger.log(`既存のエラーURL再実行トリガーを ${deletedCount} 個削除しました`);
    }

  } catch (error) {
    Logger.log(`エラーURL再実行トリガークリアエラー: ${error.toString()}`);
  }
}

/**
 * 進捗管理機能付きエラーURL再実行
 */
function retryErrorUrlsWithProgressTracking() {
  const startTime = new Date().getTime();
  const config = getConfig();

  try {
    // 実行進捗を取得
    const progress = getErrorRetryProgress();
    Logger.log(`=== 進捗管理付きエラーURL再実行開始 ===`);
    Logger.log(`開始時点での処理済み件数: ${progress.totalProcessed}`);

    const result = processErrorUrlsWithProgressTracking(startTime, progress);

    if (result.hasMoreUrls) {
      Logger.log(`実行時間制限により処理中断。今回処理件数: ${result.processedThisRun}件`);
      Logger.log(`累計処理件数: ${progress.totalProcessed + result.processedThisRun}件`);

      // 進捗を保存
      saveErrorRetryProgress(progress.totalProcessed + result.processedThisRun);

      // 次回実行をスケジュール
      try {
        scheduleNextErrorRetryExecution();
        Logger.log('次回実行スケジュール設定完了');
      } catch (scheduleError) {
        Logger.log(`スケジュール設定エラー: ${scheduleError.toString()}`);
        // スケジュール設定に失敗した場合も進捗は保存されているので、手動再実行可能
      }
    } else {
      Logger.log(`全エラーURL再実行完了。今回処理件数: ${result.processedThisRun}件`);
      Logger.log(`累計処理件数: ${progress.totalProcessed + result.processedThisRun}件`);

      // 進捗をリセット
      clearErrorRetryProgress();

      // トリガーをクリア
      clearErrorRetryProgressTriggers();
    }

  } catch (error) {
    Logger.log('進捗管理付きエラーURL再実行エラー: ' + error.toString());

    // エラーが発生した場合でも、未処理URLがあるかチェックして継続処理を試行
    try {
      const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);
      if (errorUrls.length > 0) {
        Logger.log(`エラー発生時に未処理URL ${errorUrls.length}件を確認。継続処理を試行します。`);
        // 少し時間を置いて再試行をスケジュール
        setTimeout(() => {
          try {
            scheduleNextErrorRetryExecution();
            Logger.log('エラー後の継続処理スケジュール設定完了');
          } catch (retryScheduleError) {
            Logger.log(`エラー後のスケジュール設定失敗: ${retryScheduleError.toString()}`);
          }
        }, 5000); // 5秒後に再試行
      }
    } catch (checkError) {
      Logger.log(`エラー後の状況確認失敗: ${checkError.toString()}`);
    }
  }
}

/**
 * 進捗管理付きエラーURL処理
 */
function processErrorUrlsWithProgressTracking(startTime, progress) {
  const config = getConfig();
  const EXECUTION_TIME_LIMIT = 300; // 5分（300秒）
  const SAFETY_MARGIN = 30; // 安全マージン（30秒）

  // エラーURLを取得
  const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);

  if (errorUrls.length === 0) {
    Logger.log('再実行対象のエラーURLはありません');
    return { processedThisRun: 0, successCount: 0, hasMoreUrls: false };
  }

  Logger.log(`エラーURL数: ${errorUrls.length}件`);

  let processedThisRun = 0;
  let successCount = 0;

  for (let i = 0; i < errorUrls.length; i++) {
    // 実行時間をチェック
    const currentTime = new Date().getTime();
    const elapsedTime = (currentTime - startTime) / 1000;

    if (elapsedTime > (EXECUTION_TIME_LIMIT - SAFETY_MARGIN)) {
      Logger.log(`実行時間制限に近づきました。経過時間: ${Math.round(elapsedTime)}秒`);
      return { processedThisRun, successCount, hasMoreUrls: true };
    }

    const errorData = errorUrls[i];

    try {
      Logger.log(`エラーURL再実行中 ${i + 1}/${errorUrls.length}: ${errorData.url} (経過: ${Math.round(elapsedTime)}秒, 累計: ${progress.totalProcessed + processedThisRun})`);

      // スクレイピング実行
      const hotelData = scrapeHotelFromUrl(errorData.url);

      // 結果シートの該当行を更新
      updateResultSheetRow(config.SPREADSHEET_ID, config.SHEET_NAME, errorData.rowIndex, hotelData);

      // マスタシートの最終クロール日時を更新
      const masterRowIndex = findMasterSheetRowByUrl(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, errorData.url);
      if (masterRowIndex > 0) {
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, masterRowIndex, new Date());
      }

      successCount++;
      processedThisRun++;

      // 10件ごとに進捗を保存
      if (processedThisRun % 10 === 0) {
        saveErrorRetryProgress(progress.totalProcessed + processedThisRun);
        Logger.log(`エラーURL再実行進捗保存: ${progress.totalProcessed + processedThisRun}件処理済み`);
      }

      // リクエスト間隔を空ける
      if (i < errorUrls.length - 1) {
        Utilities.sleep(config.REQUEST_DELAY);
      }

    } catch (error) {
      Logger.log(`エラーURL再実行失敗: ${errorData.url} - ${error.toString()}`);

      // エラー情報を更新
      const updatedErrorData = {
        title: 'ERROR (RETRY)',
        score: '',
        total: '',
        address: '',
        tel: '',
        totalRooms: '',
        url: errorData.url,
        error: `${error.toString()} (再実行: ${new Date().toLocaleString()})`
      };

      updateResultSheetRow(config.SPREADSHEET_ID, config.SHEET_NAME, errorData.rowIndex, updatedErrorData);
      processedThisRun++;
    }
  }

  return { processedThisRun, successCount, hasMoreUrls: false };
}

/**
 * エラーURL再実行進捗を取得
 */
function getErrorRetryProgress() {
  const properties = PropertiesService.getScriptProperties();
  const totalProcessed = parseInt(properties.getProperty('ERROR_RETRY_TOTAL_PROCESSED') || '0');
  const lastExecutionTime = properties.getProperty('ERROR_RETRY_LAST_EXECUTION_TIME');

  return {
    totalProcessed: totalProcessed,
    lastExecutionTime: lastExecutionTime
  };
}

/**
 * エラーURL再実行進捗を保存
 */
function saveErrorRetryProgress(totalProcessed) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ERROR_RETRY_TOTAL_PROCESSED': totalProcessed.toString(),
    'ERROR_RETRY_LAST_EXECUTION_TIME': new Date().toISOString()
  });
}

/**
 * エラーURL再実行進捗をクリア
 */
function clearErrorRetryProgress() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('ERROR_RETRY_TOTAL_PROCESSED');
  properties.deleteProperty('ERROR_RETRY_LAST_EXECUTION_TIME');
  Logger.log('エラーURL再実行進捗をクリアしました。');
}

/**
 * 次回エラーURL再実行をスケジュール
 */
function scheduleNextErrorRetryExecution() {
  try {
    Logger.log('次回エラーURL再実行スケジュール開始');

    // 既存のトリガーを削除
    clearErrorRetryProgressTriggers();

    // 現在時刻から2分後に実行するトリガーを作成（余裕を持たせる）
    const triggerTime = new Date();
    triggerTime.setMinutes(triggerTime.getMinutes() + 2);

    const trigger = ScriptApp.newTrigger('retryErrorUrlsWithProgressTracking')
      .timeBased()
      .at(triggerTime)
      .create();

    Logger.log(`次回エラーURL再実行時刻をスケジュール: ${triggerTime.toLocaleString()}`);
    Logger.log(`トリガーID: ${trigger.getUniqueId()}`);

    // トリガーが正常に作成されたかを確認
    const allTriggers = ScriptApp.getProjectTriggers();
    const retryTriggers = allTriggers.filter(t => t.getHandlerFunction() === 'retryErrorUrlsWithProgressTracking');
    Logger.log(`現在のエラーURL再実行トリガー数: ${retryTriggers.length}件`);

  } catch (error) {
    Logger.log('エラーURL再実行トリガー設定エラー: ' + error.toString());

    // 再試行のためのより詳細なエラー情報を記録
    Logger.log(`エラー詳細: ${JSON.stringify({
      name: error.name,
      message: error.message,
      stack: error.stack
    })}`);

    throw error; // エラーを再スローして上位で処理
  }
}

/**
 * エラーURL再実行進捗管理用トリガーをクリア
 */
function clearErrorRetryProgressTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'retryErrorUrlsWithProgressTracking') {
        try {
          ScriptApp.deleteTrigger(trigger);
          deletedCount++;
          Logger.log(`トリガー削除: ID=${trigger.getUniqueId()}`);
        } catch (deleteError) {
          Logger.log(`トリガー削除エラー: ID=${trigger.getUniqueId()}, エラー=${deleteError.toString()}`);
        }
      }
    });

    if (deletedCount > 0) {
      Logger.log(`${deletedCount}個のエラーURL再実行進捗管理トリガーを削除しました。`);
    } else {
      Logger.log('削除対象のエラーURL再実行進捗管理トリガーはありませんでした。');
    }

  } catch (error) {
    Logger.log(`トリガークリアエラー: ${error.toString()}`);
  }
}

/**
 * エラーURL再実行状況表示
 */
function showErrorRetryStatus() {
  const progress = getErrorRetryProgress();
  const config = getConfig();

  Logger.log('=== エラーURL再実行状況 ===');
  Logger.log(`累計処理件数: ${progress.totalProcessed}件`);
  Logger.log(`最終実行時刻: ${progress.lastExecutionTime || '未実行'}`);

  try {
    const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);
    Logger.log(`現在のエラーURL数: ${errorUrls.length}件`);

    const triggers = ScriptApp.getProjectTriggers().filter(t =>
      t.getHandlerFunction() === 'retryErrorUrlsWithProgressTracking'
    );
    Logger.log(`スケジュール済みエラーURL再実行トリガー数: ${triggers.length}件`);

    if (errorUrls.length > 0) {
      Logger.log('エラーURL一覧（最初の5件）:');
      errorUrls.slice(0, 5).forEach((errorData, index) => {
        Logger.log(`  ${index + 1}. ${errorData.url} (行: ${errorData.rowIndex})`);
      });

      if (errorUrls.length > 5) {
        Logger.log(`  ... その他 ${errorUrls.length - 5}件`);
      }
    }

  } catch (error) {
    Logger.log('エラーURL再実行状況取得エラー: ' + error.toString());
  }
}
