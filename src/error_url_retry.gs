/**
 * エラーURL再実行関連の機能
 * 結果シートでERRORになったURLを再実行・更新する
 */

/**
 * 結果シートからERRORになったURLを特定して再実行
 */
function retryErrorUrls() {
  const config = getConfig();
  Logger.log('=== エラーURL再実行開始 ===');

  try {
    // エラーURLを取得
    const errorUrls = getErrorUrlsFromResultSheet(config.SPREADSHEET_ID, config.SHEET_NAME);

    if (errorUrls.length === 0) {
      Logger.log('再実行対象のエラーURLはありません');
      return { processedCount: 0, successCount: 0 };
    }

    Logger.log(`エラーURL数: ${errorUrls.length}件`);

    let processedCount = 0;
    let successCount = 0;

    for (let i = 0; i < errorUrls.length; i++) {
      const errorData = errorUrls[i];

      try {
        Logger.log(`エラーURL再実行中 ${i + 1}/${errorUrls.length}: ${errorData.url}`);

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

    Logger.log(`エラーURL再実行完了: 処理 ${processedCount}件, 成功 ${successCount}件`);
    return { processedCount, successCount };

  } catch (error) {
    Logger.log(`エラーURL再実行処理エラー: ${error.toString()}`);
    throw error;
  }
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
    const result = retryErrorUrls();
    Logger.log(`日次エラーURL再実行完了: 処理 ${result.processedCount}件, 成功 ${result.successCount}件`);

  } catch (error) {
    Logger.log(`日次エラーURL再実行エラー: ${error.toString()}`);
  }
}

/**
 * エラーURL再実行のステータス確認
 */
function showErrorRetryStatus() {
  const config = getConfig();
  Logger.log('=== エラーURL再実行ステータス ===');

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
