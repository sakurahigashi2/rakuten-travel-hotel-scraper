/**
 * マスタシート連携版 - 楽天トラベルホテル情報スクレイピング
 * マスタシートからURL一覧を取得し、スクレイピング済み日時を記録
 */

// ========================================
// HTML取得・エンコーディング処理ユーティリティ
// ========================================

/**
 * HTMLコンテンツからエンコーディングを推測
 * @param {HTTPResponse} response - UrlFetchApp.fetchのレスポンス
 * @return {string} 検出されたエンコーディング
 */
function detectEncodingFromHtml(response) {
  try {
    // まずUTF-8で試してみる
    let html = response.getContentText('UTF-8');

    // メタタグでcharsetをチェック
    const charsetMatch = html.match(/<meta[^>]*charset\s*=\s*["']?([^"'>\s]+)/i);
    if (charsetMatch) {
      const charset = charsetMatch[1].toLowerCase();
      Logger.log(`HTMLのメタタグから検出されたcharset: ${charset}`);

      if (charset.includes('shift') || charset.includes('sjis')) {
        return 'Shift_JIS';
      } else if (charset.includes('euc')) {
        return 'EUC-JP';
      } else if (charset.includes('utf-8') || charset.includes('utf8')) {
        return 'UTF-8';
      }
    }

    // Content-Typeヘッダーをチェック
    const contentType = response.getHeaders()['Content-Type'] || '';
    if (contentType.includes('shift_jis') || contentType.includes('shift-jis')) {
      return 'Shift_JIS';
    } else if (contentType.includes('euc-jp')) {
      return 'EUC-JP';
    }

    // 文字化けパターンをチェック（Shift_JISをUTF-8で読んだ場合）
    if (html.includes('�') || html.includes('�u�y') || html.includes('�x��')) {
      Logger.log('文字化けを検出しました。Shift_JISの可能性があります');
      return 'Shift_JIS';
    }

    return 'UTF-8'; // デフォルト

  } catch (error) {
    Logger.log(`エンコーディング検出エラー: ${error.toString()}`);
    return 'UTF-8';
  }
}

/**
 * 適切なエンコーディングでHTMLを取得
 * @param {string} url - 取得するURL
 * @param {Object} options - UrlFetchAppのオプション（オプション）
 * @return {string} HTMLコンテンツ
 */
function getHtmlWithCorrectEncoding(url, options = {}) {
  try {
    const defaultOptions = {
      'method': 'GET',
      'followRedirects': true,
      'muteHttpExceptions': true
    };

    // オプションをマージ
    const fetchOptions = { ...defaultOptions, ...options };

    const response = UrlFetchApp.fetch(url, fetchOptions);

    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTPエラー: ${response.getResponseCode()}`);
    }

    // エンコーディングを検出
    const encoding = detectEncodingFromHtml(response);
    Logger.log(`検出されたエンコーディング: ${encoding}`);

    // 適切なエンコーディングでHTMLを取得
    const html = response.getContentText(encoding);

    return html;

  } catch (error) {
    Logger.log(`HTML取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 複数URLからHTMLを順次取得（レート制限対応）
 * @param {Array<string>} urls - 取得するURL配列
 * @param {number} delay - URL間の待機時間（ミリ秒）
 * @return {Array<Object>} {url, html, success, error}の配列
 */
function getMultipleHtmlWithEncoding(urls, delay = 2000) {
  const results = [];

  urls.forEach((url, index) => {
    Logger.log(`${index + 1}/${urls.length}: ${url} を取得中...`);

    try {
      const html = getHtmlWithCorrectEncoding(url);
      results.push({
        url: url,
        html: html,
        success: true,
        error: null
      });
      Logger.log(`✓ 成功: ${html.length} 文字`);
    } catch (error) {
      results.push({
        url: url,
        html: null,
        success: false,
        error: error.toString()
      });
      Logger.log(`✗ 失敗: ${error.toString()}`);
    }

    // 最後のURL以外は待機
    if (index < urls.length - 1) {
      Logger.log(`${delay}ms 待機中...`);
      Utilities.sleep(delay);
    }
  });

  return results;
}

// ========================================
// マスタシート連携版メイン処理
// ========================================

// 実行時間制限の設定（秒）
const EXECUTION_TIME_LIMIT = 300; // 5分（300秒）
const SAFETY_MARGIN = 30; // 安全マージン（30秒）

/**
 * 新着ホテル収集 + スクレイピング統合実行
 */
function mainWithNewHotelCollection() {
  Logger.log('=== 新着ホテル収集 + スクレイピング統合実行開始 ===');

  try {
    // 1. 新着ホテル情報を収集
    Logger.log('Step 1: 新着ホテル情報を収集中...');
    collectNewHotels();

    // 少し待機
    Utilities.sleep(2000);

    // 2. 通常のスクレイピング処理を実行
    Logger.log('Step 2: スクレイピング処理を開始...');
    mainWithProgressTracking();

  } catch (error) {
    Logger.log(`統合実行エラー: ${error.toString()}`);
  }

  Logger.log('=== 新着ホテル収集 + スクレイピング統合実行終了 ===');
}

/**
 * 実行時間制限を考慮した未処理URL処理
 */
function processUnprocessedUrlsWithTimeLimit(startTime) {
  const config = getConfig();

  // 未処理URLを取得
  const unprocessedUrls = getUnprocessedUrls(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME);

  if (unprocessedUrls.length === 0) {
    Logger.log('未処理のURLはありません。');
    return { processedCount: 0, hasMoreUrls: false };
  }

  Logger.log(`未処理URL数: ${unprocessedUrls.length}件`);

  let processedCount = 0;

  for (let i = 0; i < unprocessedUrls.length; i++) {
    // 実行時間をチェック
    const currentTime = new Date().getTime();
    const elapsedTime = (currentTime - startTime) / 1000; // 秒

    if (elapsedTime > (EXECUTION_TIME_LIMIT - SAFETY_MARGIN)) {
      Logger.log(`実行時間制限に近づきました。処理を中断します。経過時間: ${Math.round(elapsedTime)}秒`);
      return { processedCount, hasMoreUrls: true };
    }

    const urlData = unprocessedUrls[i];

    try {
      Logger.log(`処理中 ${i + 1}/${unprocessedUrls.length}: ID=${urlData.id} (経過時間: ${Math.round(elapsedTime)}秒)`);

      // スクレイピング実行
      const hotelData = scrapeHotelFromUrl(urlData.url);

      // 結果シートに書き込み
      writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, hotelData);

      // マスタシートの最終クロール日時を更新
      updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date());

      processedCount++;

      // リクエスト間隔を空ける
      if (i < unprocessedUrls.length - 1) {
        Utilities.sleep(config.REQUEST_DELAY);
      }

    } catch (error) {
      Logger.log(`URL処理エラー ID=${urlData.id}: ${error.toString()}`);

      // エラー情報を結果シートに記録
      const errorData = {
        title: 'ERROR',
        score: '',
        total: '',
        address: '',
        tel: '',
        totalRooms: '',
        url: urlData.url,
        error: error.toString()
      };

      writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, errorData);
      updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date(), error.toString());

      processedCount++;
    }
  }

  return { processedCount, hasMoreUrls: false };
}

/**
 * 未処理URLのみを取得
 */
function getUnprocessedUrls(spreadsheetId, masterSheetName) {
  const allUrls = getUrlListFromMasterSheet(spreadsheetId, masterSheetName);

  // 未処理（last_crawled_atが空）のURLのみを抽出
  return allUrls.filter(urlData => {
    return !urlData.lastCrawledAt || urlData.lastCrawledAt.toString().trim() === '';
  });
}

/**
 * 全トリガーをクリア（手動実行用）
 */
function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'mainWithMasterSheetRecursive') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  Logger.log(`${deletedCount}個のトリガーを削除しました。`);
}

/**
 * マスタシートからURL一覧を取得
 */
function getUrlListFromMasterSheet(spreadsheetId, masterSheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(masterSheetName);

    if (!masterSheet) {
      throw new Error(`マスタシート "${masterSheetName}" が見つかりません。`);
    }

    // データ範囲を取得（ヘッダー行を除く）
    const lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('マスタシートにデータがありません。');
      return [];
    }

    // ヘッダー行をチェック
    const headers = masterSheet.getRange(1, 1, 1, 3).getValues()[0];
    const expectedHeaders = ['id', 'url', 'last_crawled_at'];

    for (let i = 0; i < expectedHeaders.length; i++) {
      if (headers[i].toString().toLowerCase() !== expectedHeaders[i]) {
        Logger.log(`警告: ヘッダーが期待値と異なります。列${i + 1}: 期待値="${expectedHeaders[i]}", 実際="${headers[i]}"`);
      }
    }

    // データを取得
    const data = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const urlList = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const id = row[0];
      const url = row[1];
      const lastCrawledAt = row[2];

      // URLが空でない場合のみ処理対象とする
      if (url && url.toString().trim() !== '') {
        urlList.push({
          id: id,
          url: url.toString().trim(),
          lastCrawledAt: lastCrawledAt,
          rowIndex: i + 2 // スプレッドシートの行番号（1ベース + ヘッダー行）
        });
      }
    }

    Logger.log(`マスタシートから${urlList.length}件のURLを取得しました。`);
    return urlList;

  } catch (error) {
    Logger.log('マスタシート読み取りエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 指定URLからホテル情報をスクレイピング
 */
function scrapeHotelFromUrl(url) {
  try {
    // 共通関数を使用してHTML取得（エンコーディング自動検出）
    const html = getHtmlWithCorrectEncoding(url, {
      timeout: getConfig().TIMEOUT * 1000
    });

    return extractHotelInfoAdvanced(html, url);

  } catch (error) {
    Logger.log('スクレイピングエラー: ' + error.toString());
    throw error;
  }
}

/**
 * マスタシートの最終クロール日時を更新
 */
function updateLastCrawledAt(spreadsheetId, masterSheetName, rowIndex, crawledAt, errorMessage = '') {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(masterSheetName);

    if (!masterSheet) {
      throw new Error(`マスタシート "${masterSheetName}" が見つかりません。`);
    }

    // 最終クロール日時を更新（C列）
    let cellValue;
    if (errorMessage) {
      cellValue = `${crawledAt.toLocaleString()} (エラー: ${errorMessage.substring(0, 50)})`;
    } else {
      cellValue = crawledAt.toLocaleString();
    }

    masterSheet.getRange(rowIndex, 3).setValue(cellValue);

    // エラーの場合は背景色を変更（該当行のみ）
    const rowRange = masterSheet.getRange(rowIndex, 1, 1, 3);
    if (errorMessage) {
      rowRange.setBackground('#ffcccc');
    } else {
      rowRange.setBackground('#ccffcc');
    }

    Logger.log(`マスタシート更新: 行${rowIndex}, 日時=${cellValue}`);

  } catch (error) {
    Logger.log('マスタシート更新エラー: ' + error.toString());
  }
}

/**
 * マスタシートを初期化（テスト用）
 */
function initializeMasterSheet() {
  const config = getConfig();

  if (config.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    Logger.log('エラー: スプレッドシートIDが設定されていません。');
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    let masterSheet = spreadsheet.getSheetByName(config.MASTER_SHEET_NAME);

    // マスタシートが存在しない場合は作成
    if (!masterSheet) {
      masterSheet = spreadsheet.insertSheet(config.MASTER_SHEET_NAME);
    } else {
      // 既存のシートをクリア
      masterSheet.clear();
    }

    // ヘッダー行を作成
    const headers = ['id', 'url', 'last_crawled_at'];
    masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダーのスタイルを設定
    const headerRange = masterSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setBorder(true, true, true, true, true, true);

    // サンプルデータを追加
    const sampleData = [
      ['001', 'https://travel.rakuten.co.jp/HOTEL/104529/104529_std.html', ''],
      ['002', 'https://travel.rakuten.co.jp/HOTEL/104529/104529_std.html', ''],
      // 必要に応じて他のURLを追加
    ];

    if (sampleData.length > 0) {
      masterSheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);
    }

    // 列幅を自動調整
    masterSheet.autoResizeColumns(1, headers.length);

    Logger.log('マスタシートを初期化しました。');
    Logger.log('サンプルデータを確認し、実際のURLに変更してください。');

  } catch (error) {
    Logger.log('マスタシート初期化エラー: ' + error.toString());
  }
}

/**
 * 未処理URLのみを処理する関数
 */
function processUnprocessedUrls() {
  const config = getConfig();

  try {
    // マスタシートからURL一覧を取得
    const allUrls = getUrlListFromMasterSheet(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME);

    // 未処理（last_crawled_atが空）のURLのみを抽出
    const unprocessedUrls = allUrls.filter(urlData => {
      return !urlData.lastCrawledAt || urlData.lastCrawledAt.toString().trim() === '';
    });

    if (unprocessedUrls.length === 0) {
      Logger.log('未処理のURLはありません。');
      return;
    }

    Logger.log(`未処理URL数: ${unprocessedUrls.length}件`);

    // 未処理URLを処理
    let processedCount = 0;
    let errorCount = 0;

    for (let i = 0; i < unprocessedUrls.length; i++) {
      const urlData = unprocessedUrls[i];

      try {
        Logger.log(`未処理URL処理中 ${i + 1}/${unprocessedUrls.length}: ID=${urlData.id}`);

        // スクレイピング実行
        const hotelData = scrapeHotelFromUrl(urlData.url);

        // 結果シートに書き込み
        writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, hotelData);

        // マスタシートの最終クロール日時を更新
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date());

        processedCount++;

        // リクエスト間隔を空ける
        if (i < unprocessedUrls.length - 1) {
          Utilities.sleep(config.REQUEST_DELAY);
        }

      } catch (error) {
        Logger.log(`未処理URL処理エラー ID=${urlData.id}: ${error.toString()}`);

        // エラー情報を結果シートに記録
        const errorData = {
          title: 'ERROR',
          score: '',
          total: '',
          address: '',
          tel: '',
          totalRooms: '',
          url: urlData.url,
          error: error.toString()
        };

        writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, errorData);
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date(), error.toString());

        errorCount++;
      }
    }

    Logger.log(`未処理URL処理完了: 成功 ${processedCount}件, エラー ${errorCount}件`);

  } catch (error) {
    Logger.log('未処理URL処理エラー: ' + error.toString());
  }
}

/**
 * 指定期間より古いデータの再処理
 */
function reprocessOldData(daysOld = 7) {
  const config = getConfig();

  try {
    // マスタシートからURL一覧を取得
    const allUrls = getUrlListFromMasterSheet(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME);

    // 指定日数より古いデータを抽出
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);

    const oldUrls = allUrls.filter(urlData => {
      if (!urlData.lastCrawledAt || urlData.lastCrawledAt.toString().trim() === '') {
        return true; // 未処理も対象
      }

      try {
        const crawledDate = new Date(urlData.lastCrawledAt);
        return crawledDate < cutoffDate;
      } catch (error) {
        return true; // 日付が不正な場合も対象
      }
    });

    if (oldUrls.length === 0) {
      Logger.log(`${daysOld}日より古いデータはありません。`);
      return;
    }

    Logger.log(`${daysOld}日より古いURL数: ${oldUrls.length}件`);

    // 古いデータを再処理
    for (let i = 0; i < oldUrls.length; i++) {
      const urlData = oldUrls[i];

      try {
        Logger.log(`再処理中 ${i + 1}/${oldUrls.length}: ID=${urlData.id}`);

        const hotelData = scrapeHotelFromUrl(urlData.url);
        writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, hotelData);
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date());

        if (i < oldUrls.length - 1) {
          Utilities.sleep(config.REQUEST_DELAY);
        }

      } catch (error) {
        Logger.log(`再処理エラー ID=${urlData.id}: ${error.toString()}`);

        const errorData = {
          title: 'ERROR',
          score: '',
          total: '',
          address: '',
          tel: '',
          totalRooms: '',
          url: urlData.url,
          error: error.toString()
        };

        writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, errorData);
        updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date(), error.toString());
      }
    }

    Logger.log('再処理完了');

  } catch (error) {
    Logger.log('再処理エラー: ' + error.toString());
  }
}

/**
 * 実行状況管理機能付きメイン処理
 */
function mainWithProgressTracking() {
  const startTime = new Date().getTime();
  const config = getConfig();

  // 設定チェック
  if (config.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    Logger.log('エラー: スプレッドシートIDが設定されていません。');
    return;
  }

  if (!config.MASTER_SHEET_NAME) {
    Logger.log('エラー: マスタシート名が設定されていません。');
    return;
  }

  try {
    // 実行状況を取得
    const progress = getExecutionProgress();
    Logger.log(`=== 実行状況管理付き処理開始 ===`);
    Logger.log(`開始時点での処理済み件数: ${progress.totalProcessed}`);

    const result = processUnprocessedUrlsWithProgressTracking(startTime, progress);

    if (result.hasMoreUrls) {
      Logger.log(`実行時間制限により処理中断。今回処理件数: ${result.processedThisRun}件`);
      Logger.log(`累計処理件数: ${progress.totalProcessed + result.processedThisRun}件`);

      // 進捗を保存
      saveExecutionProgress(progress.totalProcessed + result.processedThisRun);

      // 次回実行をスケジュール
      scheduleNextExecutionWithProgress();
    } else {
      Logger.log(`全処理完了。今回処理件数: ${result.processedThisRun}件`);
      Logger.log(`累計処理件数: ${progress.totalProcessed + result.processedThisRun}件`);

      // 進捗をリセット
      clearExecutionProgress();

      // トリガーをクリア
      clearProgressTriggers();
    }

  } catch (error) {
    Logger.log('メイン処理エラー: ' + error.toString());
  }
}

/**
 * 進捗管理付き未処理URL処理
 */
function processUnprocessedUrlsWithProgressTracking(startTime, progress) {
  const config = getConfig();

  // 未処理URLを取得
  const unprocessedUrls = getUnprocessedUrls(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME);

  if (unprocessedUrls.length === 0) {
    Logger.log('未処理のURLはありません。');
    return { processedThisRun: 0, hasMoreUrls: false };
  }

  Logger.log(`未処理URL数: ${unprocessedUrls.length}件`);

  let processedThisRun = 0;

  for (let i = 0; i < unprocessedUrls.length; i++) {
    // 実行時間をチェック
    const currentTime = new Date().getTime();
    const elapsedTime = (currentTime - startTime) / 1000;

    if (elapsedTime > (EXECUTION_TIME_LIMIT - SAFETY_MARGIN)) {
      Logger.log(`実行時間制限に近づきました。経過時間: ${Math.round(elapsedTime)}秒`);
      return { processedThisRun, hasMoreUrls: true };
    }

    const urlData = unprocessedUrls[i];

    try {
      Logger.log(`処理中 ${i + 1}/${unprocessedUrls.length}: ID=${urlData.id} (経過: ${Math.round(elapsedTime)}秒, 累計: ${progress.totalProcessed + processedThisRun})`);

      // スクレイピング実行
      const hotelData = scrapeHotelFromUrl(urlData.url);

      // 結果シートに書き込み
      writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, hotelData);

      // マスタシートの最終クロール日時を更新
      updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date());

      processedThisRun++;

      // 10件ごとに進捗を保存
      if (processedThisRun % 10 === 0) {
        saveExecutionProgress(progress.totalProcessed + processedThisRun);
        Logger.log(`進捗保存: ${progress.totalProcessed + processedThisRun}件処理済み`);
      }

      // リクエスト間隔を空ける
      if (i < unprocessedUrls.length - 1) {
        Utilities.sleep(config.REQUEST_DELAY);
      }

    } catch (error) {
      Logger.log(`URL処理エラー ID=${urlData.id}: ${error.toString()}`);

      // エラー情報を記録
      const errorData = {
        title: 'ERROR',
        score: '',
        total: '',
        address: '',
        tel: '',
        totalRooms: '',
        url: urlData.url,
        error: error.toString()
      };

      writeToSpreadsheetAdvanced(config.SPREADSHEET_ID, config.SHEET_NAME, errorData);
      updateLastCrawledAt(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME, urlData.rowIndex, new Date(), error.toString());

      processedThisRun++;
    }
  }

  return { processedThisRun, hasMoreUrls: false };
}

/**
 * 実行進捗を取得
 */
function getExecutionProgress() {
  const properties = PropertiesService.getScriptProperties();
  const totalProcessed = parseInt(properties.getProperty('TOTAL_PROCESSED') || '0');
  const lastExecutionTime = properties.getProperty('LAST_EXECUTION_TIME');

  return {
    totalProcessed: totalProcessed,
    lastExecutionTime: lastExecutionTime
  };
}

/**
 * 実行進捗を保存
 */
function saveExecutionProgress(totalProcessed) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'TOTAL_PROCESSED': totalProcessed.toString(),
    'LAST_EXECUTION_TIME': new Date().toISOString()
  });
}

/**
 * 実行進捗をクリア
 */
function clearExecutionProgress() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('TOTAL_PROCESSED');
  properties.deleteProperty('LAST_EXECUTION_TIME');
  Logger.log('実行進捗をクリアしました。');
}

/**
 * 進捗管理付き次回実行スケジュール
 */
function scheduleNextExecutionWithProgress() {
  try {
    // 既存のトリガーを削除
    clearProgressTriggers();

    // 1分後に実行するトリガーを作成
    const triggerTime = new Date();
    triggerTime.setMinutes(triggerTime.getMinutes() + 1);

    ScriptApp.newTrigger('mainWithProgressTracking')
      .timeBased()
      .at(triggerTime)
      .create();

    Logger.log(`次回実行時刻をスケジュール: ${triggerTime.toLocaleString()}`);

  } catch (error) {
    Logger.log('トリガー設定エラー: ' + error.toString());
  }
}

/**
 * 進捗管理用トリガーをクリア
 */
function clearProgressTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'mainWithProgressTracking') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  if (deletedCount > 0) {
    Logger.log(`${deletedCount}個の進捗管理トリガーを削除しました。`);
  }
}

/**
 * 実行状況表示
 */
function showExecutionStatus() {
  const progress = getExecutionProgress();
  const config = getConfig();

  Logger.log('=== 実行状況 ===');
  Logger.log(`累計処理件数: ${progress.totalProcessed}件`);
  Logger.log(`最終実行時刻: ${progress.lastExecutionTime || '未実行'}`);

  try {
    const unprocessedUrls = getUnprocessedUrls(config.SPREADSHEET_ID, config.MASTER_SHEET_NAME);
    Logger.log(`残り未処理件数: ${unprocessedUrls.length}件`);

    const triggers = ScriptApp.getProjectTriggers().filter(t =>
      t.getHandlerFunction() === 'mainWithProgressTracking'
    );
    Logger.log(`スケジュール済みトリガー数: ${triggers.length}件`);

  } catch (error) {
    Logger.log('状況取得エラー: ' + error.toString());
  }
}

/**
 * 統合日次実行管理 - 新着ホテル収集とスクレイピングを統合管理
 */

/**
 * 統合日次実行トリガーを設定
 */
function setupDailyIntegratedTrigger() {
  // 既存のトリガーをクリア
  clearAllDailyTriggers();

  const config = getConfig();

  // 統合実行用トリガー（朝9時30分）
  ScriptApp.newTrigger('mainWithNewHotelCollection')
    .timeBased()
    .everyDays(1)
    .atHour(config.DAILY_EXECUTION_HOUR)
    .nearMinute(30)
    .create();

  Logger.log(`統合日次トリガーを設定しました:`);
  Logger.log(`- 新着ホテル収集: 毎日 ${config.DAILY_EXECUTION_HOUR}:00`);
  Logger.log(`- 統合実行: 毎日 ${config.DAILY_EXECUTION_HOUR}:30`);
}

/**
 * 全ての日次トリガーをクリア
 */
function clearAllDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const dailyFunctions = [
    'mainWithNewHotelCollection',
    'mainWithProgressTracking'
  ];

  let clearedCount = 0;
  triggers.forEach(trigger => {
    if (dailyFunctions.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      clearedCount++;
    }
  });

  Logger.log(`${clearedCount} 個の日次トリガーをクリアしました`);
}

/**
 * 総合ステータス確認
 */
function showIntegratedStatus() {
  Logger.log('=== 統合ステータス確認 ===');

  // 1. スクレイピング状況
  Logger.log('\n【スクレイピング状況】');
  showExecutionStatus();

  // 2. 新着ホテル収集状況
  Logger.log('\n【新着ホテル収集状況】');
  checkNewHotelCollectionStatus();

  // 3. トリガー状況
  Logger.log('\n【トリガー状況】');
  const triggers = ScriptApp.getProjectTriggers();
  const dailyTriggers = triggers.filter(trigger => {
    const funcName = trigger.getHandlerFunction();
    return funcName.includes('daily') || funcName.includes('main');
  });

  if (dailyTriggers.length === 0) {
    Logger.log('日次実行トリガーが設定されていません');
  } else {
    dailyTriggers.forEach(trigger => {
      Logger.log(`- ${trigger.getHandlerFunction()}: ${trigger.getTriggerSource()}`);
    });
  }
}
