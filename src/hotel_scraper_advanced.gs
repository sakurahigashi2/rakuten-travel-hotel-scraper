/**
 * 楽天トラベルホテル情報スクレイピング（エラーハンドリング強化版）
 */

/**
 * メイン実行関数
 */
function main() {
  const config = getConfig();

  // スプレッドシートIDが設定されているかチェック
  if (config.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    Logger.log('エラー: スプレッドシートIDが設定されていません。config.gsでSPREADSHEET_IDを設定してください。');
    return;
  }

  try {
    // 単一URLの場合
    if (config.URLS.length === 1) {
      scrapeHotelInfoSingle(config.URLS[0], config.SPREADSHEET_ID, config.SHEET_NAME);
    } else {
      // 複数URLの場合
      scrapeHotelInfoBatch(config.URLS, config.SPREADSHEET_ID, config.SHEET_NAME);
    }
  } catch (error) {
    Logger.log('メイン処理エラー: ' + error.toString());
  }
}

/**
 * 単一ホテル情報取得
 */
function scrapeHotelInfoSingle(url, spreadsheetId, sheetName) {
  try {
    Logger.log('処理開始: ' + url);

    const response = UrlFetchApp.fetch(url, {
      timeout: getConfig().TIMEOUT * 1000
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('HTTP エラー: ' + response.getResponseCode());
    }

    const html = response.getContentText();
    const hotelData = extractHotelInfoAdvanced(html, url);

    writeToSpreadsheetAdvanced(spreadsheetId, sheetName, hotelData);

    Logger.log('処理完了: ' + JSON.stringify(hotelData));

  } catch (error) {
    Logger.log('単一処理エラー: ' + error.toString());

    // エラー情報もスプレッドシートに記録
    const errorData = {
      title: 'ERROR',
      score: '',
      total: '',
      address: '',
      tel: '',
      totalRooms: '',
      url: url,
      error: error.toString()
    };

    writeToSpreadsheetAdvanced(spreadsheetId, sheetName, errorData);
  }
}

/**
 * バッチ処理
 */
function scrapeHotelInfoBatch(urls, spreadsheetId, sheetName) {
  Logger.log(`バッチ処理開始: ${urls.length}件のURL`);

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < urls.length; i++) {
    try {
      Logger.log(`処理中: ${i + 1}/${urls.length} - ${urls[i]}`);

      scrapeHotelInfoSingle(urls[i], spreadsheetId, sheetName);
      successCount++;

      // 最後以外はリクエスト間隔を空ける
      if (i < urls.length - 1) {
        Utilities.sleep(getConfig().REQUEST_DELAY);
      }

    } catch (error) {
      Logger.log(`URL処理エラー ${urls[i]}: ${error.toString()}`);
      errorCount++;
    }
  }

  Logger.log(`バッチ処理完了: 成功 ${successCount}件, エラー ${errorCount}件`);
}

/**
 * ホテル情報抽出（エラーハンドリング強化版）
 */
function extractHotelInfoAdvanced(html, url) {
  const hotelData = {
    title: '',
    score: '',
    total: '',
    address: '',
    tel: '',
    totalRooms: '',
    url: url,
    error: ''
  };

  try {
    // タイトル（施設名）を抽出
    const titlePatterns = [
      /<title>(.*?)【楽天トラベル】<\/title>/,
      /<title>(.*?)\s+設備・アメニティ・基本情報【楽天トラベル】<\/title>/,
      /<title>(.*?)<\/title>/
    ];

    for (const pattern of titlePatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.title = match[1]
          .replace(/\s+設備・アメニティ・基本情報/, '')
          .replace(/\s+【楽天トラベル】/, '')
          .trim();
        break;
      }
    }

    // 評価（score）を抽出
    const scorePatterns = [
      /hotelinfoheader\.score\s*=\s*([\d.]+);/,
      /"score"\s*:\s*([\d.]+)/,
      /評価\s*([\d.]+)/
    ];

    for (const pattern of scorePatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.score = parseFloat(match[1]);
        break;
      }
    }

    // 評価件数（total）を抽出
    const totalPatterns = [
      /hotelinfoheader\.total\s*=\s*"([^"]+)";/,
      /"total"\s*:\s*"([^"]+)"/,
      /全(\d+[,\d]*)件/
    ];

    for (const pattern of totalPatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.total = match[1];
        break;
      }
    }

    // 住所を抽出
    const addressPatterns = [
      /<span class="header__hotel-address"[^>]*>(.*?)<\/span>/,
      /<dt>住所<\/dt>\s*<dd>([^<]+)<\/dd>/,
      /〒[\d-]+[^<]+/
    ];

    for (const pattern of addressPatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.address = match[1].trim();
        break;
      }
    }

    // 電話番号を抽出
    const telPatterns = [
      /<dt>TEL<\/dt>\s*<dd>([^<]+)<\/dd>/,
      /TEL[:\s]*(\d{2,4}-\d{2,4}-\d{4})/,
      /電話[:\s]*(\d{2,4}-\d{2,4}-\d{4})/
    ];

    for (const pattern of telPatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.tel = match[1].trim();
        break;
      }
    }

    // 総部屋数を抽出
    const roomsPatterns = [
      /<dt>総部屋数<\/dt>\s*<dd>([^<]+)<\/dd>/,
      /総部屋数[:\s]*(\d+[^<]*)/,
      /(\d+)室/
    ];

    for (const pattern of roomsPatterns) {
      const match = html.match(pattern);
      if (match) {
        hotelData.totalRooms = match[1].trim();
        break;
      }
    }

  } catch (error) {
    hotelData.error = error.toString();
    Logger.log('データ抽出エラー: ' + error.toString());
  }

  return hotelData;
}

/**
 * スプレッドシート書き込み（エラーハンドリング強化版）
 */
function writeToSpreadsheetAdvanced(spreadsheetId, sheetName, data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);

      // ヘッダー行を作成
      const headers = ['施設名', '評価', '評価件数', '住所', 'TEL', '総部屋数', 'URL', 'エラー', '取得日時'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // ヘッダーのスタイルを設定
      formatHeader(sheet, headers.length);
    }

    // データを追加
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    const rowData = [
      data.title || '',
      data.score || '',
      data.total || '',
      data.address || '',
      data.tel || '',
      data.totalRooms || '',
      data.url || '',
      data.error || '',
      new Date()
    ];

    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // エラー行の場合は背景色を変更（該当行のみ）
    if (data.error) {
      const errorRange = sheet.getRange(newRow, 1, 1, rowData.length);
      errorRange.setBackground('#ffcccc');
    } else {
      // 正常行の場合は明示的に白色に設定
      const normalRange = sheet.getRange(newRow, 1, 1, rowData.length);
      normalRange.setBackground('#ffffff');
    }

    // 列幅を自動調整（最初の10行のデータがある場合のみ）
    if (newRow <= 10) {
      sheet.autoResizeColumns(1, rowData.length);
    }

    Logger.log('スプレッドシートに書き込み完了');

  } catch (error) {
    Logger.log('スプレッドシート書き込みエラー: ' + error.toString());
    throw error;
  }
}

/**
 * ヘッダーのフォーマット
 */
function formatHeader(sheet, columnCount) {
  const headerRange = sheet.getRange(1, 1, 1, columnCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setBorder(true, true, true, true, true, true);
}
