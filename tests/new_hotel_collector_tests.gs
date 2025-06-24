/**
 * 新着ホテル収集機能のテスト用ファイル
 * 実際の実行前に動作確認を行うためのテスト関数群
 */

/**
 * 新着ホテル収集のテスト実行
 */
function testNewHotelCollection() {
  Logger.log('=== 新着ホテル収集テスト開始 ===');

  try {
    // 設定確認
    const config = getConfig();
    Logger.log(`対象URL: ${config.NEW_HOTELS_URL}`);
    Logger.log(`スプレッドシートID: ${config.SPREADSHEET_ID}`);

    // 新着ホテル情報取得テスト
    const newHotels = fetchNewHotelsData(config.NEW_HOTELS_URL);

    if (newHotels && newHotels.length > 0) {
      Logger.log(`取得した新着ホテル数: ${newHotels.length} 件`);

      // 最初の3件のみ表示（ログが長くなりすぎるのを防ぐ）
      const displayCount = Math.min(3, newHotels.length);
      for (let i = 0; i < displayCount; i++) {
        const hotel = newHotels[i];
        Logger.log(`${i + 1}. ${hotel.name}`);
        Logger.log(`   URL: ${hotel.url}`);
        Logger.log(`   ID: ${hotel.id}`);
      }

      if (newHotels.length > 3) {
        Logger.log(`   ... 他 ${newHotels.length - 3} 件`);
      }
    } else {
      Logger.log('新着ホテル情報が取得できませんでした');
    }

  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.toString()}`);
  }

  Logger.log('=== 新着ホテル収集テスト終了 ===');
}

/**
 * HTML解析テスト（デバッグ用）
 */
function testHtmlParsing() {
  Logger.log('=== HTML解析テスト開始 ===');

  try {
    const config = getConfig();
    const response = UrlFetchApp.fetch(config.NEW_HOTELS_URL);
    const html = response.getContentText('UTF-8');

    // HTMLの一部を表示（最初の1000文字）
    Logger.log('取得したHTML（最初の1000文字）:');
    Logger.log(html.substring(0, 1000));

    // 新着ホテルセクションの検索
    const sectionPattern = /最近登録いただいた宿泊施設の一覧[\s\S]*?(?=<\/div>|<div[^>]*class="[^"]*section|$)/i;
    const sectionMatch = html.match(sectionPattern);

    if (sectionMatch) {
      Logger.log('新着ホテルセクションが見つかりました');
      Logger.log('セクション内容（最初の500文字）:');
      Logger.log(sectionMatch[0].substring(0, 500));
    } else {
      Logger.log('新着ホテルセクションが見つかりませんでした');

      // より広い範囲で検索
      const broadPattern = /宿泊施設/gi;
      const matches = html.match(broadPattern);
      if (matches) {
        Logger.log(`「宿泊施設」というキーワードが ${matches.length} 回見つかりました`);
      }
    }

  } catch (error) {
    Logger.log(`HTML解析テストエラー: ${error.toString()}`);
  }

  Logger.log('=== HTML解析テスト終了 ===');
}

/**
 * スプレッドシート接続テスト
 */
function testSpreadsheetConnection() {
  Logger.log('=== スプレッドシート接続テスト開始 ===');

  try {
    const config = getConfig();

    if (config.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      Logger.log('エラー: スプレッドシートIDが設定されていません');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    Logger.log(`スプレッドシート名: ${spreadsheet.getName()}`);

    // マスタシートの確認
    let masterSheet = spreadsheet.getSheetByName(config.MASTER_SHEET_NAME);
    if (masterSheet) {
      const rowCount = masterSheet.getLastRow();
      Logger.log(`マスタシート「${config.MASTER_SHEET_NAME}」が存在します（${rowCount} 行）`);
    } else {
      Logger.log(`マスタシート「${config.MASTER_SHEET_NAME}」が存在しません`);
    }

    // 新着ホテルログシートの確認
    let logSheet = spreadsheet.getSheetByName(config.NEW_HOTELS_SHEET_NAME);
    if (logSheet) {
      const rowCount = logSheet.getLastRow();
      Logger.log(`新着ホテルログシート「${config.NEW_HOTELS_SHEET_NAME}」が存在します（${rowCount} 行）`);
    } else {
      Logger.log(`新着ホテルログシート「${config.NEW_HOTELS_SHEET_NAME}」が存在しません`);
    }

  } catch (error) {
    Logger.log(`スプレッドシート接続テストエラー: ${error.toString()}`);
  }

  Logger.log('=== スプレッドシート接続テスト終了 ===');
}

/**
 * 統合テスト（全機能の動作確認）
 */
function runIntegratedTest() {
  Logger.log('=== 統合テスト開始 ===');

  // 1. 設定確認
  Logger.log('Step 1: 設定確認');
  testSpreadsheetConnection();

  Utilities.sleep(1000);

  // 2. HTML取得・解析テスト
  Logger.log('Step 2: HTML取得・解析テスト');
  testHtmlParsing();

  Utilities.sleep(1000);

  // 3. 新着ホテル収集テスト
  Logger.log('Step 3: 新着ホテル収集テスト');
  testNewHotelCollection();

  Logger.log('=== 統合テスト終了 ===');
}

/**
 * セットアップ確認
 */
function checkSetup() {
  Logger.log('=== セットアップ確認 ===');

  const config = getConfig();

  // 必須設定の確認
  const checks = [
    {
      name: 'スプレッドシートID',
      value: config.SPREADSHEET_ID,
      valid: config.SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE'
    },
    {
      name: '新着ホテルURL',
      value: config.NEW_HOTELS_URL,
      valid: config.NEW_HOTELS_URL && config.NEW_HOTELS_URL.includes('whatsnew.html')
    },
    {
      name: 'マスタシート名',
      value: config.MASTER_SHEET_NAME,
      valid: !!config.MASTER_SHEET_NAME
    },
    {
      name: '新着ホテルシート名',
      value: config.NEW_HOTELS_SHEET_NAME,
      valid: !!config.NEW_HOTELS_SHEET_NAME
    }
  ];

  let allValid = true;
  checks.forEach(check => {
    const status = check.valid ? '✓' : '✗';
    Logger.log(`${status} ${check.name}: ${check.value}`);
    if (!check.valid) allValid = false;
  });

  if (allValid) {
    Logger.log('✓ セットアップは正常です');
  } else {
    Logger.log('✗ セットアップに問題があります。config.gsを確認してください');
  }

  Logger.log('=== セットアップ確認終了 ===');
}

/**
 * URL正規化テスト
 */
function testUrlNormalization() {
  Logger.log('=== URL正規化テスト開始 ===');

  // テスト用のURL例
  const testUrls = [
    'https://travel.rakuten.co.jp/HOTEL/104529/',
    '/HOTEL/104529/',
    'https://travel.rakuten.co.jp/HOTEL/104529/104529.html',
    'https://travel.rakuten.co.jp/HOTEL/104529/104529_std.html',
    'https://travel.rakuten.co.jp/HOTEL/104529/something_else.html'
  ];

  testUrls.forEach(testUrl => {
    const normalizedUrl = normalizeHotelUrl(testUrl);
    Logger.log(`${testUrl} → ${normalizedUrl}`);
  });

  Logger.log('=== URL正規化テスト終了 ===');
}
