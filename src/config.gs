/**
 * 設定ファイル
 * Google Apps Scriptホテル情報スクレイピング用
 */

// スプレッドシート設定
const CONFIG = {
  // ここにあなたのスプレッドシートIDを入力してください
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',

  // 結果出力用シート名
  SHEET_NAME: '宿泊施設情報',

  // マスタシート名（URL一覧管理用）
  MASTER_SHEET_NAME: 'URL_マスタ',

  // 新着ホテル取得用の設定
  NEW_HOTELS_URL: 'https://travel.rakuten.co.jp/whatsnew.html',
  NEW_HOTELS_SHEET_NAME: '新着宿泊施設情報',

  // リクエスト間隔（ミリ秒）
  REQUEST_DELAY: 2000,

  // タイムアウト時間（秒）
  TIMEOUT: 30,

  // 日次実行用の設定
  DAILY_EXECUTION_HOUR: 9, // 毎日9時に実行

  // 対象URL一覧（マスタシート使用時は不要）
  URLS: []
};

/**
 * 設定を取得する関数
 */
function getConfig() {
  return CONFIG;
}

/**
 * スプレッドシートIDを設定する関数
 */
function setSpreadsheetId(id) {
  CONFIG.SPREADSHEET_ID = id;
}

// ========================================
// セットアップ・初期化機能
// ========================================

/**
 * クイックセットアップ：初回設定用
 * マスタシート初期化、新着ホテル収集、トリガー設定を一括実行
 */
function quickSetup() {
  Logger.log('=== クイックセットアップ開始 ===');

  try {
    // 1. マスタシート初期化
    Logger.log('Step 1: マスタシート初期化中...');
    initializeMasterSheet();

    // 2. 日次トリガー設定
    Logger.log('Step 2: 日次トリガー設定中...');
    setupDailyIntegratedTrigger();

    // 3. エラーURL再実行トリガー設定
    Logger.log('Step 3: エラーURL再実行トリガー設定中...');
    setupErrorRetryTrigger();

    Logger.log('=== クイックセットアップ完了 ===');
    Logger.log('これで毎日自動的に新着ホテルの収集、スクレイピング、エラーURL再実行が実行されます');

  } catch (error) {
    Logger.log(`クイックセットアップエラー: ${error.toString()}`);
  }
}

/**
 * トリガーのみセットアップ
 */
function quickSetupTriggerOnly() {
  Logger.log('=== トリガーのみセットアップ開始 ===');

  try {
    // 日次トリガー設定
    setupDailyIntegratedTrigger();

    // エラーURL再実行トリガー設定
    setupErrorRetryTrigger();

    Logger.log('=== トリガーセットアップ完了 ===');

  } catch (error) {
    Logger.log(`トリガーセットアップエラー: ${error.toString()}`);
  }
}
