# 楽天トラベル ホテル情報スクレイピング - Google Apps Script

## 概要

このスクリプトは楽天トラベルのホテル詳細ページから以下の情報を抽出し、Googleスプレッドシートに保存します：

- 施設名（タイトル）
- 評価（hotelinfoheader.score）
- 評価件数（hotelinfoheader.total）
- 住所
- TEL（電話番号）
- 総部屋数

## ホテルマスターデータ

プロジェクトには2025年6月23日時点の楽天トラベルホテル情報データが含まれています：

- **ファイル**: `docs/hotel_master_20250623.csv`
- **データ件数**: 41,471件のホテル情報
- **フォーマット**: CSV形式（id, url）
- **内容**: 楽天トラベルの全ホテルIDとURLの一覧
- **データソース**: `https://hotel.travel.rakuten.co.jp/sitemap_domestic-ja-hotel-info.xml.gz`

このマスターデータは楽天トラベルの公式サイトマップから取得され、`hotel_scraper_master.gs`で活用されて大量のホテル情報を効率的にスクレイピングできます。

## ディレクトリ構成

```text
rakuten-travel-hotel-scraper/
├── src/                    # ソースコード
│   ├── appsscript.json    # Google Apps Script設定ファイル
│   ├── config.gs          # 設定ファイル（クイックセットアップ機能含む）
│   ├── hotel_scraper_advanced.gs  # エラーハンドリングとログ機能を強化したスクレイピング機能
│   ├── hotel_scraper_master.gs    # マスタシート連携版（URL一覧をマスタシートから取得、HTML取得・エンコーディング処理含む）
│   ├── new_hotel_collector.gs     # 新着ホテル情報自動収集機能
│   ├── error_url_retry.gs         # エラーURL再実行機能
├── docs/                   # ドキュメント
│   └── hotel_master_20250623.csv  # ホテルマスターデータ（41,471件）
├── scripts/               # 実行スクリプト
│   └── setup.sh          # セットアップスクリプト
├── tests/                 # テストファイル
│   └── new_hotel_collector_tests.gs # 新着ホテル収集機能のテスト関数
├── .clasp.json           # clasp設定ファイル
├── .gitignore           # Git除外設定
├── package.json         # プロジェクト設定
└── README.md           # プロジェクト説明（このファイル）
```

## セットアップ手順

### 1. 必要な環境

- Node.js (v14以上)
- npm または yarn
- Google アカウント
- **Google Apps Script CLI (clasp)** - 開発時に推奨
- **Git** - バージョン管理用
- **VS Code** - 推奨エディタ

### 2. プロジェクトのセットアップ

#### 自動セットアップ（推奨）

```bash
# リポジトリをクローン
git clone <repository-url>
cd rakuten-travel-hotel-scraper

# 依存関係をインストール
npm install

# clasp にログイン（開発時）
npx clasp login

# セットアップスクリプトを実行
./scripts/setup.sh
```

#### 手動セットアップ

```bash
# 依存関係をインストール
npm install

# セットアップスクリプトを実行
./scripts/setup.sh
```

#### 3. Google Apps Scriptプロジェクトの作成

1. [Google Apps Script](https://script.google.com/)にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を設定（例：「ホテル情報スクレイピング」）

#### 4. スクリプトファイルのアップロード

##### claspを使用する場合（推奨）

```bash
# Google Apps Script CLIをインストール
npm install -g @google/clasp

# ログイン
clasp login

# ファイルをアップロード
clasp push
```

##### 手動アップロードの場合

1. デフォルトの`Code.gs`を削除
2. `src/`ディレクトリ内の以下のファイルを追加：
   - `config.gs`
   - `hotel_scraper_advanced.gs`
   - `hotel_scraper_master.gs`（マスタシート連携版）
   - `new_hotel_collector.gs`（新着ホテル収集機能）
   - `error_url_retry.gs`（エラーURL再実行機能）

### 5. Googleスプレッドシートの準備

1. 新しいGoogleスプレッドシートを作成
2. スプレッドシートのURLからIDを取得
   - URL例：`https://docs.google.com/spreadsheets/d/【SPREADSHEET_ID】/edit`
3. `src/config.gs`の`SPREADSHEET_ID`に設定

### 6. 設定ファイルの編集

`src/config.gs`を以下のように設定してください：

```javascript
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
```

## 使用方法

### 設定・セットアップ機能（config.gs）

#### クイックセットアップ機能

初回設定を簡単に行うための機能群です。

```javascript
// 一括初期設定（推奨）
function quickSetup() {
  // マスタシート初期化 + 新着ホテル収集 + トリガー設定 + エラーURL再実行トリガー設定
}

// トリガーのみ設定
function quickSetupTriggerOnly() {
  // 日次トリガー + エラーURL再実行トリガーのみ設定
}
```

#### config.gs の関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `getConfig()` | 設定取得 | アプリケーション設定を取得 |
| `setSpreadsheetId(id)` | スプレッドシートID設定 | スプレッドシートIDを動的に設定 |
| `quickSetup()` | 一括初期設定 | マスタシート初期化＋新着ホテル収集＋トリガー設定＋エラーURL再実行トリガー設定 |
| `quickSetupTriggerOnly()` | トリガーのみ設定 | 日次トリガー＋エラーURL再実行トリガーのみを設定 |

### hotel_scraper_advanced.gs

#### 基本機能

エラーハンドリングとログ機能を強化したスクレイピング機能です。楽天トラベルのホテル詳細ページから施設情報を自動抽出し、Googleスプレッドシートに保存します。

**主な特徴：**

- 複数パターンの正規表現を使用したデータ抽出の堅牢性
- エラーハンドリング機能による安定した処理
- 単一URL処理とバッチ処理の両対応
- スプレッドシートの自動フォーマッティング
- リクエスト間隔調整による楽天トラベルサーバーへの配慮
- エラー行の視覚的な区別（背景色変更）

**抽出可能なデータ：**

- 施設名（ホテル・旅館名）
- 評価（星評価・数値）
- 評価件数（レビュー総数）
- 住所（所在地）
- 電話番号（TEL情報）
- 総部屋数（客室数）
- URL（対象ページのURL）
- エラー情報（処理中に発生したエラー）
- 取得日時（データ取得タイムスタンプ）

#### hotel_scraper_advanced.gs の関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `main()` | メイン実行関数 | config.gsの設定に基づき、記載したURLのスクレイピングを実行 |
| `scrapeHotelInfoSingle(url, spreadsheetId, sheetName)` | 単一URLスクレイピング | 指定したURLのホテル情報を取得 |
| `scrapeHotelInfoBatch(urls, spreadsheetId, sheetName)` | 複数URLバッチ処理 | URL配列を順次処理してデータを取得 |
| `extractHotelInfoAdvanced(html, url)` | HTML解析・データ抽出 | HTMLからホテル情報を抽出 |
| `writeToSpreadsheetAdvanced(spreadsheetId, sheetName, data)` | スプレッドシート書き込み | 抽出したデータをスプレッドシートに保存 |
| `formatHeader(sheet, columnCount)` | ヘッダー書式設定 | スプレッドシートのヘッダー行を装飾 |

### マスタシート連携版（hotel_scraper_master.gs）

#### マスタシート連携の主要機能

```javascript
// 進捗管理付き実行時間制限対応版（推奨）
function mainWithProgressTracking() {
  // 進捗を保存しながら実行時間制限に対応
}

// 新着ホテル収集 + スクレイピング統合実行（🆕最新機能）
function mainWithNewHotelCollection() {
  // 新着ホテル収集とスクレイピングを統合実行
}
```

#### hotel_scraper_master.gs の関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `mainWithNewHotelCollection()` | 新着ホテル収集＋スクレイピング統合実行 | 新着ホテル収集とスクレイピングを一括で実行 |
| `mainWithProgressTracking()` | 進捗管理付きメイン処理 | 実行時間制限対応、自動継続機能付き |
| `processUnprocessedUrls()` | 未処理URL処理 | last_crawled_atが空のURLのみを処理 |
| `reprocessOldData(daysOld)` | 古いデータ再処理 | 指定日数より古いデータを再処理 |
| `setupDailyIntegratedTrigger()` | 日次統合トリガー設定 | 新着ホテル収集とスクレイピングの自動実行設定 |
| `showIntegratedStatus()` | 総合ステータス確認 | スクレイピング・新着ホテル収集・トリガーの状況表示 |
| `showExecutionStatus()` | 実行状況確認 | 処理済み件数や残り件数を表示 |
| `initializeMasterSheet()` | マスタシート初期化 | マスタシートをサンプルデータで初期化 |
| `clearProgressTriggers()` | トリガークリア | 自動実行トリガーを削除 |
| `getUnprocessedUrls(spreadsheetId, masterSheetName)` | 未処理URL取得 | マスタシートから未処理URLを抽出 |
| `updateLastCrawledAt(spreadsheetId, masterSheetName, rowIndex, crawledAt, errorMessage)` | 処理日時更新 | マスタシートの処理日時を更新 |
| `getHtmlWithCorrectEncoding(url, options)` | エンコーディング対応HTML取得 | 適切なエンコーディングでHTMLを取得 |
| `detectEncodingFromHtml(response)` | エンコーディング検出 | HTMLからエンコーディングを自動検出 |
| `getMultipleHtmlWithEncoding(urls, delay)` | 複数HTML一括取得 | レート制限対応の複数URL処理 |

### 新着ホテル収集機能（new_hotel_collector.gs）

#### 新着ホテル収集の主要機能

```javascript
// 新着ホテル収集のみ実行
function collectNewHotels() {
  // 新着ホテル情報のみを収集してマスタシートに追加
}
```

#### new_hotel_collector.gs の関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `collectNewHotels()` | 新着ホテル情報収集 | 楽天トラベル新着ページから新しいホテル情報を取得 |
| `checkNewHotelCollectionStatus()` | 新着ホテル収集状況確認 | 最後の実行状況とトリガー状況を表示 |
| `extractNewHotelsFromHtml(html)` | HTML解析 | 新着情報ページからホテル情報を抽出 |
| `normalizeHotelUrl(url)` | URL正規化 | ホテルURLを統一形式に変換 |
| `checkHotelIdDuplicates(hotels)` | 重複チェック | 既存マスタシートとの重複確認 |

### エラーURL再実行機能（error_url_retry.gs）

#### エラーURL再実行の主要機能

結果シートでERRORになったURLを自動的に検出し、再実行・更新を行う機能です。

```javascript
// エラーURL再実行
function retryErrorUrls() {
  // 結果シートからERRORになったURLを検出して再実行
}

// 日次実行用（自動トリガーで実行される）
function dailyRetryErrorUrls() {
  // 毎日自動的にエラーURLを再実行
}
```

#### error_url_retry.gs の関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `retryErrorUrls()` | エラーURL再実行 | 結果シートのERRORになったURLを再実行・更新 |
| `getErrorUrlsFromResultSheet(id, name)` | エラーURL取得 | 結果シートからエラーURLリストを取得 |
| `updateResultSheetRow(id, name, row, data)` | 結果更新 | 結果シートの指定行をホテルデータで更新 |
| `findMasterSheetRowByUrl(id, name, url)` | マスタ検索 | マスタシートから指定URLの行番号を検索 |
| `dailyRetryErrorUrls()` | 日次エラー再実行 | エラーURL再実行（日次実行用） |
| `showErrorRetryStatus()` | エラー再実行状況表示 | エラーURL再実行の実行状況表示 |
| `setupErrorRetryTrigger()` | エラー再実行トリガー設定 | エラーURL再実行トリガー設定 |
| `clearErrorRetryTriggers()` | エラー再実行トリガークリア | エラーURL再実行トリガーをクリア |

#### 注意事項

- エラーURL再実行機能は、結果シートにERRORと表示されたURLのみを対象に再実行を行います。
- 日次実行用のトリガーを設定することで、毎日自動的にエラーURLの再実行が可能です。

### テスト機能（tests/new_hotel_collector_tests.gs）

#### テスト関数一覧

| 関数名 | 説明 | 用途 |
|--------|------|------|
| `testNewHotelCollection()` | 新着ホテル収集テスト | 全体的な動作テスト |
| `testHtmlParsing()` | HTML解析テスト | HTMLパースと正規表現のテスト |
| `testSpreadsheetConnection()` | スプレッドシート接続テスト | スプレッドシート読み書きのテスト |
| `runIntegratedTest()` | 統合テスト | 全機能の統合動作テスト |
| `checkSetup()` | セットアップ確認 | 設定と環境のチェック |
| `testUrlNormalization()` | URL正規化テスト | URL変換機能のテスト |

## マスタシート形式

マスタシートは以下の形式で作成してください：

| 列 | ヘッダー | 説明 | 例 |
|---|---|---|---|
| A | id | 識別ID | 001 |
| B | url | スクレイピング対象URL | `https://travel.rakuten.co.jp/HOTEL/104529/104529_std.html` |
| C | last_crawled_at | 最終処理日時 | 2024/12/20 10:30:00 |

### マスタシートの簡単作成方法

プロジェクトに含まれている`docs/hotel_master_20250623.csv`を活用することで、41,471件の全ホテルデータを含むマスタシートを簡単に作成できます：

1. **CSVファイルのインポート**:
   - Googleスプレッドシートで新しいシートを作成
   - `ファイル` → `インポート` → `アップロード`
   - `docs/hotel_master_20250623.csv`をアップロード
   - 区切り文字を「カンマ」に設定してインポート

2. **設定の更新**:
   - `src/config.gs`でマスタシートのスプレッドシートIDとシート名を設定

## 実行手順

### config.gsに記載のURLを対象にする場合

1. Google Apps Scriptエディタで`hotel_scraper_advanced.gs` > `main`関数を選択
2. 「実行」ボタンをクリック
3. 初回実行時は権限許可が必要
4. 実行ログで進行状況を確認
5. Googleスプレッドシートで結果を確認

### マスタシート連携版の場合

1. `initializeMasterSheet()`を実行してマスタシートを作成
2. マスタシートにスクレイピング対象のURLを入力
3. `mainWithProgressTracking()`を実行（推奨）
4. 大量のURLがある場合は自動的に継続実行される

### 新着ホテル自動収集機能の場合

#### 初回セットアップ（推奨）

1. **config.gs**の`quickSetup()`を実行して一括初期設定
   - マスタシート初期化
   - 新着ホテル初回収集  
   - 日次自動実行トリガー設定
   - エラーURL再実行トリガー設定
2. 設定完了後は自動的に毎日実行される

#### 個別セットアップ

- **マスタシートのみ**: `quickSetupMasterOnly()`
- **トリガーのみ**: `quickSetupTriggerOnly()` (日次トリガー + エラーURL再実行トリガー)

#### 手動実行

1. **新着ホテル収集のみ**: `collectNewHotels()`を実行
2. **統合実行**: `mainWithNewHotelCollection()`を実行

#### 日次自動実行の設定

1. `setupDailyIntegratedTrigger()`を実行
2. 以下のスケジュールで自動実行される：
   - 毎日9:00 - 新着ホテル収集
   - 毎日9:30 - 統合スクレイピング実行

#### 状況確認

- `showIntegratedStatus()`: 全体の状況を確認
- `checkNewHotelCollectionStatus()`: 新着ホテル収集状況のみ確認

### エラーURL再実行機能の場合

1. `retryErrorUrls()`を実行してエラーURLを再実行
2. 日次自動実行を設定する場合は`setupErrorRetryTrigger()`を実行
3. エラー再実行の状況は`showErrorRetryStatus()`で確認

### 実行時間制限対応機能

- **5分制限**: Google Apps Scriptの6分制限を考慮し、5分で処理を一時停止
- **自動継続**: 未処理URLがある場合、1分後に自動的に処理を再開
- **進捗保存**: 10件処理ごとに進捗を保存し、中断時も安全
- **状況確認**: `showExecutionStatus()`で処理状況を確認可能
- **手動停止**: `clearProgressTriggers()`で自動実行を停止可能

## 出力形式

スプレッドシートには以下の列でデータが出力されます：

| 列 | 内容 | 例 |
|---|---|---|
| A | 施設名 | 亀の井ホテル　九十九里 |
| B | 評価 | 3.94 |
| C | 評価件数 | 1,008 |
| D | 住所 | 〒289-2525千葉県旭市仁玉2280-1 |
| E | TEL | 0479-63-2161 |
| F | 総部屋数 | 81室 |
| G | URL | `https://travel.rakuten.co.jp/...` |
| H | エラー | （エラーがある場合のみ） |
| I | 取得日時 | 2024/12/20 10:30:00 |

## 利用上の注意事項

### 1. 利用規約の遵守

- 楽天トラベルの利用規約を確認し、遵守してください
- 過度なリクエストはサーバーに負荷をかけるため避けてください

### 2. レート制限

- デフォルトで2秒間隔でリクエストを送信
- 必要に応じて`REQUEST_DELAY`を調整

### 3. エラーハンドリング

- ネットワークエラーやページ構造変更に対応
- エラー発生時もスプレッドシートに記録

### 4. 権限設定

初回実行時に以下の権限が必要です：

- Google スプレッドシートへの読み書き
- 外部URLへのアクセス

## トラブルシューティング

### 1. スプレッドシートIDエラー

```bash
エラー: スプレッドシートIDが設定されていません
```

→ `config.gs`でSPREADSHEET_IDを正しく設定してください

### 2. 権限エラー

```bash
Authorization required
```

→ 実行時に表示される権限許可画面で「許可」をクリックしてください

### 3. データが取得できない

```bash
データ抽出エラー
```

→ 対象ページの構造が変更された可能性があります。HTMLの確認が必要です

### 4. 文字化けが発生する

```bash
新着ホテルセクションが見つかりません
HTMLで文字化け（�u�y�V...など）が発生
```

→ エンコーディングの問題です。新着ホテル収集機能では自動的にエンコーディングを検出・変換しますが、
手動で確認する場合は以下を実行してください：

```javascript
// エンコーディングテスト
function testEncoding() {
  const response = UrlFetchApp.fetch('https://travel.rakuten.co.jp/whatsnew.html');
  Logger.log('UTF-8:', response.getContentText('UTF-8').substring(0, 200));
  Logger.log('Shift_JIS:', response.getContentText('Shift_JIS').substring(0, 200));
}
```

### 5. 実行時間制限

Google Apps Scriptには6分の実行時間制限があります。大量のURLを処理する場合は以下の対応版を使用してください：

#### 実行時間制限対応版の特徴

- **自動分割処理**: 5分で処理を一時停止し、1分後に自動継続
- **進捗保存**: PropertiesServiceで処理状況を保存
- **安全停止**: 処理途中でも安全に中断・再開可能
- **状況監視**: `showExecutionStatus()`で進捗確認
- **手動制御**: `clearProgressTriggers()`で自動実行停止

#### 推奨する実行方法

```javascript
// 大量URL処理の場合（推奨）
mainWithProgressTracking();

// 状況確認
showExecutionStatus();

// 手動停止
clearProgressTriggers();
```

### その他のよくある問題と解決方法

#### 1. clasp 認証エラー

```bash
# 再認証
npx clasp logout
npx clasp login
```

#### 2. スプレッドシート権限エラー

```javascript
// config.gs でスプレッドシートIDを確認
// 共有設定でスクリプト実行権限を確認
```

#### 3. レート制限

```javascript
// REQUEST_DELAY を増加
const CONFIG = {
  REQUEST_DELAY: 5000  // 5秒に変更
};
```

## 参考資料

- [Google Apps Script 公式ドキュメント](https://developers.google.com/apps-script)
- [clasp (Command Line Apps Script Projects)](https://github.com/google/clasp)
- [Google Sheets API](https://developers.google.com/sheets/api)
