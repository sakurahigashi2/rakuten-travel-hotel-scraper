/**
 * 新着ホテル情報取得スクリプト
 * 楽天トラベルの新着情報ページから新しく登録されたホテル情報を取得
 */

/**
 * 新着ホテル情報を取得してマスタシートに追加する
 */
function collectNewHotels() {
  try {
    Logger.log('新着ホテル収集を開始します...');

    // config.gsから設定を取得
    const config = getConfig();

    // 適切なエンコーディングで新着ホテルページを取得
    const html = getHtmlWithCorrectEncoding(config.NEW_HOTELS_URL);
    Logger.log(`取得したHTMLサイズ: ${html.length} 文字`);

    // 複数の方法で新着ホテル情報を抽出
    let newHotels = [];

    // HTMLから新着ホテル情報を抽出（メイン処理）
    const hotelsFromHtml = extractNewHotelsFromHtml(html);
    newHotels = hotelsFromHtml;
    Logger.log(`HTMLから ${hotelsFromHtml.length} 件抽出`);

    // 重複除去（同じIDの重複のみ）
    const uniqueHotels = [];
    const seenIds = new Set();

    newHotels.forEach(hotel => {
      const hotelId = hotel.id.toString();

      if (!seenIds.has(hotelId)) {
        uniqueHotels.push(hotel);
        seenIds.add(hotelId);
      } else {
        Logger.log(`抽出時重複除外: [${hotelId}] ${hotel.name} - ${hotel.url}`);
      }
    });

    Logger.log(`重複除去後: ${uniqueHotels.length} 件のユニークなホテル`);

    // 同一バッチ内での重複チェック
    checkHotelIdDuplicates(uniqueHotels);

    if (uniqueHotels.length === 0) {
      Logger.log('新着ホテルが見つかりません');
      return;
    }

    // マスタシートから既存データを取得
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const masterSheet = spreadsheet.getSheetByName(config.MASTER_SHEET_NAME);

    if (!masterSheet) {
      throw new Error(`マスタシート "${config.MASTER_SHEET_NAME}" が見つかりません`);
    }

    const existingData = masterSheet.getDataRange().getValues();
    const existingIds = new Set();

    // 既存データから重複チェック用セットを作成（IDのみ）
    existingData.forEach(row => {
      if (row.length >= 1) {
        const id = row[0];      // A列: ID
        if (id) existingIds.add(id.toString().trim());
      }
    });

    // 既存データとの重複を除外（IDのみでチェック）
    const finalNewHotels = uniqueHotels.filter(hotel => {
      const isIdDuplicate = existingIds.has(hotel.id.toString());

      if (isIdDuplicate) {
        Logger.log(`重複により除外: [${hotel.id}] ${hotel.name} (ID重複)`);
        return false;
      }
      return true;
    });

    if (finalNewHotels.length === 0) {
      Logger.log('新規のホテルはありません（すべて既存データと重複）');
      return;
    }

    // マスタシートに新規ホテルを追加
    Logger.log(`${finalNewHotels.length} 件の新規ホテルをマスタシートに追加します`);

    finalNewHotels.forEach(hotel => {
      const newRow = [
        hotel.id,                                                     // A列: ID
        hotel.url,                                                    // B列: URL
        '',                                                           // C列: 施設名
        hotel.name,                                                   // D列: 評価
        hotel.discoveredAt ? hotel.discoveredAt.toLocaleString('ja-JP') : new Date().toLocaleString('ja-JP') // E列: 発見日時
      ];
      masterSheet.appendRow(newRow);
      Logger.log(`追加: [${hotel.id}] ${hotel.name} - ${hotel.url}`);
    });

    Logger.log('新着ホテル収集が完了しました');

  } catch (error) {
    Logger.log(`新着ホテル収集中にエラーが発生: ${error.toString()}`);
    throw error;
  }
}

/**
 * HTMLから新着ホテル情報を抽出
 */
function extractNewHotelsFromHtml(html) {
  Logger.log("extractNewHotelsFromHtml: 新着ホテル抽出開始");

  // 「最近登録いただいた宿泊施設の一覧」セクションを検索
  const targetSection = html.match(/最近登録いただいた宿泊施設の一覧[\s\S]*?(?=<\/div>|$)/);
  if (!targetSection) {
    Logger.log("目的のセクションが見つかりません");
    return [];
  }

  const sectionHtml = targetSection[0];
  Logger.log(`セクションHTML取得: ${sectionHtml.length}文字`);

  // ホテルリンクを抽出
  const hotelRegex = /<a[^>]*href=["']([^"']*\/hotel[^"']*\.html)["'][^>]*>([^<]*)<\/a>/gi;
  const hotels = [];
  let match;
  let skippedCount = 0;

  while ((match = hotelRegex.exec(sectionHtml)) !== null) {
    const url = match[1];
    let name = match[2].trim();

    // HTML エンティティをデコード
    name = name.replace(/&lt;/g, '<')
               .replace(/&gt;/g, '>')
               .replace(/&amp;/g, '&')
               .replace(/&quot;/g, '"')
               .replace(/&#39;/g, "'");

    // ホテル名を正規化
    name = normalizeHotelName(name);

    if (name && url) {
      // URLを正規化
      const normalizedUrl = normalizeHotelUrl(url);

      // ホテルIDを抽出
      const hotelId = extractHotelId(normalizedUrl);

      if (!hotelId) {
        Logger.log(`⚠️ ホテルIDが取得できないためスキップ: ${name} - ${normalizedUrl}`);
        skippedCount++;
        continue; // このホテルをスキップして次のマッチに進む
      }

      hotels.push({
        id: hotelId,
        name: name,
        url: normalizedUrl,
        discoveredAt: new Date()
      });
      Logger.log(`ホテル抽出: ${name} - ${normalizedUrl} (ID: ${hotelId})`);
    }
  }

  Logger.log(`extractNewHotelsFromHtml: ${hotels.length}件のホテルを抽出、${skippedCount}件をスキップ`);
  return hotels;
}

/**
 * ホテル名を正規化する
 */
function normalizeHotelName(name) {
  if (!name) return name;

  // 全角英数字を半角に変換
  let normalized = name.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(match) {
    return String.fromCharCode(match.charCodeAt(0) - 0xFEE0);
  });

  // 全角スペースを半角スペースに変換
  normalized = normalized.replace(/　/g, ' ');

  // 連続するスペースを1つに統一
  normalized = normalized.replace(/\s+/g, ' ');

  // 前後の空白を削除
  normalized = normalized.trim();

  return normalized;
}

/**
 * URLからホテルIDを抽出
 */
function extractHotelId(url) {
  const match = url.match(/\/HOTEL\/(\d+)\//);
  return match ? match[1] : '';
}

/**
 * ホテルURLを標準形式（_std.html）に正規化
 */
function normalizeHotelUrl(url) {
  try {
    let fullUrl = url;

    // 相対URLの場合、ベースURLを追加
    if (url.startsWith('/')) {
      fullUrl = 'https://travel.rakuten.co.jp' + url;
    } else if (url.startsWith('http://')) {
      // httpをhttpsに変換
      fullUrl = url.replace('http://', 'https://');
    }

    // 既に_std.htmlで終わっている場合はそのまま返す
    if (fullUrl.includes('_std.html')) {
      return fullUrl;
    }

    // ホテルIDを抽出
    const hotelId = extractHotelId(fullUrl);
    if (!hotelId) {
      Logger.log(`警告: ホテルIDを抽出できませんでした: ${fullUrl}`);
      return fullUrl;
    }

    // 標準形式のURLを構築
    const baseUrl = `https://travel.rakuten.co.jp/HOTEL/${hotelId}`;
    const standardUrl = `${baseUrl}/${hotelId}_std.html`;

    Logger.log(`URL正規化: ${url} → ${standardUrl}`);
    return standardUrl;

  } catch (error) {
    Logger.log(`URL正規化エラー: ${url} - ${error.toString()}`);
    return url; // エラーの場合は元のURLを返す
  }
}

/**
 * HTMLの構造を詳細に解析してリンクを探す
 */
/**
 * より詳細なホテル情報抽出（実際のリンク探索強化版）
 */
/**
 * ホテルIDの重複チェック関数
 */
function checkHotelIdDuplicates(hotels) {
  const idCounts = {};
  const duplicates = [];

  hotels.forEach(hotel => {
    const id = hotel.id.toString();
    if (idCounts[id]) {
      idCounts[id]++;
      if (idCounts[id] === 2) {
        duplicates.push(id);
      }
    } else {
      idCounts[id] = 1;
    }
  });

  if (duplicates.length > 0) {
    Logger.log(`⚠️ 同一バッチ内でのホテルID重複検出: ${duplicates.join(', ')}`);
    duplicates.forEach(duplicateId => {
      const duplicateHotels = hotels.filter(h => h.id.toString() === duplicateId);
      Logger.log(`ID ${duplicateId} の重複ホテル:`);
      duplicateHotels.forEach((h, i) => {
        Logger.log(`  ${i + 1}. ${h.name} - ${h.url}`);
      });
    });
  }

  return duplicates;
}

/**
 * 新着ホテル収集の状況を確認
 */
function checkNewHotelCollectionStatus() {
  try {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);

    // マスタシートの確認
    const masterSheet = spreadsheet.getSheetByName(config.MASTER_SHEET_NAME);
    if (masterSheet) {
      const totalHotels = masterSheet.getLastRow() - 1; // ヘッダー行を除く
      Logger.log(`マスタシート総ホテル数: ${totalHotels} 件`);
    }

    // 新着ホテルログシートの確認
    const logSheet = spreadsheet.getSheetByName(config.NEW_HOTELS_SHEET_NAME);
    if (logSheet) {
      const newHotelsCount = logSheet.getLastRow() - 1; // ヘッダー行を除く
      Logger.log(`発見済み新着ホテル数: ${newHotelsCount} 件`);

      // 最新の発見日時を取得
      if (newHotelsCount > 0) {
        const lastDiscoveryDate = logSheet.getRange(logSheet.getLastRow(), 1).getValue();
        Logger.log(`最新発見日時: ${lastDiscoveryDate}`);
      }
    }
  } catch (error) {
    Logger.log(`状況確認エラー: ${error.toString()}`);
  }
}
