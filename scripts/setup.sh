#!/bin/bash

# Google Apps Script プロジェクトのセットアップスクリプト

echo "楽天トラベル ホテルスクレイパー - セットアップ開始"

# Node.js の依存関係をインストール
echo "依存関係をインストール中..."
npm install

# clasp の認証状態を確認
echo "clasp の認証状態を確認中..."
npx clasp login --status

if [ $? -ne 0 ]; then
    echo "clasp にログインしてください:"
    npx clasp login
fi

# 新しい Google Apps Script プロジェクトを作成するかどうか確認
echo ""
echo "新しい Google Apps Script プロジェクトを作成しますか？ (y/n)"
read -r response

if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
    echo "新しいプロジェクトを作成中..."
    npx clasp create --type standalone --title "楽天トラベル ホテルスクレイパー"

    echo ""
    echo "ファイルをアップロード中..."
    npx clasp push

    echo ""
    echo "プロジェクトを開きますか？ (y/n)"
    read -r open_response

    if [[ "$open_response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
        npx clasp open
    fi
else
    echo ""
    echo "既存のプロジェクトのスクリプトIDを入力してください:"
    read -r script_id

    # .clasp.json を更新
    echo "{
  \"scriptId\": \"$script_id\",
  \"rootDir\": \"./src\",
  \"filePushOrder\": [
    \"appsscript.json\",
    \"config.gs\",
    \"error_url_retry.gs\",
    \"hotel_scraper.gs\",
    \"hotel_scraper_advanced.gs\",
    \"hotel_scraper_master.gs\"
  ]
}" > .clasp.json

    echo "ファイルをアップロード中..."
    npx clasp push
fi

echo ""
echo "セットアップ完了！"
echo ""
echo "次のステップ:"
echo "1. src/config.gs でスプレッドシートIDを設定"
echo "2. 初回実行時は権限を許可"
echo ""
echo "詳細はREADME.mdを参照してください。"
