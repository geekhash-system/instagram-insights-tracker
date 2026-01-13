#!/bin/bash
set -e

echo "📊 Instagram Insights Tracker - Spreadsheet Deployment Script"
echo ""

# スプレッドシートID
NERA_SPREADSHEET_ID="1R5mKOsOUhFbTPwOklG3Pa1G2iP1LN4Ah2XxKT7dlhNA"
KARAKO_SPREADSHEET_ID="1ka2WikKaEz2-lSd2eqk2Hcb29WhWuOOrKI6W_7FFrNE"

echo "⚠️  重要: 各スプレッドシートで以下の手順を実行してください:"
echo ""
echo "1. NERAスプレッドシートを開く:"
echo "   https://docs.google.com/spreadsheets/d/$NERA_SPREADSHEET_ID/edit"
echo ""
echo "2. メニュー「拡張機能」→「Apps Script」をクリック"
echo ""
echo "3. URLからScript IDをコピー（例: https://script.google.com/...d/{SCRIPT_ID}/edit）"
echo ""

read -p "NERAのScript IDを入力してください: " NERA_SCRIPT_ID

echo ""
echo "4. KARA子スプレッドシートを開く:"
echo "   https://docs.google.com/spreadsheets/d/$KARAKO_SPREADSHEET_ID/edit"
echo ""
echo "5. 同様にメニュー「拡張機能」→「Apps Script」→ Script IDをコピー"
echo ""

read -p "KARA子のScript IDを入力してください: " KARAKO_SCRIPT_ID

echo ""
echo "============================================"
echo "📝 入力内容確認"
echo "============================================"
echo "NERA Script ID: $NERA_SCRIPT_ID"
echo "KARA子 Script ID: $KARAKO_SCRIPT_ID"
echo ""

read -p "この内容でデプロイしますか？ (y/n): " confirm
if [ "$confirm" != "y" ]; then
    echo "キャンセルしました"
    exit 1
fi

# 一時ディレクトリ作成
WORK_DIR=$(mktemp -d)
echo ""
echo "🔧 作業ディレクトリ: $WORK_DIR"

# NERAデプロイ
echo ""
echo "============================================"
echo "📤 NERAスプレッドシートにデプロイ中..."
echo "============================================"

NERA_DIR="$WORK_DIR/nera"
mkdir -p "$NERA_DIR"
cd "$NERA_DIR"

# ファイルコピー
cp ~/dev/geekhash/instagram_insights_tracker/*.gs .
cp ~/dev/geekhash/instagram_insights_tracker/appsscript.json .

# プレースホルダー置換
sed -i '' "s/{{NERA_SPREADSHEET_ID}}/$NERA_SPREADSHEET_ID/g" config.gs
sed -i '' "s/{{KARAKO_SPREADSHEET_ID}}/$KARAKO_SPREADSHEET_ID/g" config.gs

# .clasp.json作成
cat > .clasp.json << EOF
{
  "scriptId": "$NERA_SCRIPT_ID",
  "rootDir": "."
}
EOF

# デプロイ
clasp push --force

echo "✅ NERAスプレッドシートへのデプロイ完了"

# KARA子デプロイ
echo ""
echo "============================================"
echo "📤 KARA子スプレッドシートにデプロイ中..."
echo "============================================"

KARAKO_DIR="$WORK_DIR/karako"
mkdir -p "$KARAKO_DIR"
cd "$KARAKO_DIR"

# ファイルコピー
cp ~/dev/geekhash/instagram_insights_tracker/*.gs .
cp ~/dev/geekhash/instagram_insights_tracker/appsscript.json .

# プレースホルダー置換
sed -i '' "s/{{NERA_SPREADSHEET_ID}}/$NERA_SPREADSHEET_ID/g" config.gs
sed -i '' "s/{{KARAKO_SPREADSHEET_ID}}/$KARAKO_SPREADSHEET_ID/g" config.gs

# .clasp.json作成
cat > .clasp.json << EOF
{
  "scriptId": "$KARAKO_SCRIPT_ID",
  "rootDir": "."
}
EOF

# デプロイ
clasp push --force

echo "✅ KARA子スプレッドシートへのデプロイ完了"

# クリーンアップ
cd ~
rm -rf "$WORK_DIR"

echo ""
echo "============================================"
echo "🎉 デプロイ完了！"
echo "============================================"
echo ""
echo "次の手順:"
echo "1. 各スプレッドシートをリロード（F5キー）"
echo "2. メニューバーに「📊 インサイト追跡ツール」が表示されることを確認"
echo "3. メニューから「READMEシートを挿入」を実行"
echo "4. メニューから「今すぐデータ取得」をテスト実行"
echo ""
echo "NERA: https://docs.google.com/spreadsheets/d/$NERA_SPREADSHEET_ID/edit"
echo "KARA子: https://docs.google.com/spreadsheets/d/$KARAKO_SPREADSHEET_ID/edit"
echo ""
