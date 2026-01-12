# 📊 Instagram Insights Tracker

NERAとKARA子のInstagramインサイトを毎日19時に自動取得し、週次分析を行うGoogle Apps Scriptツールです。

---

## 主な機能

- **毎日19時に自動実行**: Instagram Graph APIからインサイトを自動取得
- **直近90日分のデータ管理**: 投稿データとインサイトを自動更新
- **週次ダッシュボード**: 今週/先週の比較分析（投稿数、IMP数、平均、中央値）
- **PR投稿管理**: 手動設定 + 過去10投稿の中央値ベース警告機能
- **日次履歴記録**: 毎日のIMP数推移を自動記録

---

## セットアップ

### 1. スプレッドシート作成

1. 新しいGoogleスプレッドシートを作成
2. スプレッドシートIDをコピー
   ```
   https://docs.google.com/spreadsheets/d/【ここがID】/edit
   ```

### 2. Google Apps Scriptプロジェクト作成

#### 方法A: Webエディタを使う（推奨）

1. スプレッドシートを開く
2. メニューバーから「拡張機能」→「Apps Script」をクリック
3. 以下のファイルを作成してコピー＆ペースト:
   - `config.gs` - `SPREADSHEET_ID`を実際のIDに置き換える
   - `.env.gs` - アクセストークンを設定
   - `utils.gs`
   - `instagramAPI.gs`
   - `sheetManager.gs`
   - `analytics.gs`
   - `main.gs`
   - `appsscript.json`

#### 方法B: claspを使う（開発者向け）

```bash
# claspのインストール
npm install -g @google/clasp

# Googleアカウントでログイン
clasp login

# プロジェクト作成
clasp create --type sheets --title "Instagram Insights Tracker"

# .clasp.jsonのscriptIdを確認
cat .clasp.json

# コードをプッシュ
clasp push

# Apps Scriptエディタを開く
clasp open
```

### 3. 初回実行と権限の承認

1. Apps Scriptエディタで`onOpen`関数を実行
2. 「権限を確認」→ Googleアカウントを選択
3. 「詳細」→「（プロジェクト名）に移動」→「許可」
4. スプレッドシートをリロードして、カスタムメニューが表示されることを確認

---

## 使い方

### 初回セットアップ

1. スプレッドシートを開く
2. メニューから「📊 インサイト追跡ツール」→「READMEシートを挿入」
3. メニューから「📊 インサイト追跡ツール」→「毎日19時の自動実行を開始」
4. メニューから「📊 インサイト追跡ツール」→「今すぐデータ取得」（初回テスト）

### データ確認

- **NERA** シート: NERAのインサイトデータ（2025年12月以降）
- **KARA子** シート: KARA子のインサイトデータ（2025年12月以降）
- **週次_NERA_2026W02** シート: NERAの週次ダッシュボード（週ごとに自動生成、13週間保持）
- **週次_KARA子_2026W02** シート: KARA子の週次ダッシュボード（週ごとに自動生成、13週間保持）

### PR投稿の設定

PR投稿は**手動設定**です。以下の手順で設定してください:

1. **NERAまたはKARA子シートを開く**
2. **PR投稿の行を見つける**（企業タイアップ、案件投稿など）
3. **F列（PR列）のチェックボックスをクリック**してONにする
4. 次回のデータ取得時、または「週次ダッシュボード更新」を手動実行すると、PR投稿として集計されます

**PR投稿の警告機能:**
- 過去10件のPR投稿の中央値を計算
- 中央値×70%以下のIMP数の場合、週次ダッシュボードに警告が表示されます
- 例: 中央値が10万IMPの場合、7万IMP以下で警告

**注意:** PR投稿が0件の場合、週次ダッシュボードの「PR投稿」セクションは0と表示されます

---

## スプレッドシート構造

### データシート（NERA / KARA子）

**1行目**: シート説明（例: このシートは2026年1月以降のInstagramインサイトデータを記録しています。毎日19時に自動更新されます。）

**2行目以降**:

| 列 | 項目 | 説明 |
|----|------|------|
| A | メディアID | Instagram投稿ID |
| B | 投稿日時 | 投稿された日時 |
| C | 投稿タイプ | REELS / FEED / STORY |
| D | キャプション | 投稿のキャプション |
| E | パーマリンク | Instagram投稿URL |
| F | PR | PR投稿かどうか（手動設定） |
| G | IMP数 | ビュー数（viewsメトリクス） |
| H | リーチ数 | リーチ数 |
| I | いいね数 | いいね数 |
| J | コメント数 | コメント数 |
| K | 保存数 | 保存数 |
| L | シェア数 | シェア数 |
| M | エンゲージメント数 | 総エンゲージメント数 |
| N | 最終更新日時 | データ最終更新日時 |
| O列以降 | 履歴 | 日次履歴（「12/25取得」形式） |

### 週次ダッシュボードシート（週次_NERA_2026W02 / 週次_KARA子_2026W02）

シート名は `週次_{アカウント名}_{年}W{週番号}` の形式で自動生成されます（例: 週次_NERA_2026W02）
古いシートは13週間後に自動削除されます

- **全投稿の統計**: 今週/先週の投稿数、総IMP数、平均IMP、中央値IMP
- **オーガニック投稿の統計**: オーガニック投稿のみの分析
- **PR投稿の統計**: PR投稿のみの分析
- **オーガニック投稿 トップ5**: IMP数が高い投稿
- **オーガニック投稿 ワースト5**: IMP数が低い投稿
- **PR投稿警告リスト**: 過去10投稿の中央値×70%以下の投稿

---

## Instagram Graph API仕様

### 使用メトリクス

hinome-backend実装に基づき、以下のメトリクスを使用:

**リール:**
```
comments, likes, views, reach, saved, shares, total_interactions
```

**フィード:**
```
saved, reach, total_interactions, views
```

**ストーリー:**
```
exits, views, reach, taps_forward, taps_back
```

### IMP数について

このツールでは、全ての投稿タイプ（リール・フィード・ストーリー）で **`views`** メトリクスを「IMP数」として使用しています。

**投稿タイプ別の意味:**
- **リール (REELS)**: 再生回数（動画が再生された回数）
- **フィード (FEED)**: 投稿が表示された回数（フィード上で見られた回数）
- **ストーリー (STORY)**: ストーリーが視聴された回数

**注意点:**
- Instagram Graph APIでは `views` が標準メトリクス（`plays`や`impressions`ではない）
- フィード投稿の場合、厳密には `reach`（リーチ数）の方が適切な指標ですが、hinome-backend実装に倣い、統一性のため全て `views` を使用しています
- APIバージョン: **v18.0**（安定版）

### API制限

- **レート制限**: 1時間に200リクエストまで
- **データ保持期間**: 90日（それ以前のデータは取得不可）
- **アクセストークン**: 期限切れしない長期トークン（ユーザー提供）

---

## ファイル構成

```
instagram_insights_tracker/
├── config.gs          # 設定ファイル（API設定、アカウント設定）
├── .env.gs            # 環境変数（アクセストークン）※.gitignore登録
├── main.gs            # メイン処理・UI・トリガー管理
├── instagramAPI.gs    # Instagram Graph API連携
├── sheetManager.gs    # スプレッドシート操作・データ管理
├── analytics.gs       # 週次ダッシュボード計算ロジック
├── utils.gs           # ユーティリティ関数
├── appsscript.json    # Google Apps Script設定
├── .clasp.json        # clasp設定（scriptIdは自分で設定）
├── .claspignore       # clasp除外設定
├── .gitignore         # Git除外設定
└── README.md          # このファイル
```

---

## トラブルシューティング

### データが取得できない

1. Apps Script → 実行数 でログを確認
2. アクセストークンが正しく設定されているか確認（`.env.gs`）
3. `SPREADSHEET_ID`が正しく設定されているか確認（`config.gs`）

### 週次ダッシュボードが更新されない

1. データシートに投稿データがあるか確認
2. メニューから手動で「週次ダッシュボード更新」を実行してみる
3. ログを確認してエラーがないかチェック

### PR投稿の警告が表示されない

1. PR列（F列）にチェックが入っているか確認
2. PR投稿が10件以上あるか確認（中央値計算のため）
3. IMP数が中央値×70%以下になっているか確認

---

## 拡張性

### アカウント追加

`config.gs`の`ACCOUNTS`配列に追加するだけ:

```javascript
{
  name: "新しいアカウント",
  sheetName: "新しいアカウント_25/12~",
  dashboardSheet: "週次_新しいアカウント",
  businessId: "Instagram Business Account ID",
  tokenKey: "NEW_ACCOUNT_ACCESS_TOKEN"
}
```

`.env.gs`にトークンを追加:

```javascript
const NEW_ACCOUNT_ACCESS_TOKEN = "アクセストークン";
```

---

## GitHub Actions による自動デプロイ

このプロジェクトはGitHub Actionsを使用して、mainブランチへのpush時に自動的にGoogle Apps Scriptへデプロイされます。

### セットアップ手順

1. **GitHub Secretsを設定**

   リポジトリの Settings → Secrets and variables → Actions → New repository secret から以下を追加:

   - `CLASPRC_JSON`: `~/.clasprc.json`の内容
     ```bash
     cat ~/.clasprc.json
     ```

   - `SCRIPT_ID_MAIN`: Apps ScriptのScript ID
     ```
     1mdKYeHhd6xFdFKz9Uef7JrWjkMay3FfTpttu9gomzGnh8zmo1SUztQeu
     ```

   - `SPREADSHEET_ID_MAIN`: スプレッドシートID
     ```
     1mi10KVyRkf8_svopr2Hlh4a5LJxW4QO9fSTx7EjzBGQ
     ```

   - `ENV_GS`: `.env.gs`ファイルの内容
     ```bash
     cat .env.gs
     ```

2. **デプロイ**

   mainブランチにpushすると、GitHub Actionsが自動的に:
   - プレースホルダ (`{{SCRIPT_ID}}`, `{{SPREADSHEET_ID}}`) を実際の値に置換
   - `.env.gs`をSecretから生成
   - `clasp push`でApps Scriptへデプロイ

3. **デプロイ確認**

   GitHub の Actions タブでワークフローの実行状況を確認できます。

---

## セキュリティ

⚠️ **重要**: `.env.gs`は絶対にGitHubにプッシュしないでください

- `.gitignore`に登録済み
- アクセストークンは機密情報として扱ってください
- スプレッドシートの共有は最小限にしてください
- GitHub Secretsにトークンを保存し、GitHub Actionsでデプロイします

---

## ライセンス

MIT License - 個人利用・商用利用ともに自由に使用できます。

---

**Built with ❤️ using Google Apps Script**

参考プロジェクト:
- [GAS_instagram_reel_viewcount_tracker](../buzz/GAS_instagram_reel_viewcount_tracker/)
