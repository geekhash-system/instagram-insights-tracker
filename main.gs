// ========================================
// main.gs - メイン処理・UI・トリガー管理
// ========================================

/**
 * スプレッドシート開く時に実行
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 インサイト追跡ツール")
    .addItem("今すぐデータ取得", "manualFetchAll")
    .addSeparator()
    .addItem("週次ダッシュボード更新", "manualUpdateDashboards")
    .addSeparator()
    .addItem("毎日19時の自動実行を開始", "setupDailyTrigger")
    .addItem("自動実行を停止", "removeTriggers")
    .addSeparator()
    .addItem("READMEシートを挿入", "insertReadmeSheet")
    .addToUi();
}

/**
 * 全アカウントのデータを取得（定期実行用）
 */
function fetchAllAccounts() {
  try {
    Logger.log("========================================");
    Logger.log("📅 定期実行開始: " + new Date().toLocaleString("ja-JP"));
    Logger.log("========================================");

    const { date, time } = getCurrentDateTime();

    // 各アカウントを処理
    ACCOUNTS.forEach(account => {
      Logger.log(`\n📱 アカウント: ${account.name}`);
      fetchAccountData(account, date, time);

      // API レート制限対策
      Utilities.sleep(1000);
    });

    Logger.log("\n========================================");
    Logger.log("✅ 全アカウントのデータ取得完了");
    Logger.log("========================================");

  } catch (e) {
    Logger.log(`❌ エラー in fetchAllAccounts: ${e.toString()}`);
    handleError("fetchAllAccounts", e, { severity: "HIGH" });
  }
}

/**
 * 1アカウントのデータを取得
 * @param {Object} account - アカウント設定
 * @param {string} date - 日付（YYYY-MM-DD）
 * @param {string} time - 時刻（HH:mm）
 */
function fetchAccountData(account, date, time) {
  try {
    const ss = SpreadsheetApp.openById(account.spreadsheetId);
    let sheet = ss.getSheetByName(account.sheetName);

    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet(account.sheetName);
      initializeAccountSheet(sheet);
    }

    // アクセストークン取得
    const accessToken = eval(account.tokenKey); // .env.gs から取得

    // メディア一覧取得
    const mediaList = fetchMediaList(account.businessId, accessToken, DATA_FETCH_CONFIG.MAX_DAYS_BACK);
    Logger.log(`📊 取得メディア数: ${mediaList.length}`);

    // 各メディアのインサイトを取得して更新
    mediaList.forEach((media, index) => {
      const insights = fetchMediaInsights(media.id, media.media_product_type, accessToken);
      updateMediaData(sheet, media, insights);

      // 進捗表示（10件ごと）
      if ((index + 1) % 10 === 0) {
        Logger.log(`  進捗: ${index + 1} / ${mediaList.length}`);
      }

      // API レート制限対策
      Utilities.sleep(DATA_FETCH_CONFIG.API_CALL_DELAY_MS);
    });

    // 日次履歴を記録
    addHistoryRecord(sheet, date, time);

    // 投稿日時で降順ソート（新しい投稿が上）
    sortSheetByDateDesc(sheet);

    // 週次ダッシュボード更新
    updateWeeklyDashboard(account.name);

    Logger.log(`✅ ${account.name} のデータ取得完了`);

  } catch (e) {
    Logger.log(`❌ エラー in fetchAccountData (${account.name}): ${e.toString()}`);
    handleError("fetchAccountData", e, { account: account.name });
  }
}

/**
 * NERA専用のデータ取得（並列実行用）
 */
function fetchNERA() {
  try {
    Logger.log("========================================");
    Logger.log("📅 NERA データ取得開始: " + new Date().toLocaleString("ja-JP"));
    Logger.log("========================================");

    const { date, time } = getCurrentDateTime();
    const account = ACCOUNTS.find(a => a.name === "NERA");

    if (account) {
      fetchAccountData(account, date, time);
    }

    Logger.log("✅ NERA データ取得完了");
  } catch (e) {
    Logger.log(`❌ エラー in fetchNERA: ${e.toString()}`);
    handleError("fetchNERA", e, { severity: "HIGH" });
  }
}

/**
 * KARA子専用のデータ取得（並列実行用）
 */
function fetchKARAKO() {
  try {
    Logger.log("========================================");
    Logger.log("📅 KARA子 データ取得開始: " + new Date().toLocaleString("ja-JP"));
    Logger.log("========================================");

    const { date, time } = getCurrentDateTime();
    const account = ACCOUNTS.find(a => a.name === "KARA子");

    if (account) {
      fetchAccountData(account, date, time);
    }

    Logger.log("✅ KARA子 データ取得完了");
  } catch (e) {
    Logger.log(`❌ エラー in fetchKARAKO: ${e.toString()}`);
    handleError("fetchKARAKO", e, { severity: "HIGH" });
  }
}

/**
 * 毎日19時の自動実行トリガーをセット（並列実行版）
 */
function setupDailyTrigger() {
  try {
    removeTriggers(); // 既存削除

    // NERA: 19:00に実行
    ScriptApp.newTrigger("fetchNERA")
      .timeBased()
      .atHour(DATA_FETCH_CONFIG.DAILY_TRIGGER_HOUR)
      .everyDays(1)
      .create();

    // KARA子: 19:05に実行（5分後）
    ScriptApp.newTrigger("fetchKARAKO")
      .timeBased()
      .atHour(DATA_FETCH_CONFIG.DAILY_TRIGGER_HOUR)
      .nearMinute(5)
      .everyDays(1)
      .create();

    SpreadsheetApp.getUi().alert("✅ 毎日19時の自動実行を開始しました\n・NERA: 19:00\n・KARA子: 19:05");
  } catch (e) {
    Logger.log(`エラー in setupDailyTrigger: ${e.toString()}`);
    SpreadsheetApp.getUi().alert("❌ エラー: " + e.toString());
  }
}

/**
 * トリガーを削除
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handlerName = trigger.getHandlerFunction();
    if (handlerName === "fetchAllAccounts" || handlerName === "fetchNERA" || handlerName === "fetchKARAKO") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 手動実行
 */
function manualFetchAll() {
  fetchAllAccounts();
  SpreadsheetApp.getUi().alert("✅ データ取得完了");
}

/**
 * 週次ダッシュボード手動更新
 */
function manualUpdateDashboards() {
  ACCOUNTS.forEach(account => {
    updateWeeklyDashboard(account.name);
  });
  SpreadsheetApp.getUi().alert("✅ 週次ダッシュボード更新完了");
}

/**
 * READMEシートを挿入（全アカウント）
 */
function insertReadmeSheet() {
  try {
    ACCOUNTS.forEach(account => {
      insertReadmeSheetForAccount(account);
    });
    SpreadsheetApp.getUi().alert("✅ 全アカウントのREADMEシートを挿入しました");
  } catch (e) {
    Logger.log(`エラー in insertReadmeSheet: ${e.toString()}`);
    SpreadsheetApp.getUi().alert("❌ エラーが発生しました: " + e.toString());
  }
}

/**
 * 指定アカウントのREADMEシートを挿入
 * @param {Object} account - アカウント設定
 */
function insertReadmeSheetForAccount(account) {
  try {
    const ss = SpreadsheetApp.openById(account.spreadsheetId);

    // 既存のREADMEシートがあれば削除
    const existingSheet = ss.getSheetByName(SHEET_NAMES.README);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // READMEシートを作成
    const readmeSheet = ss.insertSheet(SHEET_NAMES.README, 0); // 先頭に挿入

    const readmeContent = [
      ["📊 Instagram インサイト追跡ツール v2.0"],
      [""],
      [`このスプレッドシートは${account.name}専用です。毎日19時に自動でInstagramインサイトを取得し、週次分析を行います。`],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【📅 データ取得スケジュール】"],
      [`${account.name}のデータ取得時刻: 毎日19:${account.name === "NERA" ? "00" : "05"}`],
      ["取得対象期間: 直近30日分の投稿"],
      ["自動更新: 毎日実行（祝日含む）"],
      [""],
      ["【📊 シート構成】"],
      [`1. ${account.name}シート: インサイトデータ（直近30日分）`],
      ["   - A列: メディアID"],
      ["   - B列: 投稿日時"],
      ["   - C列: 投稿タイプ（REELS/FEED）"],
      ["   - D列: キャプション"],
      ["   - E列: パーマリンク（投稿URL）"],
      ["   - F列: PR（手動でチェック）⭐重要"],
      ["   - G列: IMP数（ビュー数）"],
      ["   - H列以降: リーチ、いいね、コメント、保存、シェア、エンゲージメント"],
      ["   - O列以降: 日次履歴（「1/13取得」形式で自動記録）"],
      [""],
      [`2. 週次_${account.name}_2026W03シート: 週次ダッシュボード`],
      ["   - 週ごとに自動生成されます（月曜日始まり）"],
      ["   - レポート生成日と計測対象期間が表示されます"],
      ["   - 全投稿、オーガニック投稿、PR投稿の統計"],
      ["   - トップ5/ワースト5の投稿リスト"],
      ["   - PR投稿の警告リスト"],
      ["   - 13週間後に自動削除されます"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【🎯 使い方】"],
      [""],
      ["■ 初回セットアップ（初めて使う場合）"],
      ["1. メニュー「📊 インサイト追跡ツール」→「毎日19時の自動実行を開始」"],
      ["2. メニュー「📊 インサイト追跡ツール」→「今すぐデータ取得」（テスト実行）"],
      [`3. ${account.name}シートにデータが表示されることを確認`],
      [""],
      ["■ 日常的な使い方"],
      [`1. ${account.name}シートで最新のインサイトを確認`],
      ["2. PR投稿のF列にチェックを入れる（企業タイアップ、案件投稿など）"],
      ["3. 週次ダッシュボードで週ごとの分析を確認"],
      [""],
      ["■ 週次ダッシュボードの見方"],
      ["- 「今週の投稿数」: 月曜〜日曜の投稿数"],
      ["- 「総IMP数」: 全投稿のビュー数合計"],
      ["- 「平均IMP」: 1投稿あたりの平均ビュー数"],
      ["- 「中央値IMP」: 投稿を並べた時の真ん中の値（外れ値の影響を受けにくい）"],
      ["- 「先週との差分」: 今週と先週の比較（↑↓で表示）"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【⭐ PR投稿の設定方法】"],
      [""],
      ["PR投稿は必ず手動で設定してください！"],
      [""],
      [`1. ${account.name}シートを開く`],
      ["2. PR投稿の行を見つける"],
      ["   例: 「#PR」「#タイアップ」「#案件」などのキャプション"],
      ["3. F列（PR列）のチェックボックスをクリック ✅"],
      ["4. メニュー「📊 インサイト追跡ツール」→「週次ダッシュボード更新」"],
      [""],
      ["■ PR投稿の警告機能"],
      ["過去10件のPR投稿の中央値を自動計算し、以下の条件で警告が表示されます:"],
      [""],
      ["⚠️ 警告条件: IMP数が中央値×70%以下"],
      [""],
      ["例: 中央値が10万IMPの場合"],
      ["  ✅ 10万IMP以上 → 正常"],
      ["  ⚠️ 7万IMP以下 → 警告表示（週次ダッシュボードに赤字で表示）"],
      [""],
      ["注意: PR投稿が10件未満の場合、中央値は計算されますが精度が低くなります"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【📈 IMP数（インプレッション数）について】"],
      [""],
      ["このツールでは「views」メトリクスを「IMP数」として使用しています。"],
      [""],
      ["投稿タイプ別の意味:"],
      ["  📹 リール: 動画の再生回数"],
      ["  📷 フィード: 投稿が表示された回数"],
      [""],
      ["なぜviewsを使うのか?"],
      ["Instagram Graph APIでは、リールとフィード両方で使える共通メトリクスが「views」です。"],
      ["「impressions」や「plays」は投稿タイプによって意味が異なるため、統一的な分析には適していません。"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【📋 重要な仕様・注意事項】"],
      [""],
      ["✅ データ保持期間: 直近30日分のみ"],
      ["   → 30日より古いデータは自動削除されます"],
      [""],
      ["✅ 週の定義: 月曜日始まり（月曜〜日曜が1週間）"],
      ["   → 例: 2026/1/6(月) 〜 2026/1/12(日) が第2週"],
      [""],
      ["✅ 週次ダッシュボードの保持期間: 13週間"],
      ["   → 約3ヶ月分の週次データが残ります"],
      [""],
      ["✅ 履歴データ: O列以降に日次のIMP数推移を記録"],
      ["   → 投稿のIMP数がどう変化したかを追跡できます"],
      [""],
      ["⚠️ このスプレッドシートは" + account.name + "専用です"],
      [`   → ${account.name === "NERA" ? "KARA子" : "NERA"}のデータは別のスプレッドシートで管理されています`],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【🔧 トラブルシューティング】"],
      [""],
      ["❌ データが取得できない"],
      ["  1. メニュー「拡張機能」→「Apps Script」を開く"],
      ["  2. 左メニュー「実行数」をクリック"],
      ["  3. エラーログを確認"],
      ["  4. よくあるエラー:"],
      ["     - アクセストークンの期限切れ → 開発者に連絡"],
      ["     - API制限超過 → 翌時間に自動復旧"],
      [""],
      ["❌ 週次ダッシュボードが更新されない"],
      [`  1. ${account.name}シートにデータがあるか確認`],
      ["  2. メニューから「週次ダッシュボード更新」を手動実行"],
      ["  3. エラーログを確認"],
      [""],
      ["❌ PR投稿の警告が表示されない"],
      ["  1. F列（PR列）にチェックが入っているか確認"],
      ["  2. PR投稿が10件以上あるか確認"],
      ["  3. 週次ダッシュボードを再生成"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【🔗 参考リンク】"],
      [""],
      ["📚 詳細マニュアル（GitHub）:"],
      ["https://github.com/geekhash-system/instagram-insights-tracker"],
      [""],
      ["💡 質問・不具合報告:"],
      ["https://github.com/geekhash-system/instagram-insights-tracker/issues"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["バージョン: v2.0"],
      ["最終更新: " + new Date().toLocaleString("ja-JP", {timeZone: "Asia/Tokyo"})],
      [`対象アカウント: ${account.name}`]
    ];

    readmeSheet.getRange(1, 1, readmeContent.length, 1).setValues(readmeContent);
    readmeSheet.getRange("A1").setFontSize(16).setFontWeight("bold");
    readmeSheet.setColumnWidth(1, 800);

    Logger.log(`✅ ${account.name} のREADMEシートを挿入しました`);
  } catch (e) {
    Logger.log(`エラー in insertReadmeSheetForAccount (${account.name}): ${e.toString()}`);
  }
}
