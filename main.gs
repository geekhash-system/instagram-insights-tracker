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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
 * READMEシートを挿入
 */
function insertReadmeSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 既存のREADMEシートがあれば削除
    const existingSheet = ss.getSheetByName(SHEET_NAMES.README);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // READMEシートを作成
    const readmeSheet = ss.insertSheet(SHEET_NAMES.README, 0); // 先頭に挿入

    const readmeContent = [
      ["📊 Instagram インサイト追跡ツール"],
      [""],
      ["このツールは、NERAとKARA子のInstagramインサイトを毎日19時に自動取得し、週次分析を行います。"],
      [""],
      ["【使い方】"],
      ["1. メニュー「インサイト追跡ツール」→「毎日19時の自動実行を開始」をクリック"],
      ["2. 毎日19時に自動でデータが取得されます"],
      ["3. 週次ダッシュボードで分析結果を確認できます"],
      [""],
      ["【PR投稿の設定】"],
      ["PR投稿は手動設定です:"],
      ["1. NERAまたはKARA子シートを開く"],
      ["2. PR投稿の行を見つける（企業タイアップ、案件投稿など）"],
      ["3. F列（PR列）のチェックボックスをクリックしてONにする"],
      ["4. 週次ダッシュボード更新を実行すると、PR投稿として集計されます"],
      [""],
      ["PR投稿の警告機能:"],
      ["- 過去10件のPR投稿の中央値を計算"],
      ["- 中央値×70%以下のIMP数の場合、週次ダッシュボードに警告が表示されます"],
      ["- 例: 中央値が10万IMPの場合、7万IMP以下で警告"],
      ["- 注意: PR投稿が0件の場合、週次ダッシュボードの「PR投稿」セクションは0と表示されます"],
      [""],
      ["【シート構成】"],
      ["- NERA: NERAのインサイトデータ（2025年12月以降）"],
      ["- KARA子: KARA子のインサイトデータ（2025年12月以降）"],
      ["- 週次_NERA_2026W02: NERAの週次ダッシュボード（週ごとに自動生成、13週間保持）"],
      ["- 週次_KARA子_2026W02: KARA子の週次ダッシュボード（週ごとに自動生成、13週間保持）"],
      [""],
      ["【IMP数について】"],
      ["このツールでは全ての投稿タイプ（リール・フィード）で「views」メトリクスを「IMP数」として使用しています。"],
      [""],
      ["投稿タイプ別の意味:"],
      ["- リール: 再生回数（動画が再生された回数）"],
      ["- フィード: 投稿が表示された回数（フィード上で見られた回数）"],
      [""],
      ["【注意事項】"],
      ["- データは直近90日分が取得されます"],
      ["- 履歴は90日分保持されます"],
      [""],
      ["【トラブルシューティング】"],
      ["- データが取得できない場合: Apps Script → 実行数 でログを確認してください"],
      ["- エラーが出る場合: アクセストークンが有効か確認してください"],
      [""],
      ["【詳細マニュアル】"],
      ["GitHub: https://github.com/geekhash-system/instagram-insights-tracker"],
      [""],
      ["最終更新: " + new Date().toLocaleString("ja-JP")]
    ];

    readmeSheet.getRange(1, 1, readmeContent.length, 1).setValues(readmeContent);
    readmeSheet.getRange("A1").setFontSize(16).setFontWeight("bold");
    readmeSheet.setColumnWidth(1, 800);

    SpreadsheetApp.getUi().alert("✅ READMEシートを挿入しました");
  } catch (e) {
    Logger.log(`エラー in insertReadmeSheet: ${e.toString()}`);
    SpreadsheetApp.getUi().alert("❌ エラーが発生しました: " + e.toString());
  }
}
