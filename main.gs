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
    .addItem("週次レポート生成", "manualGenerateWeeklyReports")
    .addSeparator()
    .addItem("日時設定の変更", "configureSchedule")
    .addSeparator()
    .addItem("READMEシートを更新", "insertReadmeSheet")
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

    // アカウントインサイト履歴を記録
    const accountInsightsSheet = getOrCreateAccountInsightsSheet(ss);
    if (accountInsightsSheet) {
      recordAccountInsights(accountInsightsSheet, account, accessToken, date);
    }

    // 投稿日時で降順ソート（新しい投稿が上）
    sortSheetByDateDesc(sheet);

    // 週次ダッシュボード更新は水曜日00:00のトリガーで実行
    // （毎日のデータ取得時には実行しない）

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
 * 日時設定ダイアログを表示
 */
function configureSchedule() {
  try {
    // 現在の設定を取得
    const properties = PropertiesService.getScriptProperties();
    const currentHour = properties.getProperty('DAILY_TRIGGER_HOUR') || '19';
    const currentWeekDay = properties.getProperty('WEEKLY_TRIGGER_DAY') || 'WEDNESDAY';
    const currentWeeklyHour = properties.getProperty('WEEKLY_TRIGGER_HOUR') || '0';

    // 現在のトリガー状況を確認
    const triggers = ScriptApp.getProjectTriggers();
    const hasTriggers = triggers.some(t => {
      const name = t.getHandlerFunction();
      return name === "fetchNERA" || name === "fetchKARAKO" || name === "generateWeeklyReportsAll";
    });

    const weekDayNames = {
      'MONDAY': '月曜日',
      'TUESDAY': '火曜日',
      'WEDNESDAY': '水曜日',
      'THURSDAY': '木曜日',
      'FRIDAY': '金曜日',
      'SATURDAY': '土曜日',
      'SUNDAY': '日曜日'
    };

    // 簡易ダイアログで設定を表示・変更
    const ui = SpreadsheetApp.getUi();
    const statusText = hasTriggers ? '自動実行: 有効' : '自動実行: 停止中';

    const response = ui.alert(
      '⚙️ 日時設定',
      `${statusText}\n\n` +
      `毎日のデータ取得: ${currentHour}時\n` +
      `週次レポート生成: ${weekDayNames[currentWeekDay]} ${currentWeeklyHour}時\n\n` +
      '設定を変更しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      showScheduleConfigDialog();
    }
  } catch (e) {
    Logger.log(`エラー in configureSchedule: ${e.toString()}`);
    SpreadsheetApp.getUi().alert("❌ エラー: " + e.toString());
  }
}

/**
 * 詳細設定ダイアログを表示
 */
function showScheduleConfigDialog() {
  const ui = SpreadsheetApp.getUi();

  // データ取得時刻の入力
  const hourResponse = ui.prompt(
    'データ取得時刻の設定',
    '毎日のデータ取得時刻を入力してください（0-23）:',
    ui.ButtonSet.OK_CANCEL
  );

  if (hourResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const dailyHour = parseInt(hourResponse.getResponseText());
  if (isNaN(dailyHour) || dailyHour < 0 || dailyHour > 23) {
    ui.alert('❌ 無効な時刻です。0から23の間で入力してください。');
    return;
  }

  // 週次レポート曜日の選択
  const dayResponse = ui.prompt(
    '週次レポート生成曜日の設定',
    '曜日を入力してください\n（月/火/水/木/金/土/日）:',
    ui.ButtonSet.OK_CANCEL
  );

  if (dayResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const dayMap = {
    '月': 'MONDAY',
    '火': 'TUESDAY',
    '水': 'WEDNESDAY',
    '木': 'THURSDAY',
    '金': 'FRIDAY',
    '土': 'SATURDAY',
    '日': 'SUNDAY'
  };

  const weekDay = dayMap[dayResponse.getResponseText()];
  if (!weekDay) {
    ui.alert('❌ 無効な曜日です。月/火/水/木/金/土/日 のいずれかを入力してください。');
    return;
  }

  // 週次レポート時刻の入力
  const weeklyHourResponse = ui.prompt(
    '週次レポート生成時刻の設定',
    '週次レポート生成時刻を入力してください（0-23）:',
    ui.ButtonSet.OK_CANCEL
  );

  if (weeklyHourResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const weeklyHour = parseInt(weeklyHourResponse.getResponseText());
  if (isNaN(weeklyHour) || weeklyHour < 0 || weeklyHour > 23) {
    ui.alert('❌ 無効な時刻です。0から23の間で入力してください。');
    return;
  }

  // 設定を適用
  const result = applyScheduleSettings(dailyHour.toString(), weekDay, weeklyHour.toString());
  ui.alert(result.message);
}

/**
 * 設定を適用してトリガーを更新
 */
function applyScheduleSettings(dailyHour, weekDay, weeklyHour) {
  try {
    // 設定を保存
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('DAILY_TRIGGER_HOUR', dailyHour);
    properties.setProperty('WEEKLY_TRIGGER_DAY', weekDay);
    properties.setProperty('WEEKLY_TRIGGER_HOUR', weeklyHour);

    // トリガーを再作成
    removeTriggers();

    // 現在のスプレッドシートがNERAかKARA子かを判定
    const currentSpreadsheetId = SpreadsheetApp.getActive().getId();
    const account = ACCOUNTS.find(a => a.spreadsheetId === currentSpreadsheetId);

    if (!account) {
      return {
        success: false,
        message: '❌ エラー: このスプレッドシートに対応するアカウントが見つかりません'
      };
    }

    // 毎日のデータ取得トリガー（このアカウント用）
    const fetchFunction = account.name === "NERA" ? "fetchNERA" : "fetchKARAKO";
    ScriptApp.newTrigger(fetchFunction)
      .timeBased()
      .atHour(parseInt(dailyHour))
      .everyDays(1)
      .create();

    // 週次レポート生成トリガー
    const weekDayEnum = ScriptApp.WeekDay[weekDay];
    ScriptApp.newTrigger("generateWeeklyReportsAll")
      .timeBased()
      .onWeekDay(weekDayEnum)
      .atHour(parseInt(weeklyHour))
      .create();

    const weekDayNames = {
      'MONDAY': '月曜日',
      'TUESDAY': '火曜日',
      'WEDNESDAY': '水曜日',
      'THURSDAY': '木曜日',
      'FRIDAY': '金曜日',
      'SATURDAY': '土曜日',
      'SUNDAY': '日曜日'
    };

    Logger.log(`✅ 設定を適用しました: 毎日${dailyHour}時, ${weekDay} ${weeklyHour}時`);
    return {
      success: true,
      message: `✅ 設定を適用しました\n\n・毎日のデータ取得: ${dailyHour}時\n・週次レポート生成: ${weekDayNames[weekDay]} ${weeklyHour}時`
    };
  } catch (e) {
    Logger.log(`❌ エラー in applyScheduleSettings: ${e.toString()}`);
    return {
      success: false,
      message: `❌ エラー: ${e.toString()}`
    };
  }
}

/**
 * 自動実行トリガーをセット（毎日のデータ取得 + 毎週水曜日の週次レポート生成）
 */
function setupTriggers() {
  try {
    removeTriggers(); // 既存削除

    // NERA: 19:00に実行
    ScriptApp.newTrigger("fetchNERA")
      .timeBased()
      .atHour(19)
      .everyDays(1)
      .create();

    // KARA子: 19:05に実行（5分後）
    ScriptApp.newTrigger("fetchKARAKO")
      .timeBased()
      .atHour(19)
      .nearMinute(5)
      .everyDays(1)
      .create();

    // 毎週水曜日00:00の週次レポート生成トリガー
    ScriptApp.newTrigger("generateWeeklyReportsAll")
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .atHour(0)
      .create();

    SpreadsheetApp.getUi().alert(
      "✅ 自動実行を開始しました\n" +
      "・毎日19時: データ取得 (NERA 19:00, KARA子 19:05)\n" +
      "・毎週水曜日00:00: 週次レポート生成"
    );
  } catch (e) {
    Logger.log(`エラー in setupTriggers: ${e.toString()}`);
    SpreadsheetApp.getUi().alert("❌ エラー: " + e.toString());
  }
}

/**
 * 旧関数名との互換性のため残す
 */
function setupDailyTrigger() {
  setupTriggers();
}

/**
 * トリガーを削除
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handlerName = trigger.getHandlerFunction();
    if (handlerName === "fetchAllAccounts" ||
        handlerName === "fetchNERA" ||
        handlerName === "fetchKARAKO" ||
        handlerName === "generateWeeklyReportsAll") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 手動実行（現在のスプレッドシートのアカウントのみ）
 */
function manualFetchAll() {
  try {
    const currentSpreadsheetId = SpreadsheetApp.getActive().getId();
    const account = ACCOUNTS.find(a => a.spreadsheetId === currentSpreadsheetId);

    if (!account) {
      SpreadsheetApp.getUi().alert("❌ エラー: このスプレッドシートに対応するアカウントが見つかりません");
      return;
    }

    Logger.log(`========================================`);
    Logger.log(`📅 ${account.name} データ取得開始: ${new Date().toLocaleString("ja-JP")}`);
    Logger.log(`========================================`);

    const { date, time } = getCurrentDateTime();
    fetchAccountData(account, date, time);

    Logger.log(`✅ ${account.name} データ取得完了`);
    SpreadsheetApp.getUi().alert(`✅ ${account.name}のデータ取得完了`);
  } catch (e) {
    Logger.log(`❌ エラー in manualFetchAll: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`❌ エラー: ${e.toString()}`);
  }
}


/**
 * READMEシートを挿入（現在のスプレッドシートのアカウントのみ）
 */
function insertReadmeSheet() {
  try {
    const currentSpreadsheetId = SpreadsheetApp.getActive().getId();
    const account = ACCOUNTS.find(a => a.spreadsheetId === currentSpreadsheetId);

    if (!account) {
      SpreadsheetApp.getUi().alert("❌ エラー: このスプレッドシートに対応するアカウントが見つかりません");
      return;
    }

    insertReadmeSheetForAccount(account);
    SpreadsheetApp.getUi().alert(`✅ ${account.name}のREADMEシートを挿入しました`);
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
      ["📊 Instagram インサイト追跡ツール"],
      [""],
      [`このスプレッドシートは${account.name}専用です。毎日19時に自動でInstagramインサイトを取得します。`],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【🔧 メニュー機能】"],
      [""],
      ["メニューバー「📊 インサイト追跡ツール」から以下の機能を実行できます:"],
      [""],
      ["■ 今すぐデータ取得"],
      ["  → 投稿データとアカウントインサイトを手動で即座に取得します"],
      ["  → 実行すると、投稿データ・フォロワー数・エンゲージメント指標が更新されます"],
      [""],
      ["■ 週次レポート生成"],
      ["  → 週次レポートシート（週次レポート_mmdd形式）を手動で生成します"],
      ["  → 2週間前（月曜〜日曜）のデータと3週間前の比較レポートを作成"],
      ["  → トップ5/ワースト5投稿、PR警告を表示"],
      [""],
      ["■ 日時設定の変更"],
      ["  → 自動実行のスケジュールを設定します（ダイアログ形式）"],
      ["  → 設定内容:"],
      ["    - 毎日のデータ取得時刻（0〜23時）"],
      ["    - 週次レポート生成の曜日（月〜日）"],
      ["    - 週次レポート生成の時刻（0〜23時）"],
      ["  → 初回セットアップ時は必ず実行してください"],
      [""],
      ["■ READMEシートを更新"],
      ["  → このREADMEシートを最新の内容に再生成します"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【📊 シート構成】"],
      [""],
      [`1. ${account.name}シート: 投稿インサイトデータ`],
      ["   - C列: 曜日（月火水木金土日）"],
      ["   - G列: PR（手動でチェック）⭐重要"],
      ["   - H列: IMP数（ビュー数）"],
      ["   - P列以降: 日次履歴（永久保存）"],
      [""],
      ["2. アカウントインサイト履歴シート: アカウント全体の指標"],
      ["   - B列: フォロワー数"],
      ["   - C列: フォロワー増減数（前日比、増加は緑、減少は赤）"],
      ["   - D列以降: リーチ、エンゲージメント指標"],
      [""],
      ["3. 週次レポート_mmddシート: 週次ダッシュボード"],
      ["   - 2週間前と3週間前のデータ比較（投稿から7日目のデータを使用）"],
      ["   - トップ5/ワースト5の投稿リスト"],
      ["   - PR投稿の警告リスト"],
      ["   - 毎週水曜日00:00に自動生成、永久保存"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【🎯 初回セットアップ】"],
      [""],
      ["1. メニュー「📊 インサイト追跡ツール」→「日時設定の変更」"],
      ["   → データ取得時刻、週次レポート生成の曜日・時刻を設定"],
      ["   → デフォルト: 毎日19時、毎週水曜日00:00"],
      [""],
      ["2. メニュー「📊 インサイト追跡ツール」→「今すぐデータ取得」でテスト実行"],
      ["   → データが正しく取得されることを確認"],
      [""],
      ["3. メニュー「📊 インサイト追跡ツール」→「週次レポート生成」でテスト実行"],
      ["   → 週次レポートシートが生成されることを確認"],
      [""],
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【⭐ PR投稿の設定方法】"],
      [""],
      [`1. ${account.name}シートでPR投稿を見つける`],
      ["   例: 「#PR」「#タイアップ」「#案件」などのキャプション"],
      ["2. G列（PR列）のチェックボックスをクリック ✅"],
      ["3. メニュー「📊 インサイト追跡ツール」→「週次レポート生成」"],
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
      ["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"],
      [""],
      ["【📈 週次ダッシュボードの見方】"],
      [""],
      ["■ 基本指標の定義"],
      ["- 「今週の投稿数」: 月曜〜日曜の投稿数"],
      ["- 「総IMP数」: 全投稿のビュー数合計（投稿から7日目のデータを使用）"],
      ["- 「平均IMP」: 1投稿あたりの平均ビュー数（投稿から7日目のデータを使用）"],
      ["- 「中央値IMP」: 投稿を並べた時の真ん中の値（外れ値の影響を受けにくい）"],
      ["- 「先週との差分」: 2週間前と3週間前の比較（↑↓で表示）"],
      [""],
      ["■ 週の定義"],
      ["月曜日始まり（月曜〜日曜が1週間）"],
      ["例: 2026/1/6(月) 〜 2026/1/12(日) が第2週"],
      [""],
      ["■ レポート対象期間"],
      ["- 計測対象: 2週間前の月曜〜日曜"],
      ["- 比較対象: 3週間前の月曜〜日曜"],
      ["- 使用データ: 各投稿の「投稿日+7日目」の履歴データ（P〜V列）"],
      ["- 理由: 投稿から7日後のデータが完全に揃っている週を分析するため"],
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
      ["【📋 データ仕様】"],
      [""],
      ["⏰ 自動実行スケジュール（デフォルト）:"],
      ["   - データ取得: 毎日19:00（祝日含む）"],
      ["   - 週次レポート生成: 毎週水曜日 00:00"],
      ["   ※「日時設定の変更」メニューから変更可能"],
      [""],
      ["📅 取得期間: 直近30日分の投稿（既存データは削除されません）"],
      ["💾 データ保存: 全データ永久保存（自動削除なし）"],
      ["📊 週次レポート: 2週間前のデータを分析（投稿+7日目の完全なデータを使用）"],
      [""],
      [`⚠️ このスプレッドシートは${account.name}専用です`],
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
      ["❌ 週次レポートが生成されない"],
      [`  1. ${account.name}シートにデータがあるか確認`],
      ["  2. メニューから「週次レポート生成」を手動実行"],
      ["  3. エラーログを確認"],
      ["  4. 2週間前のデータが存在するか確認"],
      [""],
      ["❌ PR投稿の警告が表示されない"],
      ["  1. G列（PR列）にチェックが入っているか確認"],
      ["  2. PR投稿が10件以上あるか確認"],
      ["  3. 週次レポートを再生成"],
      [""],
      ["❌ 自動実行のスケジュールを変更したい"],
      ["  1. メニュー「日時設定の変更」を実行"],
      ["  2. 希望する時刻・曜日を入力"],
      ["  3. 確認画面でOKをクリック"],
      ["  4. Apps Scriptエディタ > トリガー画面で設定を確認"],
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

/**
 * アカウントインサイトシートを取得または作成
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @return {Sheet} アカウントインサイトシート
 */
function getOrCreateAccountInsightsSheet(ss) {
  try {
    let sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNT_INSIGHTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.ACCOUNT_INSIGHTS);
      initializeAccountInsightsSheet(sheet);
      Logger.log(`📊 アカウントインサイトシートを作成しました`);
    }
    return sheet;
  } catch (e) {
    Logger.log(`エラー in getOrCreateAccountInsightsSheet: ${e.toString()}`);
    return null;
  }
}

/**
 * アカウントインサイトデータを記録
 * @param {Sheet} sheet - アカウントインサイトシート
 * @param {Object} account - アカウント設定
 * @param {string} accessToken - アクセストークン
 * @param {string} date - 日付（YYYY-MM-DD）
 */
function recordAccountInsights(sheet, account, accessToken, date) {
  try {
    Logger.log(`📊 アカウントインサイト取得開始: ${account.name}`);

    // Fetch follower count
    const accountInfo = fetchAccountInfo(account.businessId, accessToken);

    // Fetch account insights for yesterday (more stable data)
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    const since = Math.floor(yesterday.getTime() / 1000);
    const until = Math.floor((yesterday.getTime() + 86400000) / 1000); // +24 hours

    const insights = fetchAccountInsights(account.businessId, accessToken, since, until);

    // Record data
    addAccountInsightsRecord(sheet, date, accountInfo, insights);

    Logger.log(`✅ アカウントインサイト記録完了: ${account.name}`);
    Utilities.sleep(DATA_FETCH_CONFIG.API_CALL_DELAY_MS);
  } catch (e) {
    Logger.log(`❌ エラー in recordAccountInsights (${account.name}): ${e.toString()}`);
  }
}

/**
 * 全アカウントの週次レポートを生成（毎週水曜00:00実行）
 */
function generateWeeklyReportsAll() {
  try {
    Logger.log("========================================");
    Logger.log("📊 週次レポート生成開始: " + new Date().toLocaleString("ja-JP"));
    Logger.log("========================================");

    ACCOUNTS.forEach(account => {
      Logger.log(`\n📈 ${account.name}の週次レポート生成中...`);
      updateWeeklyDashboard(account.name);

      // API レート制限対策
      Utilities.sleep(1000);
    });

    Logger.log("\n========================================");
    Logger.log("✅ 全アカウントの週次レポート生成完了");
    Logger.log("========================================");

  } catch (e) {
    Logger.log(`❌ エラー in generateWeeklyReportsAll: ${e.toString()}`);
    handleError("generateWeeklyReportsAll", e, { severity: "HIGH" });
  }
}

/**
 * 週次レポート手動生成（メニューから実行）
 */
function manualGenerateWeeklyReports() {
  try {
    const currentSpreadsheetId = SpreadsheetApp.getActive().getId();
    const account = ACCOUNTS.find(a => a.spreadsheetId === currentSpreadsheetId);

    if (!account) {
      SpreadsheetApp.getUi().alert("❌ エラー: このスプレッドシートに対応するアカウントが見つかりません");
      return;
    }

    Logger.log(`📈 ${account.name}の週次レポート生成開始`);
    updateWeeklyDashboard(account.name);
    Logger.log(`✅ ${account.name}の週次レポート生成完了`);

    SpreadsheetApp.getUi().alert(`✅ ${account.name}の週次レポート生成完了`);
  } catch (e) {
    Logger.log(`❌ エラー in manualGenerateWeeklyReports: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`❌ エラー: ${e.toString()}`);
  }
}
