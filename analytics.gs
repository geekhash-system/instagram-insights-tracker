// ========================================
// analytics.gs - 週次ダッシュボード計算ロジック
// ========================================

/**
 * 週次ダッシュボードを更新
 * @param {string} accountName - アカウント名（NERA / KARA子）
 */
function updateWeeklyDashboard(accountName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const account = ACCOUNTS.find(a => a.name === accountName);

    if (!account) {
      Logger.log(`❌ アカウントが見つかりません: ${accountName}`);
      return;
    }

    const dataSheet = ss.getSheetByName(account.sheetName);
    if (!dataSheet) {
      Logger.log(`ℹ️ データシートがまだありません: ${accountName}`);
      return;
    }

    let dashboardSheet = ss.getSheetByName(account.dashboardSheet);

    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet(account.dashboardSheet);
      initializeDashboardSheet(dashboardSheet);
    }

    // データを取得
    const rows = getSheetData(dataSheet);
    if (rows.length === 0) {
      Logger.log(`ℹ️ データがまだありません: ${accountName}`);
      return;
    }

    // 今週・先週のデータを抽出
    const thisWeekData = filterByWeek(rows, 0);  // 今週
    const lastWeekData = filterByWeek(rows, -1); // 先週

    // オーガニック投稿とPR投稿に分ける
    const thisWeekOrganic = thisWeekData.filter(row => !row[COLUMNS.PR]);
    const lastWeekOrganic = lastWeekData.filter(row => !row[COLUMNS.PR]);
    const thisWeekPR = thisWeekData.filter(row => row[COLUMNS.PR] === true);
    const lastWeekPR = lastWeekData.filter(row => row[COLUMNS.PR] === true);

    // 統計計算
    const stats = {
      thisWeekPostCount: thisWeekData.length,
      lastWeekPostCount: lastWeekData.length,
      thisWeekTotalImp: sumImp(thisWeekData),
      lastWeekTotalImp: sumImp(lastWeekData),
      thisWeekAvgImp: avgImp(thisWeekData),
      lastWeekAvgImp: avgImp(lastWeekData),
      thisWeekMedianImp: medianImp(thisWeekData),
      lastWeekMedianImp: medianImp(lastWeekData),

      // オーガニック
      thisWeekOrganicPostCount: thisWeekOrganic.length,
      lastWeekOrganicPostCount: lastWeekOrganic.length,
      thisWeekOrganicTotalImp: sumImp(thisWeekOrganic),
      lastWeekOrganicTotalImp: sumImp(lastWeekOrganic),
      thisWeekOrganicAvgImp: avgImp(thisWeekOrganic),
      lastWeekOrganicAvgImp: avgImp(lastWeekOrganic),
      thisWeekOrganicMedianImp: medianImp(thisWeekOrganic),
      lastWeekOrganicMedianImp: medianImp(lastWeekOrganic),

      // PR
      thisWeekPRPostCount: thisWeekPR.length,
      lastWeekPRPostCount: lastWeekPR.length,
      thisWeekPRTotalImp: sumImp(thisWeekPR),
      lastWeekPRTotalImp: sumImp(lastWeekPR),
      thisWeekPRAvgImp: avgImp(thisWeekPR),
      lastWeekPRAvgImp: avgImp(lastWeekPR),
      thisWeekPRMedianImp: medianImp(thisWeekPR),
      lastWeekPRMedianImp: medianImp(lastWeekPR)
    };

    // ダッシュボードに書き込み
    writeDashboardStats(dashboardSheet, stats);

    // オーガニック投稿のトップ/ワースト
    writeTopBottomOrganic(dashboardSheet, thisWeekOrganic);

    // PR投稿の警告リスト
    writePRWarnings(dashboardSheet, rows, account);

    Logger.log(`✅ 週次ダッシュボード更新完了: ${accountName}`);

  } catch (e) {
    Logger.log(`エラー in updateWeeklyDashboard: ${e.toString()}`);
  }
}

/**
 * ダッシュボードシートを初期化
 * @param {Sheet} sheet - ダッシュボードシート
 */
function initializeDashboardSheet(sheet) {
  sheet.getRange("A1").setValue("週次ダッシュボード").setFontWeight("bold").setFontSize(16);
  sheet.getRange("A2").setValue("最終更新: ").setFontSize(10);
}

/**
 * ダッシュボード統計を書き込み
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {Object} stats - 統計データ
 */
function writeDashboardStats(sheet, stats) {
  const data = [
    ["【全投稿】", ""],
    ["今週の投稿数", stats.thisWeekPostCount],
    ["先週の投稿数", stats.lastWeekPostCount],
    ["今週の総IMP数", stats.thisWeekTotalImp],
    ["先週の総IMP数", stats.lastWeekTotalImp],
    ["IMP差分", stats.thisWeekTotalImp - stats.lastWeekTotalImp],
    ["IMP差分（%）", stats.lastWeekTotalImp > 0 ? ((stats.thisWeekTotalImp - stats.lastWeekTotalImp) / stats.lastWeekTotalImp * 100).toFixed(1) + "%" : "-"],
    ["今週の平均IMP", Math.round(stats.thisWeekAvgImp)],
    ["先週の平均IMP", Math.round(stats.lastWeekAvgImp)],
    ["今週の中央値IMP", Math.round(stats.thisWeekMedianImp)],
    ["先週の中央値IMP", Math.round(stats.lastWeekMedianImp)],
    ["", ""],
    ["【オーガニック投稿】", ""],
    ["今週の投稿数", stats.thisWeekOrganicPostCount],
    ["今週の総IMP数", stats.thisWeekOrganicTotalImp],
    ["今週の平均IMP", Math.round(stats.thisWeekOrganicAvgImp)],
    ["今週の中央値IMP", Math.round(stats.thisWeekOrganicMedianImp)],
    ["", ""],
    ["【PR投稿】", ""],
    ["今週の投稿数", stats.thisWeekPRPostCount],
    ["今週の総IMP数", stats.thisWeekPRTotalImp],
    ["今週の平均IMP", Math.round(stats.thisWeekPRAvgImp)],
    ["今週の中央値IMP", Math.round(stats.thisWeekPRMedianImp)]
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange("A2").setValue("最終更新: " + new Date().toLocaleString("ja-JP"));
}

/**
 * オーガニック投稿のトップ/ワースト
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {Array} organicPosts - オーガニック投稿配列
 */
function writeTopBottomOrganic(sheet, organicPosts) {
  if (organicPosts.length === 0) return;

  const sorted = organicPosts.sort((a, b) => (b[COLUMNS.IMP_COUNT] || 0) - (a[COLUMNS.IMP_COUNT] || 0));

  const startRow = 30;
  sheet.getRange(startRow, 1).setValue("【オーガニック投稿 トップ5】").setFontWeight("bold");
  sheet.getRange(startRow + 1, 1, 1, 4).setValues([["投稿日時", "キャプション", "IMP数", "リンク"]]).setFontWeight("bold");

  const top5 = sorted.slice(0, 5).map(row => [
    row[COLUMNS.POST_DATE],
    (row[COLUMNS.CAPTION] || "").substring(0, 50),
    row[COLUMNS.IMP_COUNT],
    row[COLUMNS.PERMALINK]
  ]);

  if (top5.length > 0) {
    sheet.getRange(startRow + 2, 1, top5.length, 4).setValues(top5);
  }

  // ワースト5
  const worstStartRow = startRow + 10;
  sheet.getRange(worstStartRow, 1).setValue("【オーガニック投稿 ワースト5】").setFontWeight("bold");
  sheet.getRange(worstStartRow + 1, 1, 1, 4).setValues([["投稿日時", "キャプション", "IMP数", "リンク"]]).setFontWeight("bold");

  const worst5 = sorted.slice(-5).reverse().map(row => [
    row[COLUMNS.POST_DATE],
    (row[COLUMNS.CAPTION] || "").substring(0, 50),
    row[COLUMNS.IMP_COUNT],
    row[COLUMNS.PERMALINK]
  ]);

  if (worst5.length > 0) {
    sheet.getRange(worstStartRow + 2, 1, worst5.length, 4).setValues(worst5);
  }
}

/**
 * PR投稿の警告リストを作成
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {Array} allRows - 全データ行
 * @param {Object} account - アカウント設定
 */
function writePRWarnings(sheet, allRows, account) {
  try {
    // PR投稿のみ抽出（新しい順）
    const prPosts = allRows
      .filter(row => row[COLUMNS.PR] === true)
      .sort((a, b) => new Date(b[COLUMNS.POST_DATE]) - new Date(a[COLUMNS.POST_DATE]));

    if (prPosts.length === 0) {
      Logger.log(`ℹ️ PR投稿がありません: ${account.name}`);
      return;
    }

    // 過去10投稿の中央値を計算
    const last10Posts = prPosts.slice(0, DASHBOARD_CONFIG.PR_MEDIAN_POSTS_COUNT);
    const median = medianImp(last10Posts);
    const threshold = median * DASHBOARD_CONFIG.PR_WARNING_THRESHOLD;

    Logger.log(`📊 PR投稿中央値: ${median}, 最低ライン: ${threshold}`);

    // 警告対象の投稿を抽出
    const warnings = prPosts
      .slice(0, 20)
      .filter(row => (row[COLUMNS.IMP_COUNT] || 0) < threshold)
      .map(row => [
        row[COLUMNS.POST_DATE],
        (row[COLUMNS.CAPTION] || "").substring(0, 50),
        row[COLUMNS.IMP_COUNT],
        Math.round(median),
        Math.round(threshold),
        row[COLUMNS.PERMALINK]
      ]);

    // ダッシュボードに書き込み
    const warningStartRow = 50;
    sheet.getRange(warningStartRow, 1).setValue("【PR投稿警告リスト】").setFontWeight("bold").setFontSize(14);
    sheet.getRange(warningStartRow + 1, 1, 1, 6).setValues([
      ["投稿日時", "キャプション", "IMP数", "中央値", "最低ライン", "リンク"]
    ]).setFontWeight("bold").setBackground("#FFD700");

    if (warnings.length > 0) {
      sheet.getRange(warningStartRow + 2, 1, warnings.length, 6).setValues(warnings);
      sheet.getRange(warningStartRow + 2, 1, warnings.length, 6).setBackground("#FFCCCC");
      Logger.log(`⚠️ PR投稿警告: ${warnings.length} 件`);
    } else {
      sheet.getRange(warningStartRow + 2, 1).setValue("警告なし").setFontColor("#00AA00");
    }

  } catch (e) {
    Logger.log(`エラー in writePRWarnings: ${e.toString()}`);
  }
}

/**
 * IMP数の合計
 */
function sumImp(rows) {
  return rows.reduce((sum, row) => sum + (row[COLUMNS.IMP_COUNT] || 0), 0);
}

/**
 * IMP数の平均
 */
function avgImp(rows) {
  if (rows.length === 0) return 0;
  return sumImp(rows) / rows.length;
}

/**
 * IMP数の中央値
 */
function medianImp(rows) {
  if (rows.length === 0) return 0;
  const sorted = rows.map(row => row[COLUMNS.IMP_COUNT] || 0).sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  }
  return sorted[mid];
}

/**
 * 週でフィルタリング
 * @param {Array} rows - データ行
 * @param {number} weekOffset - 0=今週、-1=先週
 */
function filterByWeek(rows, weekOffset) {
  const now = new Date();
  const startOfWeek = new Date(now);
  startOfWeek.setDate(now.getDate() - now.getDay() + weekOffset * 7);
  startOfWeek.setHours(0, 0, 0, 0);

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 7);

  return rows.filter(row => {
    const postDate = new Date(row[COLUMNS.POST_DATE]);
    return postDate >= startOfWeek && postDate < endOfWeek;
  });
}
