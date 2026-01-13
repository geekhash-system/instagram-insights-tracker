// ========================================
// analytics.gs - 週次ダッシュボード計算ロジック
// ========================================

/**
 * 週次ダッシュボードを更新
 * @param {string} accountName - アカウント名（NERA / KARA子）
 */
function updateWeeklyDashboard(accountName) {
  try {
    const account = ACCOUNTS.find(a => a.name === accountName);

    if (!account) {
      Logger.log(`❌ アカウントが見つかりません: ${accountName}`);
      return;
    }

    const ss = SpreadsheetApp.openById(account.spreadsheetId);
    const dataSheet = ss.getSheetByName(account.sheetName);
    if (!dataSheet) {
      Logger.log(`ℹ️ データシートがまだありません: ${accountName}`);
      return;
    }

    // 週次ダッシュボードシート名を生成（作成日のmmdd形式）
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const weekSheetName = `週次レポート_${month}${day}`;

    let dashboardSheet = ss.getSheetByName(weekSheetName);

    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet(weekSheetName);
      Logger.log(`📊 新しい週次シートを作成: ${weekSheetName}`);
    }

    // 2週間前の期間を計算（Monday-Sunday）
    // 分析対象は「2週間前」（投稿の7日目データが完全に揃っている週）
    const startOfWeek = new Date(now);
    const daysSinceMonday = (now.getDay() + 6) % 7;
    startOfWeek.setDate(now.getDate() - daysSinceMonday - 14); // 2週間前の月曜日
    startOfWeek.setHours(0, 0, 0, 0);
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    endOfWeek.setHours(23, 59, 59, 999);

    const dateRange = `${formatDate(startOfWeek)} - ${formatDate(endOfWeek)}`;

    initializeDashboardSheet(dashboardSheet, dateRange);

    // データを取得
    const rows = getSheetData(dataSheet);
    if (rows.length === 0) {
      Logger.log(`ℹ️ データがまだありません: ${accountName}`);
      return;
    }

    // 2週間前・3週間前のデータを抽出（レポート対象は2週間前）
    const thisWeekData = filterByWeek(rows, -2);  // 2週間前（今回の分析対象）
    const lastWeekData = filterByWeek(rows, -3); // 3週間前（比較用）

    Logger.log(`📊 全データ件数: ${rows.length}`);
    Logger.log(`📅 分析対象週（2週間前）のデータ件数: ${thisWeekData.length}`);
    Logger.log(`📅 比較対象週（3週間前）のデータ件数: ${lastWeekData.length}`);

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

    // ダッシュボードに書き込み（分析対象期間を渡す）
    writeDashboardStats(dashboardSheet, stats, startOfWeek, endOfWeek);

    // オーガニック投稿のトップ/ワースト
    writeTopBottomOrganic(dashboardSheet, thisWeekOrganic);

    // PR投稿の警告リスト
    writePRWarnings(dashboardSheet, rows, account);

    // 注: 週次シートは削除せず、永久保存されます

    Logger.log(`✅ 週次ダッシュボード更新完了: ${weekSheetName}`);

  } catch (e) {
    Logger.log(`エラー in updateWeeklyDashboard: ${e.toString()}`);
  }
}

/**
 * ダッシュボードシートを初期化
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {string} dateRange - 期間（例: 2026/01/05 - 2026/01/11）
 */
function initializeDashboardSheet(sheet, dateRange) {
  // Row 1: Title
  sheet.getRange("A1").setValue(`週次ダッシュボード（${dateRange}）`)
    .setFontWeight("bold")
    .setFontSize(16);

  // Row 2: Report Generation Date (レポート生成日)
  const now = new Date();
  const generationDate = formatDate(now);
  sheet.getRange("A2").setValue("レポート生成日:").setFontWeight("bold").setFontSize(10);
  sheet.getRange("B2").setValue(generationDate).setFontSize(10);

  // Row 3: Measurement Period (計測対象期間)
  sheet.getRange("A3").setValue("計測対象期間:").setFontWeight("bold").setFontSize(10);
  sheet.getRange("B3").setValue(dateRange).setFontSize(10);

  // Row 4: Blank row for spacing
}

/**
 * ダッシュボード統計を書き込み
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {Object} stats - 統計データ
 * @param {Date} thisWeekStart - 計測対象週の開始日（2週間前の月曜日）
 * @param {Date} thisWeekEnd - 計測対象週の終了日（2週間前の日曜日）
 */
function writeDashboardStats(sheet, stats, thisWeekStart, thisWeekEnd) {
  // 計測対象週（2週間前）と比較対象週（3週間前）の期間を計算
  const lastWeekStart = new Date(thisWeekStart);
  lastWeekStart.setDate(thisWeekStart.getDate() - 7);
  const lastWeekEnd = new Date(thisWeekEnd);
  lastWeekEnd.setDate(thisWeekEnd.getDate() - 7);

  const thisWeekRange = `${formatDate(thisWeekStart)} - ${formatDate(thisWeekEnd)}`;
  const lastWeekRange = `${formatDate(lastWeekStart)} - ${formatDate(lastWeekEnd)}`;

  const data = [
    ["【全投稿】", ""],
    [`計測対象期間の投稿数（${thisWeekRange}）`, stats.thisWeekPostCount],
    [`比較対象期間の投稿数（${lastWeekRange}）`, stats.lastWeekPostCount],
    ["計測対象の総IMP数", stats.thisWeekTotalImp],
    ["比較対象の総IMP数", stats.lastWeekTotalImp],
    ["IMP差分", stats.thisWeekTotalImp - stats.lastWeekTotalImp],
    ["IMP差分（%）", stats.lastWeekTotalImp > 0 ? ((stats.thisWeekTotalImp - stats.lastWeekTotalImp) / stats.lastWeekTotalImp * 100).toFixed(1) + "%" : "-"],
    ["計測対象の平均IMP", Math.round(stats.thisWeekAvgImp)],
    ["比較対象の平均IMP", Math.round(stats.lastWeekAvgImp)],
    ["平均IMP差分", Math.round(stats.thisWeekAvgImp - stats.lastWeekAvgImp)],
    ["平均IMP差分（%）", stats.lastWeekAvgImp > 0 ? ((stats.thisWeekAvgImp - stats.lastWeekAvgImp) / stats.lastWeekAvgImp * 100).toFixed(1) + "%" : "-"],
    ["計測対象の中央値IMP", Math.round(stats.thisWeekMedianImp)],
    ["比較対象の中央値IMP", Math.round(stats.lastWeekMedianImp)],
    ["中央値IMP差分", Math.round(stats.thisWeekMedianImp - stats.lastWeekMedianImp)],
    ["中央値IMP差分（%）", stats.lastWeekMedianImp > 0 ? ((stats.thisWeekMedianImp - stats.lastWeekMedianImp) / stats.lastWeekMedianImp * 100).toFixed(1) + "%" : "-"],
    ["", ""],
    ["【オーガニック投稿】", ""],
    ["計測対象期間の投稿数", stats.thisWeekOrganicPostCount],
    ["計測対象の総IMP数", stats.thisWeekOrganicTotalImp],
    ["計測対象の平均IMP", Math.round(stats.thisWeekOrganicAvgImp)],
    ["計測対象の中央値IMP", Math.round(stats.thisWeekOrganicMedianImp)],
    ["", ""],
    ["【PR投稿】", ""],
    ["計測対象期間の投稿数", stats.thisWeekPRPostCount],
    ["計測対象の総IMP数", stats.thisWeekPRTotalImp],
    ["計測対象の平均IMP", Math.round(stats.thisWeekPRAvgImp)],
    ["計測対象の中央値IMP", Math.round(stats.thisWeekPRMedianImp)]
  ];

  sheet.getRange(5, 1, data.length, 2).setValues(data);

  // 数値列にカンマ区切りフォーマットを適用
  sheet.getRange(5, 2, data.length, 1).setNumberFormat("#,##0");
}

/**
 * オーガニック投稿のトップ/ワースト
 * @param {Sheet} sheet - ダッシュボードシート
 * @param {Array} organicPosts - オーガニック投稿配列
 */
function writeTopBottomOrganic(sheet, organicPosts) {
  if (organicPosts.length === 0) return;

  const sorted = organicPosts.sort((a, b) => (b[COLUMNS.IMP_COUNT] || 0) - (a[COLUMNS.IMP_COUNT] || 0));

  const startRow = 32;  // +1 for blank row before organic posts section
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
    // IMP数列にカンマ区切りフォーマットを適用
    sheet.getRange(startRow + 2, 3, top5.length, 1).setNumberFormat("#,##0");
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
    // IMP数列にカンマ区切りフォーマットを適用
    sheet.getRange(worstStartRow + 2, 3, worst5.length, 1).setNumberFormat("#,##0");
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
    const warningStartRow = 51;  // Adjusted for new metadata rows
    sheet.getRange(warningStartRow, 1).setValue("【PR投稿警告リスト】").setFontWeight("bold").setFontSize(14);
    sheet.getRange(warningStartRow + 1, 1, 1, 6).setValues([
      ["投稿日時", "キャプション", "IMP数", "中央値", "最低ライン", "リンク"]
    ]).setFontWeight("bold").setBackground("#FFD700");

    if (warnings.length > 0) {
      sheet.getRange(warningStartRow + 2, 1, warnings.length, 6).setValues(warnings);
      sheet.getRange(warningStartRow + 2, 1, warnings.length, 6).setBackground("#FFCCCC");
      // IMP数、中央値、最低ライン列にカンマ区切りフォーマットを適用
      sheet.getRange(warningStartRow + 2, 3, warnings.length, 3).setNumberFormat("#,##0");
      Logger.log(`⚠️ PR投稿警告: ${warnings.length} 件`);
    } else {
      sheet.getRange(warningStartRow + 2, 1).setValue("警告なし").setFontColor("#00AA00");
    }

  } catch (e) {
    Logger.log(`エラー in writePRWarnings: ${e.toString()}`);
  }
}

/**
 * 履歴列から投稿後7日目のIMP数を取得
 * @param {Array} row - データ行
 * @return {number} 7日目のIMP数
 */
function getDay7Imp(row) {
  const historyStartCol = COLUMNS.HISTORY_START; // P列（15列目）

  // P列が1日目、Q列が2日目、...、V列が7日目
  const day7Index = historyStartCol + 6; // 7日目の列インデックス

  if (row[day7Index] !== undefined && row[day7Index] !== "" && row[day7Index] !== null) {
    const value = parseInt(row[day7Index]);
    if (!isNaN(value)) {
      return value;
    }
  }

  // 7日目のデータがない場合は現在のIMP数を使用（フォールバック）
  return row[COLUMNS.IMP_COUNT] || 0;
}

/**
 * IMP数の合計
 */
function sumImp(rows) {
  return rows.reduce((sum, row) => sum + getDay7Imp(row), 0);
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
  const sorted = rows.map(row => getDay7Imp(row)).sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  }
  return sorted[mid];
}

/**
 * 週でフィルタリング（Monday-Sunday week）
 * @param {Array} rows - データ行
 * @param {number} weekOffset - 0=今週、-1=先週
 */
function filterByWeek(rows, weekOffset) {
  const now = new Date();
  const startOfWeek = new Date(now);
  // Monday-Sunday: Calculate days since Monday
  const daysSinceMonday = (now.getDay() + 6) % 7;  // Mon=0, Tue=1, ..., Sun=6
  startOfWeek.setDate(now.getDate() - daysSinceMonday + weekOffset * 7);
  startOfWeek.setHours(0, 0, 0, 0);

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 7);

  return rows.filter(row => {
    const postDate = new Date(row[COLUMNS.POST_DATE]);
    return postDate >= startOfWeek && postDate < endOfWeek;
  });
}

/**
 * 週番号を取得（Monday-Sunday week）
 * @param {Date} date - 日付
 * @return {number} 週番号
 */
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  // Monday-Sunday week: Monday=1, Sunday=7
  const dayNum = d.getUTCDay() === 0 ? 7 : d.getUTCDay();
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * 日付をフォーマット（YYYY/MM/DD）
 * @param {Date} date - 日付
 * @return {string} フォーマット済み日付
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 古い週次ダッシュボードシートを削除
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {number} weeksToKeep - 保持する週数（デフォルト: 13週）
 */
function cleanupOldWeeklySheets(ss, weeksToKeep = 13) {
  try {
    const now = new Date();
    const cutoffDate = new Date(now);
    cutoffDate.setDate(now.getDate() - (weeksToKeep * 7));

    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      // 週次シート名のパターンマッチ: 週次_NERA_2026W02
      const match = sheetName.match(/^週次_(.+?)_(\d{4})W(\d{2})$/);
      if (match) {
        const year = parseInt(match[2]);
        const week = parseInt(match[3]);

        // 週番号から日付を推定（年の最初の日曜日 + 週数）
        const sheetDate = new Date(year, 0, 1 + (week - 1) * 7);

        if (sheetDate < cutoffDate) {
          ss.deleteSheet(sheet);
          Logger.log(`🗑️ 古い週次シートを削除: ${sheetName}`);
        }
      }
    });
  } catch (e) {
    Logger.log(`エラー in cleanupOldWeeklySheets: ${e.toString()}`);
  }
}
