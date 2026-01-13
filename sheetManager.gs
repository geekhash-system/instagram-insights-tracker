// ========================================
// sheetManager.gs - スプレッドシート操作・データ管理
// ========================================
// GAS_instagram_reel_viewcount_tracker の sheetManager.gs を参考

/**
 * アカウントシートを初期化
 * @param {Sheet} sheet - 対象シート
 */
function initializeAccountSheet(sheet) {
  // ヘッダー設定（1行目）
  const headers = [
    "メディアID",
    "投稿日時",
    "曜日",
    "投稿タイプ",
    "キャプション",
    "パーマリンク",
    "PR",
    "IMP数",
    "リーチ数",
    "いいね数",
    "コメント数",
    "保存数",
    "シェア数",
    "エンゲージメント数",
    "最終更新日時",
    "履歴→"
  ];

  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#D3D3D3");

  // G列（PR列）にチェックボックスを設定（2行目から）
  sheet.getRange("G2:G1000").insertCheckboxes();

  // 列幅調整
  sheet.setColumnWidth(1, 150);  // メディアID
  sheet.setColumnWidth(2, 150);  // 投稿日時
  sheet.setColumnWidth(3, 50);   // 曜日
  sheet.setColumnWidth(4, 100);  // 投稿タイプ
  sheet.setColumnWidth(5, 300);  // キャプション
  sheet.setColumnWidth(6, 250);  // パーマリンク
  sheet.setColumnWidth(7, 60);   // PR
  sheet.setColumnWidth(8, 100);  // IMP数
  sheet.setColumnWidth(9, 100);  // リーチ数
  sheet.setColumnWidth(10, 100); // いいね数
  sheet.setColumnWidth(11, 100); // コメント数
  sheet.setColumnWidth(12, 100); // 保存数
  sheet.setColumnWidth(13, 100); // シェア数
  sheet.setColumnWidth(14, 120); // エンゲージメント数
  sheet.setColumnWidth(15, 150); // 最終更新日時
  sheet.setColumnWidth(16, 80);  // 履歴→

  // キャプション列（E列）のテキスト折り返しと上揃えを設定
  sheet.getRange("E2:E1000").setWrap(true).setVerticalAlignment("top");

  // 全データ行の高さを21ピクセルに設定（デフォルトより低く）
  for (let row = 2; row <= 1000; row++) {
    sheet.setRowHeight(row, 21);
  }

  // 数値列にカンマ区切りフォーマットを適用（2行目以降、1000行まで）
  // H列: IMP数, I列: リーチ数, J列: いいね数, K列: コメント数, L列: 保存数, M列: シェア数, N列: エンゲージメント数
  sheet.getRange("H2:N1000").setNumberFormat("#,##0");
  // P列以降の履歴列もカンマ区切り
  sheet.getRange("P2:Z1000").setNumberFormat("#,##0");
}

/**
 * アカウントインサイト履歴シートを初期化
 * @param {Sheet} sheet - 対象シート
 */
function initializeAccountInsightsSheet(sheet) {
  const headers = [
    "日付",
    "フォロワー数",
    "フォロワー増減数",
    "フォロー数",
    "投稿数",
    "リーチ数",
    "エンゲージしたアカウント数",
    "総エンゲージメント数",
    "いいね数",
    "コメント数",
    "保存数",
    "シェア数",
    "返信数",
    "プロフィールリンクタップ数"
  ];

  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#B3E5FC");

  // Set column widths
  sheet.setColumnWidth(1, 120);  // 日付
  for (let i = 2; i <= 14; i++) {
    sheet.setColumnWidth(i, 140);
  }

  // Number formatting
  sheet.getRange("B2:N1000").setNumberFormat("#,##0");

  // Conditional formatting for follower change (C column)
  const followerChangeRange = sheet.getRange("C2:C1000");
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor("#0F9D58")
    .setRanges([followerChangeRange])
    .build();
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor("#DB4437")
    .setRanges([followerChangeRange])
    .build();

  sheet.setConditionalFormatRules([positiveRule, negativeRule]);
}

/**
 * アカウントインサイトデータを記録
 * @param {Sheet} sheet - アカウントインサイトシート
 * @param {string} date - 日付（YYYY-MM-DD）
 * @param {Object} accountInfo - アカウント情報（フォロワー数など）
 * @param {Object} insights - インサイトデータ
 */
function addAccountInsightsRecord(sheet, date, accountInfo, insights) {
  try {
    const data = sheet.getDataRange().getValues();

    // Get previous day's follower count for change calculation
    let previousFollowerCount = 0;
    if (data.length > 1) {
      previousFollowerCount = data[data.length - 1][ACCOUNT_INSIGHTS_COLUMNS.FOLLOWER_COUNT] || 0;
    }

    const currentFollowerCount = accountInfo ? accountInfo.followers_count : 0;
    const followerChange = currentFollowerCount - previousFollowerCount;
    const followsCount = accountInfo ? accountInfo.follows_count : 0;
    const mediaCount = accountInfo ? accountInfo.media_count : 0;

    // Format date as YYYY/MM/DD
    const dateParts = date.split("-");
    const formattedDate = `${dateParts[0]}/${dateParts[1]}/${dateParts[2]}`;

    const rowData = [
      formattedDate,
      currentFollowerCount,
      followerChange,
      followsCount,
      mediaCount,
      insights ? (insights.reach || 0) : 0,
      insights ? (insights.accounts_engaged || 0) : 0,
      insights ? (insights.total_interactions || 0) : 0,
      insights ? (insights.likes || 0) : 0,
      insights ? (insights.comments || 0) : 0,
      insights ? (insights.saved || 0) : 0,
      insights ? (insights.shares || 0) : 0,
      insights ? (insights.replies || 0) : 0,
      insights ? (insights.profile_links_taps || 0) : 0
    ];

    // Check if date already exists
    let existingRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][ACCOUNT_INSIGHTS_COLUMNS.DATE] === formattedDate) {
        existingRow = i + 1;
        break;
      }
    }

    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      Logger.log(`🔄 アカウントインサイト更新: ${formattedDate}`);
    } else {
      sheet.appendRow(rowData);
      Logger.log(`➕ アカウントインサイト追加: ${formattedDate}`);
    }
  } catch (e) {
    Logger.log(`エラー in addAccountInsightsRecord: ${e.toString()}`);
  }
}

/**
 * メディアデータを更新（新規追加 or 既存更新）
 * @param {Sheet} sheet - 対象シート
 * @param {Object} media - メディアデータ
 * @param {Object} insights - インサイトデータ
 */
function updateMediaData(sheet, media, insights) {
  try {
    const mediaId = media.id;
    const data = sheet.getDataRange().getValues();

    // 既存行を検索（ヘッダー行をスキップするため、i=1から開始）
    let targetRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLUMNS.MEDIA_ID] === mediaId) {
        targetRow = i + 1; // 1-indexed
        break;
      }
    }

    // データ作成
    // hinome-backend実装に基づく: リールもフィードも"views"メトリクスを使用
    const timestamp = new Date(media.timestamp);
    const dayOfWeek = getDayOfWeekJapanese(timestamp);
    const impressions = insights ? (insights.views || 0) : 0;  // viewsを使用
    const reach = insights ? (insights.reach || 0) : 0;
    const engagement = insights ? (insights.total_interactions || 0) : 0;

    const rowData = [
      mediaId,
      timestamp,
      dayOfWeek,
      media.media_product_type || media.media_type,
      media.caption || "",
      media.permalink || "",
      false, // PR（デフォルトfalse、手動で変更）
      impressions,
      reach,
      media.like_count || 0,
      media.comments_count || 0,
      insights ? (insights.saved || 0) : 0,
      insights ? (insights.shares || 0) : 0,
      engagement,
      new Date()
    ];

    if (targetRow) {
      // 既存行を更新（PR列は保持）
      const existingPR = sheet.getRange(targetRow, COLUMNS.PR + 1).getValue();
      rowData[COLUMNS.PR] = existingPR; // PR列を保持
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      Logger.log(`🔄 更新: ${mediaId}`);
    } else {
      // 新規行を追加
      sheet.appendRow(rowData);
      Logger.log(`➕ 新規追加: ${mediaId}`);
    }

  } catch (e) {
    Logger.log(`エラー in updateMediaData: ${e.toString()}`);
  }
}

/**
 * 日次履歴を記録（O列以降）
 * @param {Sheet} sheet - 対象シート
 * @param {string} date - 日付（YYYY-MM-DD）
 * @param {string} time - 時刻（HH:mm）
 */
function addHistoryRecord(sheet, date, time) {
  try {
    const historyStartCol = COLUMNS.HISTORY_START + 1; // O列（1-indexed）
    const historyHeaderRow = 2; // ヘッダー行は2行目に変更

    // 日付フォーマット: "12/25取得"
    const dateParts = date.split("-");
    const dateFormatted = dateParts[1] + "/" + dateParts[2] + "取得";

    // 既存の日付列を検索
    const lastCol = sheet.getLastColumn();
    let targetCol = null;

    for (let col = historyStartCol; col <= lastCol; col++) {
      const headerValue = sheet.getRange(historyHeaderRow, col).getValue();
      if (headerValue && headerValue.toString() === dateFormatted) {
        targetCol = col;
        break;
      }
    }

    // 新規列を作成
    if (targetCol === null) {
      targetCol = lastCol + 1;
      sheet.getRange(historyHeaderRow, targetCol)
        .setValue(dateFormatted)
        .setFontWeight("bold")
        .setBackground("#D3D3D3");
    }

    // 各行のIMP数を記録（説明行とヘッダー行をスキップするため、i=2から開始）
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const impCount = data[i][COLUMNS.IMP_COUNT]; // H列（IMP数）
      if (impCount) {
        sheet.getRange(i + 1, targetCol).setValue(impCount);
      }
    }

    Logger.log(`📝 履歴記録完了: ${dateFormatted}`);

    // 注: 履歴列は削除せず、永久保存されます

  } catch (e) {
    Logger.log(`エラー in addHistoryRecord: ${e.toString()}`);
  }
}

/**
 * 古い履歴列を削除
 * @param {Sheet} sheet - 対象シート
 * @param {number} daysToKeep - 保持する日数
 */
function cleanupOldHistoryColumns(sheet, daysToKeep) {
  try {
    const historyStartCol = COLUMNS.HISTORY_START + 1;
    const lastCol = sheet.getLastColumn();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

    const columnsToDelete = [];

    for (let col = historyStartCol; col <= lastCol; col++) {
      const headerValue = sheet.getRange(2, col).getValue(); // ヘッダー行は2行目
      if (!headerValue) continue;

      // "12/25取得" → Date オブジェクト
      const match = headerValue.toString().match(/(\d+)\/(\d+)取得/);
      if (match) {
        const month = parseInt(match[1]);
        const day = parseInt(match[2]);
        const year = new Date().getFullYear();
        const columnDate = new Date(year, month - 1, day);

        if (columnDate < cutoffDate) {
          columnsToDelete.push(col);
        }
      }
    }

    // 後ろから削除（列番号のずれ防止）
    for (let i = columnsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteColumn(columnsToDelete[i]);
      Logger.log(`🗑️ 古い履歴列を削除: ${columnsToDelete[i]}`);
    }

  } catch (e) {
    Logger.log(`エラー in cleanupOldHistoryColumns: ${e.toString()}`);
  }
}

/**
 * シートのデータを全て取得
 * @param {Sheet} sheet - 対象シート
 * @return {Array} データ配列（ヘッダー除く）
 */
function getSheetData(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    return data.slice(1); // ヘッダー除く
  } catch (e) {
    Logger.log(`エラー in getSheetData: ${e.toString()}`);
    return [];
  }
}

/**
 * シートを投稿日時で降順ソート（新しい投稿が上）
 * @param {Sheet} sheet - 対象シート
 */
function sortSheetByDateDesc(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) return; // データがない場合はスキップ

    // データ範囲を取得（説明行1行目、ヘッダー2行目を除く、3行目から）
    const dataRange = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn());

    // B列（投稿日時）で降順ソート
    dataRange.sort({column: 2, ascending: false});

    Logger.log(`📊 シートを投稿日時で降順ソート完了`);
  } catch (e) {
    Logger.log(`エラー in sortSheetByDateDesc: ${e.toString()}`);
  }
}
