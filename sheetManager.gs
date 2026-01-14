// ========================================
// sheetManager.gs - スプレッドシート操作・データ管理
// ========================================
// GAS_instagram_reel_viewcount_tracker の sheetManager.gs を参考

/**
 * アカウントシートを初期化
 * @param {Sheet} sheet - 対象シート
 */
function initializeAccountSheet(sheet) {
  // ヘッダー設定（1行目）- メディアIDを最後に移動
  const headers = [
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
    "メディアID",
    "履歴→"
  ];

  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#D3D3D3");

  // F列（PR列）にチェックボックスを設定（2行目から）- G列→F列に移動
  sheet.getRange("F2:F1000").insertCheckboxes();

  // 列幅調整（列順序変更に対応）
  sheet.setColumnWidth(1, 150);  // 投稿日時
  sheet.setColumnWidth(2, 50);   // 曜日
  sheet.setColumnWidth(3, 100);  // 投稿タイプ
  sheet.setColumnWidth(4, 300);  // キャプション
  sheet.setColumnWidth(5, 250);  // パーマリンク
  sheet.setColumnWidth(6, 60);   // PR
  sheet.setColumnWidth(7, 100);  // IMP数
  sheet.setColumnWidth(8, 100);  // リーチ数
  sheet.setColumnWidth(9, 100);  // いいね数
  sheet.setColumnWidth(10, 100); // コメント数
  sheet.setColumnWidth(11, 100); // 保存数
  sheet.setColumnWidth(12, 100); // シェア数
  sheet.setColumnWidth(13, 120); // エンゲージメント数
  sheet.setColumnWidth(14, 150); // 最終更新日時
  sheet.setColumnWidth(15, 150); // メディアID
  sheet.setColumnWidth(16, 80);  // 履歴→

  // キャプション列（D列）のテキスト折り返しと上揃えを設定
  sheet.getRange("D:D")
    .setWrap(true)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("top");

  // 全データ範囲にCLIP戦略を適用（行の高さ自動調整を防ぐ）
  sheet.getRange("2:1000")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 全データ行の高さを40ピクセルに設定（高速化版）
  sheet.setRowHeights(2, 999, 40);

  // 数値列にカンマ区切りフォーマットを適用（2行目以降、1000行まで）
  // G列: IMP数, H列: リーチ数, I列: いいね数, J列: コメント数, K列: 保存数, L列: シェア数, M列: エンゲージメント数
  sheet.getRange("G2:M1000").setNumberFormat("#,##0");
  // P列以降の履歴列もカンマ区切り
  sheet.getRange("P2:Z1000").setNumberFormat("#,##0");
}

/**
 * アカウントインサイト履歴シートを初期化
 * @param {Sheet} sheet - 対象シート
 */
function initializeAccountInsightsSheet(sheet) {
  // 説明行（1行目）
  const description = [
    "データ取得日",
    "その日の値",
    "前日比",
    "その日の値",
    "その日の値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値",
    "前日（0時～24時）の集計値"
  ];

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

  // 説明行を書き込み（1行目）
  const descRow = sheet.getRange(1, 1, 1, description.length);
  descRow.setValues([description]);
  descRow.setFontSize(9);
  descRow.setFontColor("#666666");
  descRow.setBackground("#E3F2FD");

  // ヘッダー行を書き込み（2行目）
  const headerRow = sheet.getRange(2, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#B3E5FC");

  // Set column widths
  sheet.setColumnWidth(1, 120);  // 日付
  for (let i = 2; i <= 14; i++) {
    sheet.setColumnWidth(i, 140);
  }

  // Set row heights to default (21 pixels) for all data rows (高速化版)
  sheet.setRowHeights(1, 1000, 21);

  // Number formatting (データは3行目から)
  sheet.getRange("B3:N1000").setNumberFormat("#,##0");

  // Conditional formatting for follower change (C column、3行目から)
  const followerChangeRange = sheet.getRange("C3:C1000");
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
    // データは3行目から開始（1行目: 説明行、2行目: ヘッダー）
    let previousFollowerCount = 0;
    if (data.length > 2) {
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
      insights ? (insights.saves || 0) : 0,
      insights ? (insights.shares || 0) : 0,
      insights ? (insights.replies || 0) : 0,
      insights ? (insights.profile_links_taps || 0) : 0
    ];

    // Check if date already exists (データは3行目から: i=2から開始)
    let existingRow = null;
    for (let i = 2; i < data.length; i++) {
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
      insights ? (insights.saves || 0) : 0,
      insights ? (insights.shares || 0) : 0,
      engagement,
      new Date(),
      mediaId  // メディアIDを最後に移動
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
    const historyHeaderRow = 1; // ヘッダー行は1行目

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

    // 各行のIMP数を記録（ヘッダー行をスキップするため、i=1から開始）
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
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
      const headerValue = sheet.getRange(1, col).getValue(); // ヘッダー行は1行目
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
 * シートを投稿日時で降順ソート（新しい投稿が上）
 * @param {Sheet} sheet - 対象シート
 */
function sortSheetByDateDesc(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return; // データがない場合はスキップ

    // ソート対象をA～O列（メディアIDまで）に制限
    // 履歴列（P列以降）はソート不要で、含めるとタイムアウトの原因になる
    const sortColumns = COLUMNS.MEDIA_ID + 1; // 15列（A～O）
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sortColumns);

    // 1列（投稿日時）で降順ソート
    dataRange.sort({column: 1, ascending: false});

    Logger.log(`📊 シートを投稿日時で降順ソート完了`);
  } catch (e) {
    Logger.log(`エラー in sortSheetByDateDesc: ${e.toString()}`);
  }
}
