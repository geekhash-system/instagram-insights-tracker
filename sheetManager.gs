// ========================================
// sheetManager.gs - スプレッドシート操作・データ管理
// ========================================
// GAS_instagram_reel_viewcount_tracker の sheetManager.gs を参考

/**
 * アカウントシートを初期化
 * @param {Sheet} sheet - 対象シート
 */
function initializeAccountSheet(sheet) {
  // 説明行を追加（1行目）
  const explanation = `このシートはInstagramインサイトデータを記録しています。毎日19時に自動更新されます。`;

  sheet.getRange(1, 1).setValue(explanation);
  sheet.getRange(1, 1).setFontWeight("bold").setBackground("#FFF9C4").setFontSize(10);
  sheet.getRange(1, 1, 1, 15).merge(); // A1からO1までマージ

  // ヘッダー設定（2行目に移動）
  const headers = [
    "メディアID",
    "投稿日時",
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

  const headerRow = sheet.getRange(2, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#D3D3D3");

  // F列（PR列）にチェックボックスを設定（3行目から）
  sheet.getRange("F3:F1000").insertCheckboxes();

  // 列幅調整
  sheet.setColumnWidth(1, 150);  // メディアID
  sheet.setColumnWidth(2, 150);  // 投稿日時
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
  sheet.setColumnWidth(15, 80);  // 履歴→

  // 1行目の高さを調整
  sheet.setRowHeight(1, 30);

  // 数値列にカンマ区切りフォーマットを適用（3行目以降、1000行まで）
  // G列: IMP数, H列: リーチ数, I列: いいね数, J列: コメント数, K列: 保存数, L列: シェア数, M列: エンゲージメント数
  sheet.getRange("G3:M1000").setNumberFormat("#,##0");
  // O列以降の履歴列もカンマ区切り
  sheet.getRange("O3:Z1000").setNumberFormat("#,##0");
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

    // 既存行を検索（説明行とヘッダー行をスキップするため、i=2から開始）
    let targetRow = null;
    for (let i = 2; i < data.length; i++) {
      if (data[i][COLUMNS.MEDIA_ID] === mediaId) {
        targetRow = i + 1; // 1-indexed
        break;
      }
    }

    // データ作成
    // hinome-backend実装に基づく: リールもフィードも"views"メトリクスを使用
    const timestamp = new Date(media.timestamp);
    const impressions = insights ? (insights.views || 0) : 0;  // viewsを使用
    const reach = insights ? (insights.reach || 0) : 0;
    const engagement = insights ? (insights.total_interactions || 0) : 0;

    const rowData = [
      mediaId,
      timestamp,
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
      const impCount = data[i][COLUMNS.IMP_COUNT]; // G列（IMP数）
      if (impCount) {
        sheet.getRange(i + 1, targetCol).setValue(impCount);
      }
    }

    Logger.log(`📝 履歴記録完了: ${dateFormatted}`);

    // 古い履歴の削除（90日以上前）
    cleanupOldHistoryColumns(sheet, 90);

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
