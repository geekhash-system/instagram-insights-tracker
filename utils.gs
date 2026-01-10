// ========================================
// utils.gs - ユーティリティ関数
// ========================================

/**
 * 現在の日時を取得
 * @return {Object} {date: "YYYY-MM-DD", time: "HH:mm", datetime: Date}
 */
function getCurrentDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");

  return {
    date: `${year}-${month}-${day}`,
    time: `${hours}:${minutes}`,
    datetime: now
  };
}

/**
 * 数値をカンマ区切りでフォーマット
 * @param {number} num - 数値
 * @return {string} フォーマット済み文字列
 */
function formatNumber(num) {
  if (!num) return "0";
  return num.toLocaleString();
}

/**
 * Slack通知（オプション）
 * @param {string} message - 送信するメッセージ
 */
function sendSlackNotification(message) {
  if (!SLACK_WEBHOOK_URL || SLACK_WEBHOOK_URL === "") return;

  try {
    const payload = {
      text: message
    };

    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log(`Slack通知失敗: ${e.toString()}`);
  }
}

/**
 * エラーハンドリング
 * @param {string} functionName - 関数名
 * @param {Error} error - エラーオブジェクト
 * @param {Object} context - コンテキスト情報
 */
function handleError(functionName, error, context = {}) {
  const errorMessage = `❌ エラー in ${functionName}: ${error.toString()}`;
  Logger.log(errorMessage);

  if (context.severity === "HIGH") {
    sendSlackNotification(errorMessage);
  }
}
