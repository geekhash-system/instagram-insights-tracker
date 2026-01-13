// ========================================
// config.gs - Instagram インサイト追跡ツール設定
// ========================================
// 注：秘密情報は .env.gs に記載（GitHubに公開しない）

// Instagram Graph API設定
// 全APIエンドポイントでv21.0を使用（最新メトリクス対応）
const INSTAGRAM_API_BASE = "https://graph.facebook.com/v21.0";

// データ取得設定
const DATA_FETCH_CONFIG = {
  MAX_DAYS_BACK: 30,              // 取得対象期間（30日）
  POSTS_PER_PAGE: 25,             // 1回のAPIコールで取得される投稿数
  MAX_POSTS_PER_ACCOUNT: 100,     // 1アカウントあたりの最大取得投稿数
  API_CALL_DELAY_MS: 500,         // API呼び出し間の待機時間（ミリ秒）
  DAILY_TRIGGER_HOUR: 19          // 毎日19時に実行
};

// 週次ダッシュボード設定
const DASHBOARD_CONFIG = {
  PR_MEDIAN_POSTS_COUNT: 10,      // PR投稿の中央値計算に使う投稿数
  PR_WARNING_THRESHOLD: 0.7       // 警告閾値（中央値の70%以下）
};

// シート名定数
const SHEET_NAMES = {
  NERA: "NERA",
  KARAKO: "KARA子",
  DASHBOARD_NERA: "週次_NERA",
  DASHBOARD_KARAKO: "週次_KARA子",
  README: "はじめに（README）",
  ACCOUNT_INSIGHTS: "アカウントインサイト履歴"
};

// アカウント設定（拡張可能）
const ACCOUNTS = [
  {
    name: "NERA",
    spreadsheetId: "{{NERA_SPREADSHEET_ID}}",  // NERAの専用スプレッドシートID（GitHub Actionsで置換）
    sheetName: SHEET_NAMES.NERA,
    dashboardSheet: SHEET_NAMES.DASHBOARD_NERA,
    businessId: "17841432104009484",
    tokenKey: "NERA_ACCESS_TOKEN"
  },
  {
    name: "KARA子",
    spreadsheetId: "{{KARAKO_SPREADSHEET_ID}}",  // KARA子の専用スプレッドシートID（GitHub Actionsで置換）
    sheetName: SHEET_NAMES.KARAKO,
    dashboardSheet: SHEET_NAMES.DASHBOARD_KARAKO,
    businessId: "17841443872237114",
    tokenKey: "KARAKO_ACCESS_TOKEN"
  }
];

// APIエンドポイント定義（参考: hinome-backend の実装）
const API_ENDPOINTS = {
  // メディア一覧取得（フィード・リール）
  MEDIA: (businessId) => `${INSTAGRAM_API_BASE}/${businessId}/media`,

  // メディアインサイト取得
  MEDIA_INSIGHTS: (mediaId) => `${INSTAGRAM_API_BASE}/${mediaId}/insights`,

  // ストーリー一覧取得
  STORIES: (businessId) => `${INSTAGRAM_API_BASE}/${businessId}/stories`,

  // アカウントインサイト取得
  ACCOUNT_INSIGHTS: (businessId) => `${INSTAGRAM_API_BASE}/${businessId}/insights`,

  // アカウント情報取得（フォロワー数など）
  ACCOUNT_INFO: (businessId) => `${INSTAGRAM_API_BASE}/${businessId}`
};

// メディア取得時のフィールド
const MEDIA_FIELDS = [
  "id",
  "media_type",
  "media_product_type",
  "permalink",
  "caption",
  "timestamp",
  "like_count",
  "comments_count",
  "thumbnail_url",
  "media_url"
].join(",");

// インサイトメトリクス（hinome-backend実装に基づく）
const INSIGHT_METRICS = {
  // リール: viewsを使用（playsではなくviews）
  REELS: "comments,likes,views,reach,saved,shares,total_interactions",

  // フィード: viewsを使用（impressionsではなくviews）
  FEED: "saved,reach,total_interactions,views",

  // ストーリー: viewsを使用
  STORIES: "exits,views,reach,taps_forward,taps_back"
};

// アカウントインサイトメトリクス（日次取得）
const ACCOUNT_INSIGHT_METRICS = {
  DAILY: "reach,accounts_engaged,total_interactions,likes,comments,saved,shares,replies,profile_links_taps",
  PERIOD: "day",
  METRIC_TYPE: "total_value"
};

// カラム定数（スプレッドシート列番号）
const COLUMNS = {
  MEDIA_ID: 0,           // A列
  POST_DATE: 1,          // B列
  DAY_OF_WEEK: 2,        // C列
  POST_TYPE: 3,          // D列
  CAPTION: 4,            // E列
  PERMALINK: 5,          // F列
  PR: 6,                 // G列
  IMP_COUNT: 7,          // H列
  REACH: 8,              // I列
  LIKES: 9,              // J列
  COMMENTS: 10,          // K列
  SAVED: 11,             // L列
  SHARES: 12,            // M列
  ENGAGEMENT: 13,        // N列
  LAST_UPDATE: 14,       // O列
  HISTORY_START: 15      // P列以降
};

// アカウントインサイト履歴シートのカラム定数
const ACCOUNT_INSIGHTS_COLUMNS = {
  DATE: 0,                    // A列: 日付
  FOLLOWER_COUNT: 1,          // B列: フォロワー数
  FOLLOWER_CHANGE: 2,         // C列: フォロワー増減数
  FOLLOWS_COUNT: 3,           // D列: フォロー数
  MEDIA_COUNT: 4,             // E列: 投稿数
  REACH: 5,                   // F列: リーチ数
  ACCOUNTS_ENGAGED: 6,        // G列: エンゲージしたアカウント数
  TOTAL_INTERACTIONS: 7,      // H列: 総エンゲージメント数
  LIKES: 8,                   // I列: いいね数
  COMMENTS: 9,                // J列: コメント数
  SAVED: 10,                  // K列: 保存数
  SHARES: 11,                 // L列: シェア数
  REPLIES: 12,                // M列: 返信数
  PROFILE_LINKS_TAPS: 13      // N列: プロフィールリンクタップ数
};
