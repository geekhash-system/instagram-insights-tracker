// ========================================
// config.gs - Instagram インサイト追跡ツール設定
// ========================================
// 注：秘密情報は .env.gs に記載（GitHubに公開しない）

// スプレッドシート ID
const SPREADSHEET_ID = "1mi10KVyRkf8_svopr2Hlh4a5LJxW4QO9fSTx7EjzBGQ";

// Instagram Graph API設定
// hinome-backendではv18.0とv21.0を使い分けている
// メディア・インサイト取得: v18.0（安定版）
// アカウントインサイト: v21.0（新メトリクス対応）
const INSTAGRAM_API_BASE = "https://graph.facebook.com/v18.0";

// データ取得設定
const DATA_FETCH_CONFIG = {
  MAX_DAYS_BACK: 90,              // 取得対象期間（90日）
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
  NERA: "NERA_2512~",
  KARAKO: "KARA子_2512~",
  DASHBOARD_NERA: "週次_NERA",
  DASHBOARD_KARAKO: "週次_KARA子",
  README: "はじめに（README）"
};

// アカウント設定（拡張可能）
const ACCOUNTS = [
  {
    name: "NERA",
    sheetName: SHEET_NAMES.NERA,
    dashboardSheet: SHEET_NAMES.DASHBOARD_NERA,
    businessId: "17841432104009484",
    tokenKey: "NERA_ACCESS_TOKEN"
  },
  {
    name: "KARA子",
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
  ACCOUNT_INSIGHTS: (businessId) => `${INSTAGRAM_API_BASE}/${businessId}/insights`
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

// カラム定数（スプレッドシート列番号）
const COLUMNS = {
  MEDIA_ID: 0,           // A列
  POST_DATE: 1,          // B列
  POST_TYPE: 2,          // C列
  CAPTION: 3,            // D列
  PERMALINK: 4,          // E列
  PR: 5,                 // F列
  IMP_COUNT: 6,          // G列
  REACH: 7,              // H列
  LIKES: 8,              // I列
  COMMENTS: 9,           // J列
  SAVED: 10,             // K列
  SHARES: 11,            // L列
  ENGAGEMENT: 12,        // M列
  LAST_UPDATE: 13,       // N列
  HISTORY_START: 14      // O列以降
};
