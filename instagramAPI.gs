// ========================================
// instagramAPI.gs - Instagram Graph API連携
// ========================================
// hinome-backend の feed_storage_per_day/main.py を参考

/**
 * 指定アカウントのメディア一覧を取得（フィード・リール）
 * hinome-backend の feed_storage_per_day/main.py を参考
 *
 * @param {string} businessId - Instagram Business Account ID
 * @param {string} accessToken - アクセストークン
 * @param {number} maxDaysBack - 取得対象期間（日数）
 * @return {Array} メディア配列
 */
function fetchMediaList(businessId, accessToken, maxDaysBack = 90) {
  const allMedia = [];
  let nextUrl = API_ENDPOINTS.MEDIA(businessId);
  let apiCallCount = 0;
  const maxApiCalls = 50; // 安全のため上限設定

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - maxDaysBack);

  try {
    while (nextUrl && apiCallCount < maxApiCalls) {
      apiCallCount++;

      const params = {
        access_token: accessToken,
        fields: MEDIA_FIELDS,
        limit: DATA_FETCH_CONFIG.POSTS_PER_PAGE
      };

      const url = nextUrl + (nextUrl.includes('?') ? '&' : '?') +
                  Object.keys(params).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');

      Logger.log(`📥 メディア取得 API呼び出し #${apiCallCount}`);

      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true
      });

      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (statusCode !== 200) {
        Logger.log(`❌ API エラー (${statusCode}): ${responseText}`);
        break;
      }

      const result = JSON.parse(responseText);
      const mediaData = result.data || [];

      // 日付フィルタリング
      for (let i = 0; i < mediaData.length; i++) {
        const media = mediaData[i];
        const postDate = new Date(media.timestamp);

        if (postDate < cutoffDate) {
          Logger.log(`📊 ${maxDaysBack}日以前の投稿に到達: ${postDate.toISOString()}`);
          nextUrl = null; // これ以上取得しない
          break;
        }

        allMedia.push(media);
      }

      // ページネーション
      if (result.paging && result.paging.next && nextUrl) {
        nextUrl = result.paging.next;
        Utilities.sleep(DATA_FETCH_CONFIG.API_CALL_DELAY_MS);
      } else {
        nextUrl = null;
      }

      Logger.log(`✅ ${mediaData.length} 件取得 (累計: ${allMedia.length} 件)`);
    }

    Logger.log(`🎉 メディア取得完了: 合計 ${allMedia.length} 件`);
    return allMedia;

  } catch (e) {
    Logger.log(`エラー in fetchMediaList: ${e.toString()}`);
    return allMedia;
  }
}

/**
 * メディアのインサイトを取得
 * hinome-backend の feed_storage_per_day/main.py (Line 348-360) を参考
 *
 * @param {string} mediaId - メディアID
 * @param {string} mediaProductType - REELS or FEED
 * @param {string} accessToken - アクセストークン
 * @return {Object} インサイトデータ（nullの場合は取得失敗）
 */
function fetchMediaInsights(mediaId, mediaProductType, accessToken) {
  try {
    const metrics = (mediaProductType === "REELS") ?
                    INSIGHT_METRICS.REELS :
                    INSIGHT_METRICS.FEED;

    const url = `${API_ENDPOINTS.MEDIA_INSIGHTS(mediaId)}?metric=${metrics}&access_token=${accessToken}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      Logger.log(`❌ インサイト取得エラー (${statusCode}): ${mediaId}`);
      return null;
    }

    const result = JSON.parse(response.getContentText());
    const insights = {};

    // データを整形（hinome-backend実装に基づく）
    // APIレスポンス形式: {data: [{name: "views", values: [{value: 12345}]}]}
    if (result.data) {
      result.data.forEach(item => {
        insights[item.name] = item.values && item.values[0] ? item.values[0].value : 0;
      });
    }

    return insights;

  } catch (e) {
    Logger.log(`エラー in fetchMediaInsights: ${e.toString()}`);
    return null;
  }
}

/**
 * ストーリー一覧を取得
 * hinome-backend の story_storage_per_hour/main.py (Line 56-78) を参考
 *
 * @param {string} businessId - Instagram Business Account ID
 * @param {string} accessToken - アクセストークン
 * @return {Array} ストーリー配列
 */
function fetchStories(businessId, accessToken) {
  const allStories = [];
  let nextUrl = API_ENDPOINTS.STORIES(businessId);

  try {
    while (nextUrl) {
      const params = {
        access_token: accessToken,
        fields: "caption,id,like_count,media_product_type,media_type,media_url,permalink,thumbnail_url,timestamp"
      };

      const url = nextUrl + (nextUrl.includes('?') ? '&' : '?') +
                  Object.keys(params).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');

      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true
      });

      const statusCode = response.getResponseCode();
      if (statusCode !== 200) {
        Logger.log(`❌ ストーリー取得エラー (${statusCode})`);
        break;
      }

      const result = JSON.parse(response.getContentText());

      if (result.data && result.data.length > 0) {
        allStories.push(...result.data);
      }

      // ページネーション
      if (result.paging && result.paging.next) {
        nextUrl = result.paging.next;
        Utilities.sleep(DATA_FETCH_CONFIG.API_CALL_DELAY_MS);
      } else {
        nextUrl = false;
      }
    }

    Logger.log(`📖 ストーリー取得完了: ${allStories.length} 件`);
    return allStories;

  } catch (e) {
    Logger.log(`エラー in fetchStories: ${e.toString()}`);
    return allStories;
  }
}

/**
 * ストーリーのインサイトを取得
 * @param {string} storyId - ストーリーID
 * @param {string} accessToken - アクセストークン
 * @return {Object} インサイトデータ
 */
function fetchStoryInsights(storyId, accessToken) {
  try {
    const metrics = INSIGHT_METRICS.STORIES;
    const url = `${API_ENDPOINTS.MEDIA_INSIGHTS(storyId)}?metric=${metrics}&access_token=${accessToken}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      Logger.log(`❌ ストーリーインサイト取得エラー (${statusCode}): ${storyId}`);
      return null;
    }

    const result = JSON.parse(response.getContentText());
    const insights = {};

    if (result.data) {
      result.data.forEach(item => {
        insights[item.name] = item.values && item.values[0] ? item.values[0].value : 0;
      });
    }

    return insights;

  } catch (e) {
    Logger.log(`エラー in fetchStoryInsights: ${e.toString()}`);
    return null;
  }
}

/**
 * アカウント情報を取得（フォロワー数）
 * @param {string} businessId - Instagram Business Account ID
 * @param {string} accessToken - アクセストークン
 * @return {Object} アカウント情報
 */
function fetchAccountInfo(businessId, accessToken) {
  try {
    const url = `${API_ENDPOINTS.ACCOUNT_INFO(businessId)}?fields=followers_count&access_token=${accessToken}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      Logger.log(`❌ アカウント情報取得エラー (${statusCode})`);
      return null;
    }

    const result = JSON.parse(response.getContentText());
    return { followers_count: result.followers_count || 0 };
  } catch (e) {
    Logger.log(`エラー in fetchAccountInfo: ${e.toString()}`);
    return null;
  }
}

/**
 * アカウントインサイトを取得
 * @param {string} businessId - Instagram Business Account ID
 * @param {string} accessToken - アクセストークン
 * @param {string} since - 開始日時（Unixタイムスタンプ）
 * @param {string} until - 終了日時（Unixタイムスタンプ）
 * @return {Object} アカウントインサイトデータ
 */
function fetchAccountInsights(businessId, accessToken, since, until) {
  try {
    const metrics = ACCOUNT_INSIGHT_METRICS.DAILY;
    const period = ACCOUNT_INSIGHT_METRICS.PERIOD;
    const metricType = ACCOUNT_INSIGHT_METRICS.METRIC_TYPE;

    const url = `${API_ENDPOINTS.ACCOUNT_INSIGHTS(businessId)}?metric=${metrics}&period=${period}&metric_type=${metricType}&since=${since}&until=${until}&access_token=${accessToken}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      Logger.log(`❌ アカウントインサイト取得エラー (${statusCode})`);
      return null;
    }

    const result = JSON.parse(response.getContentText());
    const insights = {};

    // Parse response - extract total_value from each metric
    if (result.data) {
      result.data.forEach(item => {
        if (item.total_value && item.total_value.value !== undefined) {
          insights[item.name] = item.total_value.value;
        }
      });
    }

    return insights;
  } catch (e) {
    Logger.log(`エラー in fetchAccountInsights: ${e.toString()}`);
    return null;
  }
}
