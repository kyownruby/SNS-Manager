// ====================
// Web App エントリーポイント
// ====================

function doGet(e) {
  // X OAuth 2.0 のコールバック（?code=... 付き）はトークン交換処理へ振り分ける。
  // それ以外は従来どおり Index.html を配信する。
  if (e && e.parameter && e.parameter.code) {
    return handleXAuthCallback(e);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SNS Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====================
// スプレッドシート操作
// ====================

function getSheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    throw new Error('スクリプトプロパティに SPREADSHEET_ID を設定してください');
  }
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName('posts');
  if (!sheet) {
    sheet = ss.insertSheet('posts');
    sheet.appendRow(['id', 'status', 'body', 'image_url', 'scheduled_at', 'posted_at', 'platform', 'error_message']);
    return sheet;
  }
  migrateSheetColumns(sheet);
  return sheet;
}

// 既存シートに platform / error_message 列が無ければヘッダーに追加する（後方互換マイグレーション）。
function migrateSheetColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var required = ['platform', 'error_message'];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(required[i]);
    }
  }
}

function generateId() {
  return Utilities.getUuid();
}

// ====================
// 投稿取得
// ====================

function getPosts() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
      }
      row[headers[j]] = value;
    }
    // platform 列が無い/空の既存データは threads 投稿として扱う（後方互換）。
    if (!row.platform) row.platform = 'threads';
    row._row = i + 1;
    posts.push(row);
  }
  return posts.reverse();
}

// ====================
// 下書き保存
// ====================

function saveDraft(body, imageUrl) {
  var sheet = getSheet();
  var id = generateId();
  sheet.appendRow([id, 'draft', body, imageUrl || '', '', '']);
  return { success: true, id: id };
}

// ====================
// 予約投稿保存
// ====================

function schedulePost(body, imageUrl, scheduledAt) {
  var sheet = getSheet();
  var id = generateId();
  sheet.appendRow([id, 'scheduled', body, imageUrl || '', scheduledAt, '']);
  ensureTrigger();
  return { success: true, id: id };
}

// ====================
// 即時投稿
// ====================

function postNow(body, imageUrl) {
  var sheet = getSheet();
  var id = generateId();
  var result = publishToThreads(body, imageUrl);
  if (result.success) {
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    sheet.appendRow([id, 'posted', body, imageUrl || '', '', now]);
  }
  return result;
}

// ====================
// 投稿編集
// ====================

function updatePost(id, body, imageUrl, scheduledAt, status) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      var row = i + 1;
      sheet.getRange(row, 2).setValue(status || data[i][1]);
      sheet.getRange(row, 3).setValue(body);
      sheet.getRange(row, 4).setValue(imageUrl || '');
      if (scheduledAt) {
        sheet.getRange(row, 5).setValue(scheduledAt);
      }
      if (status === 'scheduled') {
        ensureTrigger();
      }
      return { success: true };
    }
  }
  return { success: false, error: '投稿が見つかりません' };
}

// ====================
// 投稿削除
// ====================

function deletePost(id) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: '投稿が見つかりません' };
}

// ====================
// 画像アップロード（Google Drive）
// ====================

function uploadImage(fileData, fileName, mimeType) {
  var blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);

  var folderId = getOrCreateImageFolder();
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId = file.getId();
  var publicUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
  return { success: true, url: publicUrl, fileId: fileId };
}

function getOrCreateImageFolder() {
  var folderName = 'SNSManager_Images';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  }
  var folder = DriveApp.createFolder(folderName);
  return folder.getId();
}

// ====================
// Threads API
// ====================

function getThreadsConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    accessToken: props.getProperty('THREADS_ACCESS_TOKEN'),
    userId: props.getProperty('THREADS_USER_ID') || 'me'
  };
}

function publishToThreads(body, imageUrls) {
  var config = getThreadsConfig();
  if (!config.accessToken) {
    return { success: false, error: 'THREADS_ACCESS_TOKEN が設定されていません' };
  }

  // imageUrls: カンマ区切り文字列 or 配列 or 空
  var urls = [];
  if (imageUrls) {
    if (typeof imageUrls === 'string') {
      urls = imageUrls.split(',').map(function(u) { return u.trim(); }).filter(function(u) { return u; });
    } else {
      urls = imageUrls;
    }
  }

  var apiBase = 'https://graph.threads.net/v1.0/' + config.userId;

  try {
    var containerId;

    if (urls.length === 0) {
      // テキストのみ
      containerId = createThreadsContainer(apiBase, config.accessToken, {
        text: body,
        media_type: 'TEXT'
      });
    } else if (urls.length === 1) {
      // 画像1枚
      containerId = createThreadsContainer(apiBase, config.accessToken, {
        text: body,
        media_type: 'IMAGE',
        image_url: urls[0]
      });
    } else {
      // カルーセル（2〜4枚）
      var childIds = [];
      for (var i = 0; i < urls.length; i++) {
        var childId = createThreadsContainer(apiBase, config.accessToken, {
          media_type: 'IMAGE',
          image_url: urls[i],
          is_carousel_item: 'true'
        });
        childIds.push(childId);
      }
      // 全子コンテナが FINISHED になるまで待機
      for (var j = 0; j < childIds.length; j++) {
        waitForContainerReady(childIds[j], config.accessToken);
      }
      containerId = createThreadsContainer(apiBase, config.accessToken, {
        text: body,
        media_type: 'CAROUSEL',
        children: childIds.join(',')
      });
    }

    if (!containerId) {
      return { success: false, error: 'コンテナ作成失敗' };
    }

    // 公開
    var publishRes = UrlFetchApp.fetch(apiBase + '/threads_publish', {
      method: 'post',
      payload: {
        creation_id: containerId,
        access_token: config.accessToken
      },
      muteHttpExceptions: true
    });

    var publishData = JSON.parse(publishRes.getContentText());
    if (publishData.id) {
      return { success: true, threadId: publishData.id };
    } else {
      return { success: false, error: '公開失敗: ' + publishRes.getContentText() };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function waitForContainerReady(containerId, accessToken) {
  var maxAttempts = 15;
  for (var i = 0; i < maxAttempts; i++) {
    var res = UrlFetchApp.fetch(
      'https://graph.threads.net/v1.0/' + containerId + '?fields=status&access_token=' + accessToken,
      { muteHttpExceptions: true }
    );
    var data = JSON.parse(res.getContentText());

    if (data.status === 'FINISHED') {
      return;
    }
    if (data.status === 'ERROR' || data.status === 'EXPIRED') {
      throw new Error('コンテナ準備失敗 (ID: ' + containerId + ', status: ' + data.status + ')');
    }
    Utilities.sleep(2000);
  }
  throw new Error('コンテナ準備タイムアウト (ID: ' + containerId + ')');
}

function createThreadsContainer(apiBase, accessToken, params) {
  params.access_token = accessToken;
  var res = UrlFetchApp.fetch(apiBase + '/threads', {
    method: 'post',
    payload: params,
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (!data.id) {
    throw new Error('コンテナ作成失敗: ' + res.getContentText());
  }
  return data.id;
}

// ====================
// 既存投稿を即時公開
// ====================

function publishExistingPost(id, body, imageUrls) {
  var result = publishToThreads(body, imageUrls);
  if (result.success) {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        var row = i + 1;
        sheet.getRange(row, 2).setValue('posted');
        sheet.getRange(row, 3).setValue(body);
        sheet.getRange(row, 4).setValue(imageUrls || '');
        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        sheet.getRange(row, 6).setValue(now);
        break;
      }
    }
  }
  return result;
}

// ====================
// 予約投稿トリガー
// ====================

function checkScheduledPosts() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== 'scheduled') continue;

    var scheduledAt = new Date(data[i][4]);
    if (isNaN(scheduledAt.getTime())) continue;
    if (scheduledAt > now) continue;

    var body = data[i][2];
    var imageUrl = data[i][3];
    var row = i + 1;

    var result = publishToThreads(body, imageUrl);
    if (result.success) {
      var postedAt = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      sheet.getRange(row, 2).setValue('posted');
      sheet.getRange(row, 6).setValue(postedAt);
    } else {
      Logger.log('予約投稿失敗 (row ' + row + '): ' + result.error);
    }
  }
}

function ensureTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkScheduledPosts') {
      return;
    }
  }
  ScriptApp.newTrigger('checkScheduledPosts')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// ====================
// アクセストークン管理
// ====================

function exchangeForLongLivedToken() {
  var props = PropertiesService.getScriptProperties();
  var shortToken = props.getProperty('THREADS_ACCESS_TOKEN');
  var appSecret = props.getProperty('THREADS_APP_SECRET');

  if (!shortToken || !appSecret) {
    return { success: false, error: 'THREADS_ACCESS_TOKEN または THREADS_APP_SECRET が未設定です' };
  }

  try {
    var res = UrlFetchApp.fetch(
      'https://graph.threads.net/access_token'
      + '?grant_type=th_exchange_token'
      + '&client_secret=' + appSecret
      + '&access_token=' + shortToken,
      { muteHttpExceptions: true }
    );

    var data = JSON.parse(res.getContentText());
    if (data.access_token) {
      props.setProperty('THREADS_ACCESS_TOKEN', data.access_token);
      Logger.log('長期トークンに交換しました（有効期限: ' + data.expires_in + '秒）');
      ensureTokenRefreshTrigger();
      return { success: true, expires_in: data.expires_in };
    } else {
      return { success: false, error: 'トークン交換失敗: ' + res.getContentText() };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function refreshLongLivedToken() {
  var props = PropertiesService.getScriptProperties();
  var currentToken = props.getProperty('THREADS_ACCESS_TOKEN');

  if (!currentToken) {
    Logger.log('THREADS_ACCESS_TOKEN が未設定です');
    return { success: false, error: 'THREADS_ACCESS_TOKEN が未設定です' };
  }

  try {
    var res = UrlFetchApp.fetch(
      'https://graph.threads.net/refresh_access_token'
      + '?grant_type=th_refresh_token'
      + '&access_token=' + currentToken,
      { muteHttpExceptions: true }
    );

    var data = JSON.parse(res.getContentText());
    if (data.access_token) {
      props.setProperty('THREADS_ACCESS_TOKEN', data.access_token);
      Logger.log('トークンを更新しました（有効期限: ' + data.expires_in + '秒）');
      return { success: true, expires_in: data.expires_in };
    } else {
      Logger.log('トークン更新失敗: ' + res.getContentText());
      return { success: false, error: 'トークン更新失敗: ' + res.getContentText() };
    }
  } catch (e) {
    Logger.log('トークン更新エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

function ensureTokenRefreshTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshLongLivedToken') {
      return;
    }
  }
  ScriptApp.newTrigger('refreshLongLivedToken')
    .timeBased()
    .everyDays(30)
    .create();
}

// ====================
// 投稿ディスパッチ（共通入口）
// ====================

// platform を見て Threads / X に振り分ける。未指定は従来どおり Threads。
function publishPost(body, imageUrls, platform) {
  if (platform === 'x') {
    return publishToX(body, imageUrls);
  }
  return publishToThreads(body, imageUrls);
}

// 画像URL（カンマ区切り文字列 or 配列 or 空）を配列に正規化する共通ヘルパー。
function normalizeImageUrls(imageUrls) {
  var urls = [];
  if (imageUrls) {
    if (typeof imageUrls === 'string') {
      urls = imageUrls.split(',').map(function(u) { return u.trim(); }).filter(function(u) { return u; });
    } else {
      urls = imageUrls;
    }
  }
  return urls;
}

// ====================
// X OAuth 2.0（PKCE）
// ====================

function getXConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    clientId: props.getProperty('X_CLIENT_ID'),
    clientSecret: props.getProperty('X_CLIENT_SECRET'),
    redirectUri: props.getProperty('X_REDIRECT_URI')
  };
}

// PKCE: code_verifier を生成（UUID 由来の 96 文字。許可文字 [0-9a-f] のみで PKCE 仕様内）。
function generateXCodeVerifier() {
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// PKCE: code_challenge = BASE64URL( SHA-256( code_verifier ) )
function computeXCodeChallenge(verifier) {
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, verifier, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(hash).replace(/=+$/, '');
}

// 認可フローを開始：code_verifier/state を一時保存し、認可 URL を返す。
function startXAuth() {
  var config = getXConfig();
  if (!config.clientId || !config.redirectUri) {
    return { success: false, error: 'X_CLIENT_ID または X_REDIRECT_URI が未設定です' };
  }

  var props = PropertiesService.getScriptProperties();
  var verifier = generateXCodeVerifier();
  var challenge = computeXCodeChallenge(verifier);
  var state = Utilities.getUuid();
  props.setProperty('X_CODE_VERIFIER', verifier);
  props.setProperty('X_OAUTH_STATE', state);

  var scope = 'tweet.read tweet.write users.read media.write offline.access';
  var authUrl = 'https://x.com/i/oauth2/authorize'
    + '?response_type=code'
    + '&client_id=' + encodeURIComponent(config.clientId)
    + '&redirect_uri=' + encodeURIComponent(config.redirectUri)
    + '&scope=' + encodeURIComponent(scope)
    + '&state=' + encodeURIComponent(state)
    + '&code_challenge=' + encodeURIComponent(challenge)
    + '&code_challenge_method=S256';

  return { success: true, authUrl: authUrl };
}

// /2/oauth2/token を叩く共通処理（confidential client は Basic 認証）。
function exchangeXToken(payload) {
  var config = getXConfig();
  var headers = {};
  if (config.clientSecret) {
    headers['Authorization'] = 'Basic ' + Utilities.base64Encode(config.clientId + ':' + config.clientSecret);
  }
  var res = UrlFetchApp.fetch('https://api.x.com/2/oauth2/token', {
    method: 'post',
    headers: headers,
    payload: payload,
    muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

// 取得したトークン情報をスクリプトプロパティに保存（失効時刻も計算して保存）。
function saveXTokens(data) {
  var props = PropertiesService.getScriptProperties();
  if (data.access_token) {
    props.setProperty('X_ACCESS_TOKEN', data.access_token);
  }
  if (data.refresh_token) {
    // X は refresh_token をローテーションするため毎回保存し直す。
    props.setProperty('X_REFRESH_TOKEN', data.refresh_token);
  }
  if (data.expires_in) {
    var expiresAt = Math.floor(Date.now() / 1000) + parseInt(data.expires_in, 10);
    props.setProperty('X_TOKEN_EXPIRES_AT', String(expiresAt));
  }
}

// doGet 経由で受けた認可コードをトークンに交換して保存する。
function handleXAuthCallback(e) {
  var props = PropertiesService.getScriptProperties();
  var config = getXConfig();
  var code = e && e.parameter ? e.parameter.code : null;
  var state = e && e.parameter ? e.parameter.state : null;

  if (!code) {
    return HtmlService.createHtmlOutput('<h2>認証エラー</h2><p>認可コードがありません。</p>');
  }

  var savedState = props.getProperty('X_OAUTH_STATE');
  if (savedState && state !== savedState) {
    return HtmlService.createHtmlOutput('<h2>認証エラー</h2><p>state が一致しません（CSRF の可能性）。</p>');
  }

  var verifier = props.getProperty('X_CODE_VERIFIER');
  if (!verifier) {
    return HtmlService.createHtmlOutput('<h2>認証エラー</h2><p>code_verifier が見つかりません。認可をやり直してください。</p>');
  }

  var data = exchangeXToken({
    grant_type: 'authorization_code',
    code: code,
    redirect_uri: config.redirectUri,
    code_verifier: verifier,
    client_id: config.clientId
  });

  if (data.access_token) {
    saveXTokens(data);
    props.deleteProperty('X_CODE_VERIFIER');
    props.deleteProperty('X_OAUTH_STATE');
    return HtmlService.createHtmlOutput('<h2>X 連携が完了しました！</h2><p>このタブを閉じて、アプリに戻ってください。</p>');
  }

  return HtmlService.createHtmlOutput('<h2>X 連携に失敗しました</h2><pre>' + JSON.stringify(data) + '</pre>');
}

// refresh_token を使ってアクセストークンを再取得する。
function refreshXToken() {
  var props = PropertiesService.getScriptProperties();
  var config = getXConfig();
  var refreshToken = props.getProperty('X_REFRESH_TOKEN');

  if (!refreshToken) {
    return { success: false, error: 'X_REFRESH_TOKEN が未設定です。OAuth 認可を行ってください' };
  }

  var data = exchangeXToken({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: config.clientId
  });

  if (data.access_token) {
    saveXTokens(data);
    return { success: true, expires_in: data.expires_in };
  }
  return { success: false, error: 'トークン更新失敗: ' + JSON.stringify(data) };
}

// 有効なアクセストークンを返す。失効間近なら自動でリフレッシュする共通関数。
function getValidXAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('X_ACCESS_TOKEN');
  var expiresAt = parseInt(props.getProperty('X_TOKEN_EXPIRES_AT') || '0', 10);
  var now = Math.floor(Date.now() / 1000);

  // トークンが無い、または失効 60 秒前なら更新する。
  if (!token || now >= expiresAt - 60) {
    var result = refreshXToken();
    if (!result.success) {
      throw new Error(result.error || 'X アクセストークンの更新に失敗しました');
    }
    token = props.getProperty('X_ACCESS_TOKEN');
  }
  return token;
}

// ====================
// X API v2（投稿・メディア）
// ====================

// 画像 URL からバイナリ（Blob）を取得する。X のメディアアップロード入力に使用。
function fetchImageBytes(imageUrl) {
  var res = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('画像取得失敗 (' + res.getResponseCode() + '): ' + imageUrl);
  }
  return res.getBlob();
}

// メディアアップロード（INIT → APPEND → FINALIZE）。media_id を返す。
function uploadXMedia(blob, token) {
  var base = 'https://api.x.com/2/media/upload';
  var bytes = blob.getBytes();
  var totalBytes = bytes.length;
  var mimeType = blob.getContentType() || 'image/jpeg';
  var authHeader = { 'Authorization': 'Bearer ' + token };

  // INIT
  var initRes = UrlFetchApp.fetch(base, {
    method: 'post',
    headers: authHeader,
    payload: {
      command: 'INIT',
      total_bytes: String(totalBytes),
      media_type: mimeType,
      media_category: 'tweet_image'
    },
    muteHttpExceptions: true
  });
  var initData = JSON.parse(initRes.getContentText());
  var mediaId = (initData.data && initData.data.id) || initData.media_id_string || initData.media_id;
  if (!mediaId) {
    throw new Error('メディア INIT 失敗: ' + initRes.getContentText());
  }

  // APPEND（5MB 以下のチャンクに分割し segment_index を連番で送信）
  var chunkSize = 5 * 1024 * 1024;
  var segmentIndex = 0;
  for (var offset = 0; offset < totalBytes; offset += chunkSize) {
    var end = Math.min(offset + chunkSize, totalBytes);
    var chunkBlob = Utilities.newBlob(bytes.slice(offset, end), mimeType, 'chunk' + segmentIndex);
    var appendRes = UrlFetchApp.fetch(base, {
      method: 'post',
      headers: authHeader,
      payload: {
        command: 'APPEND',
        media_id: mediaId,
        segment_index: String(segmentIndex),
        media: chunkBlob
      },
      muteHttpExceptions: true
    });
    if (appendRes.getResponseCode() >= 300) {
      throw new Error('メディア APPEND 失敗 (seg ' + segmentIndex + '): ' + appendRes.getContentText());
    }
    segmentIndex++;
  }

  // FINALIZE
  var finalizeRes = UrlFetchApp.fetch(base, {
    method: 'post',
    headers: authHeader,
    payload: {
      command: 'FINALIZE',
      media_id: mediaId
    },
    muteHttpExceptions: true
  });
  var finalizeData = JSON.parse(finalizeRes.getContentText());
  var finalId = (finalizeData.data && finalizeData.data.id) || mediaId;

  // 動画/GIF で processing_info がある場合のみ完了までポーリング（静止画は不要）。
  var processingInfo = (finalizeData.data && finalizeData.data.processing_info) || finalizeData.processing_info;
  if (processingInfo) {
    waitForXMediaReady(finalId, token);
  }
  return finalId;
}

// （動画/GIF 用）processing_info が succeeded になるまでポーリング。
function waitForXMediaReady(mediaId, token) {
  var base = 'https://api.x.com/2/media/upload';
  var maxAttempts = 15;
  for (var i = 0; i < maxAttempts; i++) {
    var res = UrlFetchApp.fetch(base + '?command=STATUS&media_id=' + encodeURIComponent(mediaId), {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    var info = (data.data && data.data.processing_info) || data.processing_info;

    if (!info || info.state === 'succeeded') {
      return;
    }
    if (info.state === 'failed') {
      throw new Error('メディア処理失敗: ' + res.getContentText());
    }
    var wait = info.check_after_secs ? info.check_after_secs * 1000 : 2000;
    Utilities.sleep(wait);
  }
  throw new Error('メディア処理タイムアウト (ID: ' + mediaId + ')');
}

// X へ投稿。投稿直前に必ずトークンをリフレッシュ →（画像があれば）メディアアップロード → POST /2/tweets。
function publishToX(body, imageUrls) {
  try {
    // ⚠️ X のアクセストークンは約 2 時間で失効するため、投稿直前に必ず有効化する。
    var token = getValidXAccessToken();

    var urls = normalizeImageUrls(imageUrls);
    var mediaIds = [];
    for (var i = 0; i < urls.length && i < 4; i++) {
      var blob = fetchImageBytes(urls[i]);
      mediaIds.push(uploadXMedia(blob, token));
    }

    var payload = { text: body };
    if (mediaIds.length > 0) {
      payload.media = { media_ids: mediaIds };
    }

    var res = UrlFetchApp.fetch('https://api.x.com/2/tweets', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(res.getContentText());
    if (data.data && data.data.id) {
      return { success: true, tweetId: data.data.id };
    }
    return { success: false, error: 'X 投稿失敗: ' + res.getContentText() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
