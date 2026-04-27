// ====================
// Web App エントリーポイント
// ====================

function doGet() {
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
    sheet.appendRow(['id', 'status', 'body', 'image_url', 'scheduled_at', 'posted_at']);
  }
  return sheet;
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
