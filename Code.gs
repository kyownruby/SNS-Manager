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
  return posts;
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
