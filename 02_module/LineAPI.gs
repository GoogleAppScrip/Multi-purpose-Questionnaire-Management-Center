/**
 * 透過 LINE Messaging API 回覆訊息
 * @param {String} replyToken - 回覆用的憑證
 * @param {String} message - 要回送的文本
 * @param {Boolean} isSimulation - 是否為模擬模式 (寫回試算表 B2)
 */
function replyLINE(replyToken, message, isSimulation = false) {
  if (isSimulation) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("模擬測試");
    if (sheet) {
      sheet.getRange("B2").setValue(message);
    }
    Log("模擬模式：訊息已寫回『模擬測試』分頁 B2。內容：" + message);
    return;
  }

  const url = "https://api.line.me/v2/bot/message/reply";
  const accessToken = ss.getSheetByName("參數設定").getRange("B2").getValue();
  UrlFetchApp.fetch(url, {
    "headers": {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + accessToken,
    },
    "method": "post",
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": [{ "type": "text", "text": message }]
    }),
  });
}

/**
 * 獲取使用者個人檔案 (Profile)
 * 呼叫對象：業務邏輯函式 (如 processEngine) 或事件處理函式。
 * 用途：取得使用者顯示名稱、頭像、狀態。
 * @param {String} userId - 使用者 ID
 * @returns {Object} 包含 displayName, pictureUrl, statusMessage 的物件
 */
function getProfile(userId) {
  try {
    Log(`[DEBUG] getProfile() TODO: 實作 getProfile API 呼叫"`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = "https://api.line.me/v2/bot/profile/" + userId;
    const accessToken = ss.getSheetByName("參數設定").getRange("B2").getValue();
    const response = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + accessToken },
      "method": "get"
    });
    return JSON.parse(response.getContentText());
  } catch (err) {
    Log(`[ERROR] getProfile 失敗: ${err.message}`);
    return null;
  }
}

/**
 * 獲取群組摘要 (Summary)
 * 呼叫對象：業務邏輯函式或事件處理函式。
 * 用途：取得群組名稱、圖示。
 * @param {String} groupId - 群組 ID
 * @returns {Object} 包含 groupName, pictureUrl 的物件
 */
function getGroupSummary(groupId) {
  try {
    Log(`[DEBUG] getGroupSummary() 實作 getGroupSummary API 呼叫"`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = "https://api.line.me/v2/bot/group/" + groupId + "/summary";
    const accessToken = ss.getSheetByName("參數設定").getRange("B2").getValue();
    const response = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + accessToken },
      "method": "get"
    });

    const info = JSON.parse(response.getContentText());
    Log(`[DEBUG] group information: ${info}`);

    return info;

  } catch (err) {
    Log(`[ERROR] getGroupSummary 失敗: ${err.message}`);
    return null;
  }
}

/**
 * 獲取群組成員人數
 * 呼叫對象：業務邏輯函式或事件處理函式。
 * 用途：取得群組內人數。
 * @param {String} groupId - 群組 ID
 * @returns {Object} 包含 count 的物件
 */
function getGroupMembersCount(groupId) {
  try {
    Log(`[DEBUG] getGroupMembersCount() TODO: 實作 getGroupMembersCount API 呼叫"`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = "https://api.line.me/v2/bot/group/" + groupId + "/members/count";
    const accessToken = ss.getSheetByName("參數設定").getRange("B2").getValue();
    const response = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + accessToken },
      "method": "get"
    });
    return JSON.parse(response.getContentText());
  } catch (err) {
    Log(`[ERROR] getGroupMembersCount 失敗: ${err.message}`);
    return null;
  }
}

/**
 * 獲取特定群組成員的個人檔案
 * 呼叫對象：業務邏輯函式或事件處理函式。
 * 用途：取得群組中特定成員的資訊。
 * @param {String} groupId - 群組 ID
 * @param {String} userId - 使用者 ID
 * @returns {Object} 包含 displayName, pictureUrl 的物件
 */
function getGroupMemberProfile(groupId, userId) {
  try {
    Log(`[DEBUG] getGroupMemberProfile() TODO: 實作 getGroupMemberProfile API 呼叫"`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = "https://api.line.me/v2/bot/group/" + groupId + "/member/" + userId;
    const accessToken = ss.getSheetByName("參數設定").getRange("B2").getValue();
    const response = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + accessToken },
      "method": "get"
    });
    return JSON.parse(response.getContentText());
  } catch (err) {
    Log(`[ERROR] getGroupMemberProfile 失敗: ${err.message}`);
    return null;
  }
}


/*
 * 其他常用的 LINE Messaging API 清單 (供參考)：
 * 
 * --- 訊息發送 ---
 * - pushMessage: 主動對特定用戶發送訊息 (需要 Channel Access Token)
 * - multicastMessage: 對多個特定用戶發送相同訊息
 * - broadcastMessage: 對所有加入好友的用戶發送訊息
 * 
 * --- 使用者/群組資訊 ---
 * - getProfile: 獲取用戶的名稱、頭像、身分證言 (Status Message)
 * - getGroupSummary: 獲取群組名稱、圖示
 * - getGroupMembersCount: 獲取群組內的人數
 * - getGroupMemberProfile: 獲取特定的群組成員資訊
 * 
 * --- 其他操作 ---
 * - leaveGroup: 令機器人主動退出群組
 * - richmenu: 管理與設定圖文選單
 * - insight: 獲取統計資料 (如訊息送達數、好友增加數等)
 */
