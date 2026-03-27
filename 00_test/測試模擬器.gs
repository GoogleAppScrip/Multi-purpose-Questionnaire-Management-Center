function test2() {
  const now = new Date();
  Logger.log(getRecentDay(now,1));
}

/**
 * 測試模擬器 - 用於在不連結 LINE 的情況下測試邏輯
 */
function testIncomingMessage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("模擬測試");
  
  if (!sheet) {
    sheet = ss.insertSheet("模擬測試");
    sheet.getRange("A1:B1").setValues([["模擬輸入 (User Message)", "模擬回覆 (Bot Response)"]]);
    sheet.getRange("A1:B1").setBackground("#cfe2f3").setFontWeight("bold");
    sheet.getRange("A2").setValue("我的善行"); // 預設測試文字
  }
  
  const inputMessage = sheet.getRange("A2").getValue();

  // 模擬 LINE 的事件物件結構
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        events: [{
          type: "message",
          replyToken: "simulated_reply_token",
          source: { userId: "U123456" },
          message: { type: "text", text: inputMessage },
          isSimulation: true // 加入模擬旗標
        }]
      })
    }
  };

  // 執行主程式
  doPost(mockEvent);
  
  SpreadsheetApp.flush();
  Log("模擬完成，請查看「模擬測試」分頁 B2 欄位。");
}

function forceResetAll() {
  PropertiesService.getUserProperties().deleteAllProperties();
  console.log("所有 Session 已清空");
}

/**
 * 列出特定使用者的 Session 資料 (用於測試除錯)
 * @param {String} userId - 使用者 ID (如不填則預設為 U123456)
 */
function listUserSession(userId = "U123456") {
  const session = getSession(userId);
  console.log(`[DEBUG] User Session (${userId}):`, JSON.stringify(session, null, 2));
  return session;
}

/**
 * 刪除特定使用者的 Session 資料
 * @param {String} userId - 使用者 ID (如不填則預設為 U123456)
 */
function deleteUserSession(userId = "U123456") {
  clearSession(userId);
  console.log(`[DEBUG] User Session (${userId}) 已刪除。`);
}

/**
 * 列出所有目前的 Session ID (PropertiesService 中的所有 Key)
 */
function listAllSessionIds() {
  const keys = PropertiesService.getUserProperties().getKeys();
  console.log("[DEBUG] 目前所有的 Session ID:", JSON.stringify(keys, null, 2));
  return keys;
}
