/**
 * 萬用問卷與統計引擎 (Universal Questionnaire Engine)
 * 實作規格：萬用問卷設計規格書 v1.0
 */

const ss = SpreadsheetApp.getActiveSpreadsheet();
const projectIndex = ss.getSheetByName("專案索引");
const log_sheet = ss.getSheetByName("紀錄資料");

const cache = PropertiesService.getUserProperties(); // 現階段以 Session 記憶為主

/**
 * LINE Webhook 入口
 * 處理來自 LINE 平台的消息事件，並根據邏輯進行回覆
 * @param {Object} e - Apps Script 傳入的事件對象
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.events[0];
    if (!event) return;

    // --- 事件分流處理 ---
    const type = event.type;
    if (type === "follow") return handleFollow(event);
    if (type === "unfollow") return handleUnfollow(event);
    if (type === "join") return handleJoin(event);
    if (type === "leave") return handleLeave(event);
    if (type === "message") return handleMessage(event);

  } catch (err) {
    Log(`[ERROR] doPost 執行異常: ${err.message}`);
    console.error("系統錯誤: " + err.message);
  }
}


/**
 * 處理「加入好友 (Follow)」事件
 */
function handleFollow(event) {
  Log(`[DEBUG] handleFollow(): 處理新增好友邏輯"`);

  // 1. 取得使用者ID (userId)
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  // 2. 檢查使用者是否已經存在(userId)於 [使用者資料]分頁裡
  const userName = getUserName(userId);
  if (userName !== "同行善友") {
    Log(`[INFO] 使用者 ${userId} (${userName}) 已存在，跳過加入好友問卷。`);
    return;
  }

  // 3. 取得使用者詳細資料 資料填寫到變數裡 "${用戶編號} ${用戶名稱} ${用戶圖像} ${用戶狀態}"
  const profile = getProfile(userId) || {};
  let session = getSession(userId);
  if (!session.data) session.data = {};

  session.data["用戶編號"] = userId;
  session.data["用戶名稱"] = profile.displayName || "匿名";
  session.data["用戶圖像"] = profile.pictureUrl || "";
  session.data["用戶狀態"] = profile.statusMessage || "";

  // 4. 觸發 "專案索引" 裡面的 "加入好友"之問卷
  const responseText = processEngine(userId, "加入好友", session);
  
  if (responseText) {
    replyLINE(replyToken, responseText);
  } else {
    Log(`[WARN] handleFollow: 未觸發問卷或 responseText 為空`);
  }
}

/**
 * 處理「取消好友 (Unfollow)」事件
 */
function handleUnfollow(event) {
  Log(`[DEBUG] handleUnfollow() TODO: 處理封鎖/刪除好友邏輯"`);
}

/**
 * 處理「加入群組 (Join)」事件
 */
function handleJoin(event) {
  Log(`[DEBUG] handleJoin(): 處理加入群組邏輯"`);

  // 1. 取得群組ID (groupId)
  const groupId = event.source.groupId || event.source.roomId;
  const replyToken = event.replyToken;
  if (!groupId) return;

  // 2. 檢查群組ID是否已經存在於 [使用者資料]分頁裡的"用戶名稱"
  const groupNameCheck = getUserName(groupId);
  if (groupNameCheck !== "同行善友") {
    Log(`[INFO] 群組 ${groupId} (${groupNameCheck}) 已存在，跳過加入群組問卷。`);
    return;
  }

  // 3. 取得群組詳細資料 資料填寫到變數裡 "${用戶編號} ${用戶名稱} ${用戶圖像}"
  const summary = getGroupSummary(groupId) || {};
  let session = getSession(groupId);
  if (!session.data) session.data = {};

  session.data["用戶編號"] = groupId;
  session.data["用戶名稱"] = summary.groupName || "未知群組";
  session.data["用戶圖像"] = summary.pictureUrl || "";

  // 4. 觸發 "專案索引" 裡面的 "加入群組"之問卷
  const responseText = processEngine(groupId, "加入群組", session);
  
  if (responseText) {
    replyLINE(replyToken, responseText);
  }
}

/**
 * 處理「離開群組 (Leave)」事件
 */
function handleLeave(event) {
  Log(`[DEBUG] handleLeave() TODO: 處理機器人被踢出或群組廢止邏輯"`);
}


function handleMessage(event) {
  try {
    const userId = event.source.userId;
    const userMessage = event.message.text;
    const replyToken = event.replyToken;

    // 1. 取得使用者 Session 狀態
    let session = getSession(userId);
    Log(`[RECEIVE] userId: ${userId}, message: ${userMessage}, session: ${JSON.stringify(session)}`);

    // 2. 引擎處理邏輯
    const responseText = processEngine(userId, userMessage, session);
    Log(`[DEBUG] Final responseText: "${responseText}"`);
    
    // 3. 回覆訊息
    if (responseText) {
      const isSimulation = (event.parameter && event.parameter.isSimulation);
      replyLINE(replyToken, responseText, isSimulation);
    } else {
      Log(`[WARN] responseText 為空，未執行 replyLINE (請確認關鍵字是否匹配或狀態是否正確)`);
    }
  } catch (err) {
    Log(`[ERROR] handleMessage 執行異常: ${err.message}`);
    console.error("系統錯誤: " + err.message);
  }
}

/**
 * 核心引擎處理器
 * 負責識別專案關鍵字、控制問卷狀態流轉與管理資料儲存
 * @param {String} userId - 使用者的獨一 ID
 * @param {String} userInput - 使用者的輸入文本
 * @param {Object} session - 目前的使用者 Session 對象
 * @param {Number} depth - 遞迴深度 (內部使用，防止配置錯誤導致死循環)
 * @returns {String} render後的系統回覆訊息
 */
function processEngine(userId, userInput, session, depth = 1) {
    Log(`Debug 1> userId: ${userId}, message: ${userInput}, session: ${JSON.stringify(session)}`);
  // 1. 安全閥：遞迴深度限制，防止配置錯誤導致死循環
  if (depth >= 15) {
    return "系統遞迴執行超過上限，請檢查問卷邏輯是否有死循環。";
  }

  // 2. 專案識別與初始化 (僅在無專案且首層時執行)
  if (!session.projectId) {
    // 使用專案識別工具檢查使用者輸入是否觸發特定問卷
    const project = checkProjectTrigger(userInput);
    if (!project) return null; // 未觸發任何專案

    // 確保 session.data 存在
    if (!session.data) session.data = {};

    // 取得基本系統變數，但不覆蓋已存在的 (例如 handleFollow 預填的資料)
    const sysData = { 
      "UserName": getUserName(userId), 
      "Group": getUserGroup(userId)
    };
    for (let key in sysData) {
      if (session.data[key] === undefined) session.data[key] = sysData[key];
    }

    // 根據識別結果填入 Session 初始資料
    session.projectId = renderTemplate(project.description, session); // 專案描述或標題
    session.confSheet = renderTemplate(project.projectName, session); // 配置表分頁
    session.resultSheet = renderTemplate(project.resultSheet, session); // 結果儲存分頁
    session.currentState = "START";
    session.data["問卷名稱"] = session.confSheet;
    session.data["儲存分頁"] = session.resultSheet;
    session.data["USER_INPUT"] = project.matchedInput; // 已去關鍵字後的輸入
    Log(`[INIT] Project: ${session.projectId}, Config: ${session.confSheet}, Result: ${session.resultSheet}`);
  }

    Log(`Debug 2> userId: ${userId}, message: ${userInput}, session: ${JSON.stringify(session)}`);

  let responseText = "";
  // 3. 獲取當前狀態之顯示內容與處理邏輯
  Log(`[FLOW] Current State: ${session.currentState} (Depth: ${depth})`);
  const config = getConfig(session.confSheet, session.currentState);

  // 5. 流程分支與特殊狀態處理
  // 如果狀態是 "SAVE" 則要另外處理 (進行資料永久寫入)
  if (session.currentState === "SAVE") {

    // 1. 直接讀取 "輸入"欄位裡面的值, 作為存入到結果分頁的 矩陣
    const inputTemplates = convertToArray(renderTemplate(config.Input, session));
    const dataRow = inputTemplates.map(tpl => renderTemplate(String(tpl), session));

    // 2. 從 session.resultSheet 取得分頁名稱, 建議再做一次轉換 renderTemplate
    const targetSheetName = renderTemplate(session.resultSheet, session);
    let targetSheet = ss.getSheetByName(targetSheetName);

    // 3. 如果分頁不存在, 則建立新的分頁, 並且在第一列裡面, 填入 "變數"欄位, 當作 title
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetSheetName);
      const titleRow = convertToArray(config.Storage);
      if(titleRow.length > 0) {
        targetSheet.getRange(1, 1, 1, titleRow.length).setValues([titleRow]);
        targetSheet.getRange(1, 1, 1, titleRow.length).setBackground("#d0e0e3").setFontWeight("bold");
      }
    }

    // 4. 將結果存入 分頁的 第二列裡面, 原本的第二列會被擠到下面
    targetSheet.insertRowBefore(2);
    targetSheet.getRange(2, 1, 1, dataRow.length).setValues([dataRow]);
    SpreadsheetApp.flush(); // 強制更新

  } else {
    Log(`[EXEC] Handling Input & Process for State: ${session.currentState}`);
    handleInputConfig(config, userInput, session);
    handleProcessConfig(config, userInput, session);

    // 處理輸出模板渲染
    responseText = handleOutputConfig(config, session);
    Log(`[EXEC] Output Rendered: "${responseText}"`);
    
    // 暫存目前的 Session 狀態 (確保 handle*Config 變動已被紀錄)
    saveSession(userId, session);
  }

  // 4. 異常與錯誤提示處理
  // 如果 handleOutputConfig 傳回的是 null 則去讀取 "錯誤提示"，並且經過 renderTemplate 轉換
  if (responseText === null) {
    // 暫存目前的 Session 狀態 (確保 handle*Config 變動已被紀錄)
    saveSession(userId, session);

    responseText = renderTemplate(config.Error, session);
    // 發生錯誤時回傳提示訊息，但不進行狀態轉換，讓使用者重新嘗試
    return responseText;
  }

  // 6. END
  if (session.currentState === "END") {
    // TODO
    if (session.undefine && session.undefine.length > 0)
      responseText = renderTemplate(responseText, session);

    clearSession(userId); // 結束流程清除記憶
    return responseText;
  }
  
  // 7. 則轉換到下一個狀態
  const nextState = getNextState(config, session);
  Log(`[FLOW] State Transition: ${session.currentState} -> ${nextState}`);
  session.currentState = nextState;

  if (session.currentState === "ERROR") {
    clearSession(userId); // 結束流程清除記憶
    return "系統內部錯誤, 請洽工程師";
  }

  // 8. 遞迴呼叫 (處理靜態過渡狀態)
  // 如果輸出為空字串，代表這是內部過渡狀態，直接進入下一輪遞迴
  if (responseText === "") {
    return processEngine(userId, userInput, session, depth + 1);
  }

  // 9. 最終更新 Session 並回傳渲染後的文本
  saveSession(userId, session);
  Log(`Debug 3> userId: ${userId}, message: ${userInput}, session: ${JSON.stringify(session)}`);

  if (session.undefine && session.undefine.length > 0)
    responseText = renderTemplate(responseText, session);

  return responseText;
}



// --- 底層輔助工具 ---

/**
 * 自定義日誌函數，將訊息與呼叫端資訊寫入「紀錄資料」分頁
 * 包含：時間戳記、檔案名稱、函式名稱、行號與自定義訊息
 * @param {String} message - 要記錄的訊息內容
 */
function Log(message) {
  try {
    const stack = new Error().stack.split("\n");
    // 取得呼叫者的資訊 (通常在 stack 的第三行)
    const callerInfo = stack.length > 2 ? stack[2].trim() : "unknown";

    const match = callerInfo.match(/at (\S+) \((.*):(\d+):(\d+)\)/) || callerInfo.match(/at (.*):(\d+):(\d+)/);
    let functionName = match?.[1] || "unknown";
    let fileName = match?.[2] || match?.[1] || "unknown";
    let lineNumber = match?.[3] || "unknown";

    log_sheet.insertRows(2);
    const timeStamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
    log_sheet.getRange(2, 1, 1, 5).setValues([[timeStamp, fileName, functionName, lineNumber, message]]);
    Logger.log(message);
  } catch (err) {
    Logger.log(`Log 函數錯誤: ${err.message}`);
  }
}

/**
 * 從 Cache (PropertiesService) 中獲取使用者的 Session 狀態
 * @param {String} userId - 使用者 ID
 * @returns {Object} 解析後的 Session 對象
 */
function getSession(userId) {
  const data = cache.getProperty(userId);
  if (data) {
    const parsed = JSON.parse(data);
    parsed.userId = userId; // 補回 ID
    return parsed;
  }
  return { userId: userId, data: {} };
}

/**
 * 檢查使用者輸入是否觸發特定問卷專案
 * 讀取「專案索引」分頁中的配置：觸發關鍵字(0), 問卷名稱(1), 儲存分頁(2), 專案描述(3)
 * @param {String} userInput - 使用者輸入文本
  * @returns {Object|null} 匹配到的專案資料物件，未觸發則回傳 null
 */
function checkProjectTrigger(userInput) {
  const projects = projectIndex.getDataRange().getValues();
  
  // 跳過第一行標題
  for (let i = 1; i < projects.length; i++) {
    const keyword = projects[i][0];
    const projectName = projects[i][1];
    const resultSheet = projects[i][2];
    const description = projects[i][3];

    // 使用從試算表讀取的關鍵字（支援正則表達式）來測試使用者目前的輸入內容
    // 如果 userInput 匹配成功（例如輸入了「我的善行」觸發了「^我的善行」），則進入執行邏輯
    const pattern = new RegExp(keyword);
    if (pattern.test(userInput)) {
      // 回傳識別結果供呼叫者使用
      return {
        keyword: keyword,
        projectName: projectName,
        resultSheet: resultSheet,
        description: description,
        matchedInput: userInput.replace(pattern, '').trim()
      };
    }
  }
  return null;
}

/**
 * 根據分頁名稱與狀態 ID 獲取狀態標籤配置 (狀態機核心配置)
 * 配置欄位順序：狀態編號(0), 輸入(1), 處理(2), 變數(3), 輸出(4), 下一狀態(5), 錯誤提示(6)
 * @param {String} sheetName - 配置表分頁名稱
 * @param {String} stateId - 當前的狀態 ID
 * @returns {Object} 包含輸入預設、處理指令、儲存變數、輸出文本、下一狀態路徑與錯誤訊息的配置對象
 */
function getConfig(sheetName, stateId) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { Output: "錯誤：找不到分頁 " + sheetName };
  
  const range = sheet.getDataRange();
  const rows = range.getValues();      // 獲取所有儲存格值

  // 遍歷所有行，跳過標題 (i=1)
  // 欄位對應：[0]狀態編號, [1]輸入, [2]處理, [3]變數, [4]輸出, [5]下一狀態, [6]錯誤提示
  const searchId = String(stateId).trim();
  for (let i = 1; i < rows.length; i++) {
    const rowId = String(rows[i][0]).trim();
    if (rowId === searchId) {
      return {
        Input: rows[i][1],     // 狀態進入後的初始輸入值
        Process: rows[i][2],   // {} 格式的處理指令 (JSON 映射)
        Storage: normalizeStorageKey(rows[i][3]),   // 儲存到 Session 的變數名稱
        Output: rows[i][4],    // 系統回覆的文本模板
        NextState: rows[i][5], // 下一狀態 ID 或轉換條件
        Error: rows[i][6],     // 驗證失敗時的提示訊息
      };
    }
  }
  Log(`[ERROR] getConfig 找不到狀態: "${searchId}" (正在搜尋分頁: ${sheetName})`);
  return { Output: "錯誤：找不到狀態配置 " + searchId };
}

/**
 * 處理配置表中的「輸入」欄位邏輯
 * 1. 從 getConfig 取得完整資料
 * 2. 根據「輸入」欄位的值做處理
 * 3. 屬性為 "<USER INPUT>"：變數 = userInput
 * 4. 屬性為空白：不做任何處理
 * 5. 剩下的：原封不動填入變數裡面
 * @param {Object} config - 狀態配置對象
 * @param {String} userInput - 使用者輸入
 * @param {Object} session - 目前 Session (會更新 session.data)
 */
function handleInputConfig(config, userInput, session) {
  const inputVal = config.Input;
  const storageKey = config.Storage;
  
  // 1. 如果「輸入」值為空白，不做任何處理
  if (inputVal === "" || inputVal === undefined || inputVal === null) {
    return;
  }
  
  // 2. 判斷是否具備有效的儲存變數名稱
  if (!storageKey || storageKey === "(null)") {
    return;
  }

  // 3. 執行處理邏輯
  if (inputVal === "<USER INPUT>") {
    // 如果是 "<USER INPUT>" 則 變數 = userInput
    session.data[storageKey] = userInput;
  } else {
    // 剩下的 原封不動的 填入到 變數 裡面
    session.data[storageKey] = renderTemplate(inputVal, session);
  }
}

/**
 * 處理配置表中的「處理」欄位邏輯
 * 1. 解析 JSON 指令 (例如 {"B23":"<USER INPUT>","OUTPUT":"C23"})
 * 2. 如果值為 "<USER INPUT>"，則將 userInput 填入該 key 指定的座標
 * 3. 如果 key 為 "OUTPUT"，則讀取該 value (座標) 的值並存入 session.data[config.Storage]
 * 4. 否則將渲染後的 value 填入該 key 指定的座標
 * @param {Object} config - 狀態配置對象
 * @param {String} userInput - 使用者輸入
 * @param {Object} session - 目前 Session
 */
function handleProcessConfig(config, userInput, session) {
  const processStr = config.Process;
  if (!processStr || processStr === "" || processStr === "(null)") return;

  try {
    const processMap = JSON.parse(processStr);
    const sheet = ss.getSheetByName(session.confSheet);
    
    for (let key in processMap) {
      const value = processMap[key];

      if(key === "OUTPUT") {
        // 從座標讀取值並存入 session.data[config.Storage]
        const cellValue = sheet.getRange(value).getValue();
        session.data[config.Storage] = renderTemplate(String(cellValue), session);
      } else if (value === "<USER INPUT>") {
        // 將 userInput 填入座標
        sheet.getRange(key).setValue(userInput);
        SpreadsheetApp.flush(); // 強制更新以利後續讀取公式結果
      } else {
        // 將渲染後的內容寫入指定座標
        sheet.getRange(key).setValue(renderTemplate(String(value), session));
      }
    }
  } catch (e) {
    Log(`解析「處理」欄位失敗: ${e.message}, 內容: ${processStr}`);
  }
}

/**
 * 處理配置表中的「輸出」欄位邏輯
 * 1. 檢查 "輸出" 是否有值，沒有則返回空字串
 * 2. 如果是 JSON 物件，則根據 "變數" (config.Storage) 的值當做 key 取得字串
 * 3. 若 JSON 映射找不到對應內容且無預設值，則返回 null
 * 4. 最後將字串透過 renderTemplate 轉換後傳回
 * @param {Object} config - 狀態配置對象
 * @param {Object} session - 目前 Session
 * @returns {String|null} 渲染後的輸出內容，找不到配置則返回 null
 */
function handleOutputConfig(config, session) {
  let outputTemplate = config.Output;
  const storageKey = config.Storage;

  // 1. 檢查是否有值
  if (!outputTemplate || outputTemplate === "" || outputTemplate === "(null)") {
    return "";
  }

  // 2. 判斷是否為 JSON 物件
  if (outputTemplate.trim().startsWith("{")) {
    try {
      const outputMap = JSON.parse(outputTemplate);
      // 從 session.data[storageKey] 取得對應的輸出當作 Key
      const lookupKey = String(session.data[storageKey] || "").trim();
      outputTemplate = outputMap[lookupKey] || outputMap["default"];
      
      if (outputTemplate === undefined) {
        Log(`[WARN] handleOutputConfig 在映射中找不到 Key: "${lookupKey}" 且無 default`);
        return null; // 找不到對應配置且無預設值時回傳 null
      }
    } catch (e) {
      // 若解析失敗則視為普通字串
    }
  }

  // 4. 透過 renderTemplate 轉換並回傳
  return renderTemplate(String(outputTemplate), session);
}

//
/**
 * 根據配置與 Session 內容獲取下一個狀態 ID
 * 1. 若 NextState 非 JSON，直接返回該值 (靜態跳轉)
 * 2. 若為 JSON，則根據 config.Storage 變數的值當作 key 取得對應目標狀態
 * @param {Object} config - 當前狀態之配置對象
 * @param {Object} session - 目前的 Session 對象
 * @returns {String} 下一狀態 ID，找不到則返回 "ERROR"
 */
function getNextState(config, session) {
  const nextConfig = config.NextState;
  const storageKey = config.Storage;

  // 1. 不是 JSON 則直接返回
  if (!nextConfig || !nextConfig.trim().startsWith("{")) {
    return nextConfig || "ERROR";
  }

  try {
    const stateMap = JSON.parse(nextConfig);
    // 2. 是 JSON 則根據 Storage 裡面的值當作 key
    const lookupKey = String(session.data[storageKey] || "").trim();
    const targetState = stateMap[lookupKey] || stateMap["default"] || "ERROR";
    Log(`[FLOW] JSON 跳轉映射: Key="${lookupKey}" -> Target="${targetState}"`);
    return targetState;
  } catch (e) {
    return nextConfig || "ERROR";
  }
}

/**
 * 將字串輸入轉換為陣列
 * 支援 JSON 陣列格式或以逗號分隔的字串格式 (例如 ["A", "B"] 與 CSV "A, B")
 * @param {String} input - 可能包含陣列格式的字串
 * @param {Boolean} keepEmpty - 是否保留空字串元素 (例如由連續逗號產生的空項)，預設為 true
 * @returns {Array} 轉換後的陣列
 */
function convertToArray(input, keepEmpty = true) {
  let result = [];

  if (!input) return result;

  let trimmed = String(input).trim();

  // 1. 嘗試解析為 JSON 陣列
  if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) {
        // 先映射並去空白，統一後續處理
        result = parsed.map(item => String(item).trim());
        // --- 修正：這裡應該立即處理 keepEmpty 並回傳結果 ---
        return keepEmpty ? result : result.filter(item => item !== "");
      }
    } catch (e) {
      // JSON 解析失敗 (通常因為內容有換行或特殊字元)，剝除外層括號繼續往下處理
      trimmed = trimmed.substring(1, trimmed.length - 1).trim();
    }
  }

  // 2. 處理以逗號分隔的格式 (CSV 備援方案)
  result = trimmed.split(",").map(item => {
    let s = item.trim();
    // 移除包裹在外的引號 (如果有)
    if (s.startsWith('"') && s.endsWith('"')) s = s.substring(1, s.length - 1);
    if (s.startsWith("'") && s.endsWith("'")) s = s.substring(1, s.length - 1);
    return s.trim();
  });
  
  // 3. 通用過濾邏輯：根據參數決定是否移除空字串 (包含原 JSON 陣列中的空項)
  return keepEmpty ? result : result.filter(item => item !== "");
}


/**
 * 將使用者的 Session 對象序列化並儲存至 Cache (PropertiesService)
 * @param {String} userId - 使用者的獨一 ID
 * @param {Object} session - 目前的使用者 Session 對象
 */
function saveSession(userId, session) {
  cache.setProperty(userId, JSON.stringify(session));
}

/**
 * 結束問卷流程後，從 Cache 中清除使用者的 Session 記憶
 * @param {String} userId - 使用者的獨一 ID
 */
function clearSession(userId) {
  cache.deleteProperty(userId);
}

/**
 * 規格化變數名稱
 * 如果輸入是 "${變數}" 格式，則移除 "${" 與 "}"
 * @param {String} key - 原始變數名稱
 * @returns {String} 移除標籤後的純變數名稱
 */
function normalizeStorageKey(key) {
  if (!key || key === "(null)") return null;
  const trimmed = String(key).trim();
  const match = trimmed.match(/^\$\{(.+?)\}$/);
  if (match) {
    return match[1].trim();
  }
  return trimmed;
}
