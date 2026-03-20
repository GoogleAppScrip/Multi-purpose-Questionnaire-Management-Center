/**
 * 系統變數解析模組
 * 負責處理問卷、分頁名稱與看板中使用的動態標籤
 */

/**
 * 樣板渲染引擎 (支援動態變數、標籤解析、遞回巢狀解析)
 * 1. 處理預設系統標籤 (日期、週數、時間)
 * 2. 處理 Session 專屬標籤 (使用者名稱、組別)
 * 3. 處理一般資料變數 (支援中文名稱與巢狀解析)
 * 4. 自動紀錄未定義標籤至 session.undefine
 * @param {String} template - 包含 ${var} 的原始文本
 * @param {Object} session - 目前的 Session 對象 (需包含 data 子物件)
 * @param {Number} depth - 遞回深度 (內部使用，防止死循環)
 * @returns {String} 替換後的最終文本，若超過遞回深度則回傳半成品
 */
function renderTemplate(template, session, depth = 0) {
  if (!template) return "";
  if (depth > 3) return String(template); // 安全閥：限制遞回深度

  const now = new Date();
  let rendered = String(template);

  // 1. 解析 ${WeekNumber} - 符合 ISO-8601 (1-53)
  rendered = rendered.replace(/\$\{WeekNumber\}/g, () => getWeekNumber(now));

  // 2. 解析 ${DateStart(n)} - n=1對應星期一, n=0對應星期天
  rendered = rendered.replace(/\$\{DateStart\((\d+)\)\}/g, (match, n) => {
    return getRecentDay(now, parseInt(n));
  });

  // 3. 解析 ${Now} - 當下時間 (年/月/日 時:分:秒)
  rendered = rendered.replace(/\$\{Now\}/g, () => {
    return Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  });

  // 4. 解析 ${UserName} (從 session.data 讀取，需配合萬用問卷設計規格書)
  rendered = rendered.replace(/\$\{UserName\}/g, () => {
    return session.data.UserName;
  });

  // 5. 解析 ${Group} (從 session.data 讀取)
  rendered = rendered.replace(/\$\{Group\}/g, () => {
    return session.data.Group;
  });
  
  // 6. 解析一般資料變數 ${var} (支援中文名稱，如 ${問卷名稱}、${儲存分頁})
  rendered = rendered.replace(/\$\{(.+?)\}/g, (match, key) => {
    return session.data[key] !== undefined ? session.data[key] : match;
  });

  // 7. 檢查 session.undefine 的變數，現在是否已經出現了，出現了就將他移除掉
  if (session.undefine && session.undefine.length > 0) {
    session.undefine = session.undefine.filter(key => session.data[key] === undefined);
  }

  // 8. 檢查遞回與未定義標籤需求
  const allvar = rendered.match(/\$\{(.+?)\}/g);
  if (allvar) {
    if (!session.undefine) session.undefine = [];
    
    // 檢查是否還有 ${var} 標籤，且該標籤的 key 存在於 session.data 中 (代表是剛替換出來的巢狀標籤)
    // 篩選有效的標籤 (存在於 session.data 中)
    const validVars = allvar.filter(match => {
      const key = match.slice(2, -1);
      if (session.data[key] === undefined) {
        // 若標籤不在 data 中，紀錄至 undefine 並從當前回合排除
        if (session.undefine.indexOf(key) === -1) session.undefine.push(key);
        return false;
      }
      return true;
    });

    // 9. 如果還有「有效且待解析」的標籤，則進行遞回處理
    if (validVars.length > 0) {
      return renderTemplate(rendered, session, depth + 1);
    }
  }

  return rendered;
}






/**
 * 根據使用者 ID 獲取顯示名稱 (優先生別名，無則回傳用戶名)
 * @param {String} userId - 使用者 ID
 * @returns {String} 格式化的使用者稱呼
 */
function getUserName(userId) {
  const userData = ss.getSheetByName("使用者資料");
  if (userData) {
    const data = userData.getDataRange().getValues();
    for (let row of data) {
      if (row[0] === userId && row[6] === "Active") return row[4] || row[1];
    }
  }
  return "同行善友";
}

/**
 * 根據使用者 ID 獲取組別名稱
 * @param {String} userId - 使用者 ID
 * @returns {String} 格式化的使用者組別
 */
function getUserGroup(userId) {
  const userData = ss.getSheetByName("使用者資料");
  if (userData) {
    const data = userData.getDataRange().getValues();
    for (let row of data) {
      if (row[0] === userId && row[6] === "Active") return row[5];
    }
  }
  return "未知組別";
}

/**
 * 取得年度第幾週 (ISO-8601)
 * @param {Date} d - 日期對象
 * @returns {Number} 年度週次
 */
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

/**
 * 計算最近一次指定星期幾的日期 (YYYYMMDD)
 * @param {Date} targetDate 基礎日期
 * @param {Number} targetDay 星期索引 (1=Mon, 2=Tue... 6=Sat, 0=Sun)
 */
function getRecentDay(targetDate, targetDay) {
  let d = new Date(targetDate);
  let currentDay = d.getDay(); // 0(Sun) - 6(Sat)
  
  // 計算差距
  let diff = currentDay - targetDay;
  if (diff < 0) diff += 7;
  
  d.setDate(d.getDate() - diff);
  return Utilities.formatDate(d, "Asia/Taipei", "yyyyMMdd");
}
