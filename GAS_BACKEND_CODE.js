/**
 * ==============================================================================
 * 臺南市安南區南興國民小學 115 學年度校慶運動大會
 * LINE 智慧賽事秘書 ＆ 雲端試算表橋接核心 (GAS 後端)
 * ==============================================================================
 * 核心功能：
 * 1. LINE Bot 接收自然語言（文字/語音），經 Gemini 2.5 Flash 自動解析結構化資料。
 * 2. 嚴格白名單權限防護（非授權者直接阻擋，保護資料庫防竄改並節省 API 額度）。
 * 3. 自動寫入 Google 試算表（精神加分明細、戰隊總積分、各項賽事成績、個人短跑得獎榜）。
 * 4. 提供 doGet Webhook JSON API，讓官方網站（index/scores/spirit）每 30 秒自動讀取最新分數！
 * ==============================================================================
 */

// 系統預設組態
const CONFIG = {
  GEMINI_MODEL: 'gemini-2.5-flash',
  ENABLE_WHITELIST: true, // 開啟白名單機制（防外人亂改）
  SHEET_NAMES: {
    TOTALS: '戰隊總積分榜',
    EVENTS: '各項賽事成績明細',
    SPRINT: '個人短跑得獎榜',
    SPIRIT: '精神加分明細',
    GUESTS: '貴賓線上回條名冊'
  }
};

/**
 * ----------------------------------------------------------------------
 * 1. 提供給官方網頁讀取的公開唯讀 JSON API (doGet)
 * - 支援 ?type=guests 讀取貴賓回條清單（供接待管理後台 admin.html 載入）
 * - 預設回傳：三大戰隊總積分、各項賽事成績、個人短跑得獎名冊、精神事蹟日誌
 * ----------------------------------------------------------------------
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetHeaders(ss);

    // A. 查詢貴賓線上回條名冊
    if (e && e.parameter && e.parameter.type === 'guests') {
      const guestSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.GUESTS);
      const data = guestSheet.getDataRange().getValues();
      const guests = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue;
        guests.push({
          id: i,
          created_at: Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'yyyy/MM/dd HH:mm'),
          org: row[1],
          name: row[2],
          title: row[3],
          attendance: row[4],
          delegate: row[5],
          car: row[6],
          speech: row[7],
          sponsor: row[8],
          wishes: row[9]
        });
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        total: guests.length,
        guests: guests.reverse()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // B. 讀取「精神加分明細」
    const spiritSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);
    const spiritData = spiritSheet.getDataRange().getValues();
    const stories = [];
    let greenSpirit = 0, blueSpirit = 0, redSpirit = 0;

    for (let i = 1; i < spiritData.length; i++) {
      const row = spiritData[i];
      if (!row[0]) continue;
      const team = String(row[2]).trim(); // 'green', 'blue', 'red'
      const pts = Number(row[4]) || 0;

      if (team === 'green' || team.includes('綠')) greenSpirit += pts;
      else if (team === 'blue' || team.includes('藍')) blueSpirit += pts;
      else if (team === 'red' || team.includes('紅')) redSpirit += pts;

      stories.push({
        id: i,
        time: Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'MM/dd HH:mm'),
        team: normalizeTeam(team),
        category: row[3],
        points: pts,
        story: row[5],
        author: row[6] || '大會裁判',
        place: row[7] || '操場'
      });
    }
    stories.reverse();

    // C. 讀取「各項賽事成績明細」
    const eventSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVENTS);
    const eventData = eventSheet.getDataRange().getValues();
    const eventScores = {};
    let greenSports = 0, blueSports = 0, redSports = 0;

    for (let i = 1; i < eventData.length; i++) {
      const row = eventData[i];
      if (!row[0]) continue;
      const eventCode = String(row[2]).trim(); // e.g. 'dodge', 'runLow', etc.
      const gPts = Number(row[3]) || 0;
      const bPts = Number(row[4]) || 0;
      const rPts = Number(row[5]) || 0;

      eventScores[eventCode] = { green: gPts, blue: bPts, red: rPts, memo: row[6] || '' };
      greenSports += gPts;
      blueSports += bPts;
      redSports += rPts;
    }

    // D. 讀取「個人短跑得獎榜」
    const sprintSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPRINT);
    const sprintData = sprintSheet.getDataRange().getValues();
    const sprintWinners = {};

    for (let i = 1; i < sprintData.length; i++) {
      const row = sprintData[i];
      if (!row[0]) continue;
      const groupCode = String(row[0]).trim(); // 'lowBoy', 'lowGirl', etc.
      sprintWinners[groupCode] = {
        name: row[1],
        dist: row[2],
        gold: row[3] || '尚未決賽',
        silver: row[4] || '尚未決賽',
        bronze: row[5] || '尚未決賽'
      };
    }

    // E. 彙整三大戰隊總分（體育分 + 精神分）
    const teamTotals = {
      green: { sports: greenSports, spirit: greenSpirit, total: greenSports + greenSpirit },
      blue: { sports: blueSports, spirit: blueSpirit, total: blueSports + blueSpirit },
      red: { sports: redSports, spirit: redSpirit, total: redSports + redSpirit }
    };

    // F. 組合完整結果回傳
    const result = {
      success: true,
      timestamp: new Date().toISOString(),
      teamTotals: teamTotals,
      eventScores: eventScores,
      sprintWinners: sprintWinners,
      spiritScores: { green: greenSpirit, blue: blueSpirit, red: redSpirit },
      stories: stories
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

    for (let i = 1; i < spiritData.length; i++) {
      const row = spiritData[i];
      if (!row[0]) continue;
      const team = String(row[2]).trim(); // 'green', 'blue', 'red'
      const pts = Number(row[4]) || 0;

      if (team === 'green' || team.includes('綠')) greenSpirit += pts;
      else if (team === 'blue' || team.includes('藍')) blueSpirit += pts;
      else if (team === 'red' || team.includes('紅')) redSpirit += pts;

      stories.push({
        id: i,
        time: Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'MM/dd HH:mm'),
        team: normalizeTeam(team),
        category: row[3],
        points: pts,
        story: row[5],
        author: row[6] || '大會裁判',
        place: row[7] || '操場'
      });
    }

    // 最新故事排在最前面
    stories.reverse();

    // 2. 組合回傳結果
    const result = {
      success: true,
      timestamp: new Date().toISOString(),
      spiritScores: {
        green: greenSpirit,
        blue: blueSpirit,
        red: redSpirit
      },
      stories: stories
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ----------------------------------------------------------------------
 * 2. 接收 LINE 訊息之 Webhook 與 貴賓線上回條提交 (doPost)
 * ----------------------------------------------------------------------
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 🌟 A. 處理來自「貴賓線上回條 (invite.html)」的填寫提交
    if (data && data.action === 'SUBMIT_INVITE') {
      const res = recordGuestInvite(data.guest);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 🌟 B. 處理來自 LINE 官方帳號的 Webhook 訊息
    const events = data.events;
    if (!events || events.length === 0) return ContentService.createTextOutput('OK');

    const event = events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;
    const messageType = event.message ? event.message.type : '';

    // 🛡️ 支柱 1：白名單防禦檢核
    if (!isUserAuthorized(userId)) {
      replyLineMessage(replyToken, '您好！南興國小 115 校慶賽事秘書系統目前為【管理員與裁判教師專用】。\n若您是南興教職員，請先將您的 User ID 回報給吳校長進行綁定授權：\n' + userId);
      return ContentService.createTextOutput('BLOCKED');
    }

    // 處理管理員特殊指令
    if (messageType === 'text') {
      const text = event.message.text.trim();
      
      // 查詢自己的 ID
      if (text === '我的ID' || text === '我的id') {
        replyLineMessage(replyToken, `您的 LINE User ID 為：\n${userId}\n（請將此代碼交由管理者加入白名單）`);
        return ContentService.createTextOutput('OK');
      }

      // 撤銷上一筆
      if (text === '/undo' || text === '撤銷上一筆' || text === '復原') {
        const undoMsg = undoLastRecord();
        replyLineMessage(replyToken, undoMsg);
        return ContentService.createTextOutput('OK');
      }
    }

    // 提取使用者輸入文字（純文字或語音）
    let userPrompt = '';
    if (messageType === 'text') {
      userPrompt = event.message.text.trim();
    } else if (messageType === 'audio') {
      // 語音多模態辨識（需使用 LINE Content API 下載後傳入 Gemini）
      userPrompt = fetchAndTranscribeAudio(event.message.id);
      if (!userPrompt) {
        replyLineMessage(replyToken, '語音辨識未完成，請嘗試改以打字輸入，謝謝！');
        return ContentService.createTextOutput('OK');
      }
    } else {
      replyLineMessage(replyToken, '目前支援【文字輸入】與【語音訊息】登記賽事與精神加分喔！');
      return ContentService.createTextOutput('OK');
    }

    // 🧠 支柱 2：Gemini 2.5 Flash 智慧語意解析
    const parsed = callGeminiParser(userPrompt, userId);

    if (!parsed || !parsed.intent) {
      replyLineMessage(replyToken, '收到訊息，但未能清楚辨識加分意圖。\n建議範例：\n「綠隊精神加分2分，六年級主動牽一年級去檢錄」\n「高年級男生100m第一名陳小明、第二名林大華、第三名張小傑」');
      return ContentService.createTextOutput('OK');
    }

    // 執行寫入動作
    let replyText = '';
    if (parsed.intent === 'ADD_SPIRIT') {
      replyText = recordSpiritPoint(parsed);
    } else if (parsed.intent === 'UPDATE_SPRINT') {
      replyText = recordSprintWinner(parsed);
    } else if (parsed.intent === 'UPDATE_EVENT') {
      replyText = recordEventScores(parsed);
    } else {
      replyText = `收到您的訊息：${userPrompt}\n（目前未對應至已定義的賽事項目）`;
    }

    replyLineMessage(replyToken, replyText);
    return ContentService.createTextOutput('OK');

  } catch (err) {
    console.error('doPost Error:', err);
    return ContentService.createTextOutput('ERROR: ' + err.toString());
  }
}

/**
 * ----------------------------------------------------------------------
 * 3. 呼叫 Gemini 進行意圖分析與結構化 JSON 輸出
 * ----------------------------------------------------------------------
 */
function callGeminiParser(userText, userId) {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('請先在 Script Properties 中設定 GEMINI_API_KEY！');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const systemInstruction = `
你是一位專門為「臺南市安南區南興國小 115 學年度校慶運動大會」服務的 AI 賽事秘書。
你的任務是將校長或老師輸入的隨性文字或語音轉錄內容，解析為嚴格的結構化 JSON。

學校賽事基本定義：
1. 戰隊：'green' (綠隊/翠玉蜈蚣), 'blue' (藍隊/碧海晨曦), 'red' (紅隊/烈火金獅)。
2. 精神加分向度（四大向度）：
   - 'cross' (跨齡扶持 / 大帶小)
   - 'sports' (運動家風度)
   - 'eco' (環保自律)
   - 'cheer' (團隊歡呼 / 隊呼)
3. 個人短跑組別：
   - 'lowBoy' (低年級男生組 60m)
   - 'lowGirl' (低年級女生組 60m)
   - 'midBoy' (中年級男生組 80m)
   - 'midGirl' (中年級女生組 80m)
   - 'highBoy' (高年級男生組 100m)
   - 'highGirl' (高年級女生組 100m)

請根據使用者輸入判斷 intent：
- intent: "ADD_SPIRIT" (例如：綠隊精神加分2分，六年級幫學弟綁鞋帶)
  輸出欄位: { "intent": "ADD_SPIRIT", "team": "green"|"blue"|"red", "category": "cross"|"sports"|"eco"|"cheer", "points": 1~3, "story": "完整事蹟全文", "author": "老師名稱若有提及否則填寫裁判", "place": "操場/風雨球場/司令台等" }

- intent: "UPDATE_SPRINT" (例如：高年級男生組第一名六年級陳小明、第二名林大華、第三名張小傑)
  輸出欄位: { "intent": "UPDATE_SPRINT", "group": "lowBoy"|"lowGirl"|"midBoy"|"midGirl"|"highBoy"|"highGirl", "gold": "姓名", "silver": "姓名", "bronze": "姓名" }

- intent: "UPDATE_EVENT" (例如：拔河低年級組第一名綠隊、第二名紅隊、第三名藍隊)
  輸出欄位: { "intent": "UPDATE_EVENT", "event": "dodge"|"tugLow"|"tugHigh"|"relayLow"|"relayHigh", "rank1": "green"|"blue"|"red", "rank2": "green"|"blue"|"red", "rank3": "green"|"blue"|"red" }

必須只回傳合法的純 JSON，絕不加 markdown 標籤（如 \`\`\`json ）。
`;

  const payload = {
    contents: [
      { role: 'user', parts: [{ text: userText }] }
    ],
    systemInstruction: {
      parts: [{ text: systemInstruction }]
    },
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.candidates && json.candidates[0].content.parts[0].text) {
    const raw = json.candidates[0].content.parts[0].text.trim();
    return JSON.parse(raw);
  }
  return null;
}

/**
 * ----------------------------------------------------------------------
 * 4. 寫入資料庫：精神加分記錄
 * ----------------------------------------------------------------------
 */
function recordSpiritPoint(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders(ss);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);

  const teamNameMap = { green: '🟢 綠隊', blue: '🔵 藍隊', red: '🔴 紅隊' };
  const catNameMap = { cross: '🤝 跨齡扶持', sports: '🏃 運動家風度', eco: '🌿 環保自律', cheer: '📣 團隊歡呼' };

  const now = new Date();
  const row = [
    now,
    teamNameMap[p.team] || p.team,
    p.team,
    catNameMap[p.category] || p.category,
    p.points || 2,
    p.story,
    p.author || '大會裁判',
    p.place || '操場'
  ];

  sheet.appendRow(row);

  // 暫存上一筆資訊供 /undo 撤銷使用
  PropertiesService.getScriptProperties().setProperty('LAST_SPIRIT_ROW', sheet.getLastRow().toString());

  return `🎉【精神加分登記成功！】\n` +
         `━━━━━━━━━━━━━━\n` +
         `戰隊：${teamNameMap[p.team] || p.team} (+${p.points || 2}分)\n` +
         `向度：${catNameMap[p.category] || p.category}\n` +
         `事蹟：${p.story}\n` +
         `見證：${p.author || '大會裁判'} (${p.place || '操場'})\n` +
         `━━━━━━━━━━━━━━\n` +
         `💡 官網首頁與精神榮譽館已即時同步更新！\n若登記有誤，請於 3 分鐘內輸入 /undo 即可撤銷。`;
}

/**
 * ----------------------------------------------------------------------
 * 5. 撤銷上一筆記錄 (/undo)
 * ----------------------------------------------------------------------
 */
function undoLastRecord() {
  const scriptProps = PropertiesService.getScriptProperties();
  const lastRowStr = scriptProps.getProperty('LAST_SPIRIT_ROW');
  if (!lastRowStr) return '目前沒有可撤銷的精神加分紀錄。';

  const rowNum = parseInt(lastRowStr, 10);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);

  if (rowNum > 1 && rowNum <= sheet.getLastRow()) {
    const deletedData = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];
    sheet.deleteRow(rowNum);
    scriptProps.deleteProperty('LAST_SPIRIT_ROW');
    return `↩️【已成功撤銷上一筆精神加分】\n戰隊：${deletedData[1]} (-${deletedData[4]}分)\n事蹟：${deletedData[5]}\n官網資料已即時更正！`;
  }
  return '撤銷失敗，該列已不存在或已經撤銷過。';
}

/**
 * ----------------------------------------------------------------------
 * 6. 權限防禦：白名單檢核
 * ----------------------------------------------------------------------
 */
function isUserAuthorized(userId) {
  const scriptProps = PropertiesService.getScriptProperties();
  if (CONFIG.ENABLE_WHITELIST === false) return true;
  if (!userId) return false;

  let adminId = scriptProps.getProperty('ADMIN_USER_ID');
  if (!adminId) {
    // 第一次傳訊的老師自動綁定為最高管理員（通常是校長）
    scriptProps.setProperty('ADMIN_USER_ID', userId);
    return true;
  }
  if (userId === adminId) return true;

  // 檢查額外白名單（逗號隔開的 LINE ID）
  const whitelist = scriptProps.getProperty('EXTRA_WHITELIST_USERS') || '';
  if (whitelist.split(',').map(s => s.trim()).includes(userId)) return true;

  return false;
}

/**
 * ----------------------------------------------------------------------
 * 7. 試算表自癒建表函式（自動建立五大核心分頁與標題列）
 * ----------------------------------------------------------------------
 */
function ensureSheetHeaders(ss) {
  // 1. 各項賽事成績明細分頁
  let eventSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVENTS);
  if (!eventSheet) {
    eventSheet = ss.insertSheet(CONFIG.SHEET_NAMES.EVENTS);
    eventSheet.appendRow(['賽事日期', '賽事項目名稱', '賽事代碼', '綠隊得分', '藍隊得分', '紅隊得分', '名次與備註說明', '裁判登記人', '更新時間']);
    eventSheet.getRange(1, 1, 1, 9).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
    eventSheet.setFrozenRows(1);

    // 寫入預設賽事架構
    const defaultEvents = [
      ['11/26 (四)', '熱血美式躲避球循環賽', 'dodge', 12, 4, 8, '綠隊冠軍(+12)、紅隊亞軍(+8)、藍隊季軍(+4)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-低年級3圈(5分內)', 'runLow', 6, 7, 5, '藍7人、綠6人、紅5人達標(每人1分)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-中年級800m(7分內)', 'runMid', 6, 4, 5, '綠6人、紅5人、藍4人達標(每人1分)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-高年級800m(6分內)', 'runHigh', 5, 4, 6, '紅6人、綠5人、藍4人達標(每人1分)', '大會裁判組', new Date()],
      ['12/05 (六)', '拔河賽-一二三年級組', 'tugLow', 8, 6, 4, '綠隊冠軍(+8)、藍隊亞軍(+6)、紅隊季軍(+4)', '大會裁判組', new Date()],
      ['12/05 (六)', '拔河賽-四五六年級組', 'tugHigh', 6, 4, 8, '紅隊冠軍(+8)、綠隊亞軍(+6)、藍隊季軍(+4)', '大會裁判組', new Date()],
      ['12/05 (六)', '趣味競賽-蟲蟲危機', 'rainbow', 6, 12, 8, '藍隊冠軍(+12)、紅隊亞軍(+8)、綠隊季軍(+6)', '大會裁判組', new Date()],
      ['12/05 (六)', '大隊接力-一二三年級組', 'relayLow', 10, 12, 8, '藍隊冠軍(+12)、綠隊亞軍(+10)、紅隊季軍(+8)', '大會裁判組', new Date()],
      ['12/05 (六)', '大隊接力-四五六年級組', 'relayHigh', 12, 8, 10, '綠隊冠軍(+12)、紅隊亞軍(+10)、藍隊季軍(+8)', '大會裁判組', new Date()]
    ];
    eventSheet.getRange(2, 1, defaultEvents.length, 9).setValues(defaultEvents);
  }

  // 2. 個人短跑得獎榜分頁
  let sprintSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPRINT);
  if (!sprintSheet) {
    sprintSheet = ss.insertSheet(CONFIG.SHEET_NAMES.SPRINT);
    sprintSheet.appendRow(['組別代碼', '競賽組別名稱', '競賽距離', '🥇金牌(第1名)', '🥈銀牌(第2名)', '🥉銅牌(第3名)', '裁判簽名', '更新時間']);
    sprintSheet.getRange(1, 1, 1, 8).setBackground('#0369a1').setFontColor('#ffffff').setFontWeight('bold');
    sprintSheet.setFrozenRows(1);

    const defaultSprints = [
      ['lowBoy', '低年級男生組', '60m', '二年級代表A', '二年級代表B', '一年級代表A', '大會裁判組', new Date()],
      ['lowGirl', '低年級女生組', '60m', '二年級代表C', '一年級代表B', '一年級代表C', '大會裁判組', new Date()],
      ['midBoy', '中年級男生組', '80m', '四年級代表A', '四年級代表B', '三年級代表A', '大會裁判組', new Date()],
      ['midGirl', '中年級女生組', '80m', '四年級代表C', '三年級代表B', '三年級代表C', '大會裁判組', new Date()],
      ['highBoy', '高年級男生組', '100m', '六年級代表A', '五年級代表A', '六年級代表B', '大會裁判組', new Date()],
      ['highGirl', '高年級女生組', '100m', '六年級代表C', '五年級代表B', '五年級代表C', '大會裁判組', new Date()]
    ];
    sprintSheet.getRange(2, 1, defaultSprints.length, 8).setValues(defaultSprints);
  }

  // 3. 精神加分明細分頁
  let spiritSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);
  if (!spiritSheet) {
    spiritSheet = ss.insertSheet(CONFIG.SHEET_NAMES.SPIRIT);
    spiritSheet.appendRow(['記錄時間', '戰隊名稱', '戰隊代碼', '評審向度', '得分', '優良事蹟全文', '見證認證教師', '地點位置']);
    spiritSheet.getRange(1, 1, 1, 8).setBackground('#4338ca').setFontColor('#ffffff').setFontWeight('bold');
    spiritSheet.setFrozenRows(1);
  }

  // 4. 貴賓線上回條名冊分頁
  let guestSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.GUESTS);
  if (!guestSheet) {
    guestSheet = ss.insertSheet(CONFIG.SHEET_NAMES.GUESTS);
    guestSheet.appendRow([
      '填寫時間', '服務機關/單位', '貴賓姓名', '職稱', '出席意向', 
      '代表出席人', '車牌號碼', '致詞意願', '贊助意願', '賀詞與期勉'
    ]);
    guestSheet.getRange(1, 1, 1, 10).setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold');
    guestSheet.setFrozenRows(1);
  }
}

/**
 * ----------------------------------------------------------------------
 * 8. 寫入資料庫：貴賓線上回條
 * ----------------------------------------------------------------------
 */
function recordGuestInvite(g) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetHeaders(ss);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.GUESTS);

    const now = new Date();
    const row = [
      now,
      g.org || '',
      g.name || '',
      g.title || '',
      g.attendance || '',
      g.delegate || '',
      g.car || '',
      g.speech || '',
      g.sponsor || '無',
      g.wishes || ''
    ];

    sheet.appendRow(row);

    // 🔔 立即透過 LINE 主動推播通知管理員（校長）
    const notifyMsg = 
      `🔔【南興115校慶 • 新貴賓回條通報】\n` +
      `━━━━━━━━━━━━━━\n` +
      `單位：${g.org || '未填寫'}\n` +
      `貴賓：${g.name || '長官'} (${g.title || '貴賓'})\n` +
      `出席意向：${g.attendance}\n` +
      (g.delegate ? `指派代表：${g.delegate}\n` : '') +
      (g.car ? `車牌安排：${g.car}\n` : '') +
      `致詞意願：${g.speech}\n` +
      (g.sponsor && g.sponsor !== '無' ? `贊助意向：${g.sponsor}\n` : '') +
      (g.wishes ? `祝福賀詞：${g.wishes}\n` : '') +
      `━━━━━━━━━━━━━━\n` +
      `名冊已自動存入 Google 試算表！`;

    pushLineNotification(notifyMsg);

    return { success: true, message: '貴賓線上回條已成功記錄！' };
  } catch (err) {
    console.error('recordGuestInvite Error:', err);
    return { success: false, error: err.toString() };
  }
}

/**
 * ----------------------------------------------------------------------
 * 9. 發送 LINE 回覆訊息 (Reply)
 * ----------------------------------------------------------------------
 */
function replyLineMessage(replyToken, textMessage) {
  const scriptProps = PropertiesService.getScriptProperties();
  const channelToken = scriptProps.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!channelToken) return;

  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: textMessage }]
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + channelToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

/**
 * ----------------------------------------------------------------------
 * 10. 主動推播訊息給管理員 (Push Notification)
 * ----------------------------------------------------------------------
 */
function pushLineNotification(textMessage) {
  const scriptProps = PropertiesService.getScriptProperties();
  const channelToken = scriptProps.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const adminId = scriptProps.getProperty('ADMIN_USER_ID');
  if (!channelToken || !adminId) return;

  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: adminId,
    messages: [{ type: 'text', text: textMessage }]
  };

  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + channelToken
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('Push Notification Error:', e);
  }
}

function normalizeTeam(t) {
  if (t === 'green' || t.includes('綠')) return 'green';
  if (t === 'blue' || t.includes('藍')) return 'blue';
  if (t === 'red' || t.includes('紅')) return 'red';
  return 'green';
}
