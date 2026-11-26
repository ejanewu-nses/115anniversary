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
  GEMINI_MODEL: 'gemini-2.0-flash', // 穩定高速主力模型
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
 * 手動一鍵初始化資料庫分頁（在 Apps Script 編輯器上方選擇此函式並點擊「執行」即可秒速建好 4 大核心分頁！）
 * ----------------------------------------------------------------------
 */
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders(ss);
  
  // 若預設的「工作表1」是空的，自動移除保持整潔
  const defaultSheet = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }
  console.log('✅ 資料庫核心工作表已成功建立！');
}

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

      const rawCat = String(row[3] || '');
      let catCode = 'cross';
      let catLabel = '🤝 跨齡扶持';
      if (rawCat.includes('風度') || rawCat.includes('sports')) { catCode = 'sports'; catLabel = '🏃 運動家風度'; }
      else if (rawCat.includes('環保') || rawCat.includes('eco')) { catCode = 'eco'; catLabel = '🌿 環保自律'; }
      else if (rawCat.includes('歡呼') || rawCat.includes('cheer')) { catCode = 'cheer'; catLabel = '📣 團隊歡呼'; }

      const storyText = String(row[5] || '');
      const storyTitle = generateStoryTitle(storyText, catCode);

      stories.push({
        id: i,
        time: Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'MM/dd HH:mm'),
        team: normalizeTeam(team),
        category: catCode,
        categoryLabel: catLabel,
        title: storyTitle,
        points: pts,
        story: storyText,
        author: row[6] || '大會裁判',
        place: row[7] ? String(row[7]).trim() : ''
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
    const sourceType = (event.source && event.source.type) ? event.source.type : 'user';
    const isGroup = (sourceType === 'group' || sourceType === 'room');
    const userId = (event.source && event.source.userId) ? event.source.userId : '';
    const messageType = event.message ? event.message.type : '';

    // 🛡️ 支柱 1：白名單與群組防禦檢核
    // 若來自校內工作群組（group / room），全員自動授權免個別登記！若為個人 1 對 1 私聊，則檢核白名單防禦
    if (!isGroup && !isUserAuthorized(userId)) {
      replyLineMessage(replyToken, '您好！南興國小 115 校慶賽事秘書系統目前為【管理員與裁判教師專用】。\n若您是南興教職員，請先將您的 User ID 回報給吳校長進行綁定授權：\n' + userId);
      return ContentService.createTextOutput('BLOCKED');
    }

    // 處理管理員與系統特殊指令
    if (messageType === 'text') {
      const text = event.message.text.trim();
      
      // 查詢自己的 ID
      if (text === '我的ID' || text === '我的id') {
        replyLineMessage(replyToken, `您的 LINE User ID 為：\n${userId || '（未提供）'}\n${isGroup ? '💡 提示：您正在南興校慶裁判工作群中，本群同仁皆已自動具備登記權限！' : '（請將此代碼交由管理者加入白名單）'}`);
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
        if (!isGroup) {
          replyLineMessage(replyToken, '語音辨識未完成，請嘗試改以打字輸入，謝謝！');
        }
        return ContentService.createTextOutput('OK');
      }
    } else {
      // 貼圖或其它非文字訊息：在群組中靜默忽略防洗版，私聊才提示
      if (!isGroup) {
        replyLineMessage(replyToken, '目前支援【文字輸入】與【語音訊息】登記賽事與精神加分喔！');
      }
      return ContentService.createTextOutput('OK');
    }

    // 🧠 支柱 2：雙引擎智慧語意解析 (Fast Local + Gemini Flash)
    const parsed = callGeminiParser(userPrompt, userId);

    if (!parsed || !parsed.intent) {
      // 🔇 關鍵防洗版機制：
      // 若在群組內，一般同仁日常討論或非賽事指令（如「辛苦了」、「便當到了」、「收到」），機器人保持靜音（Silent），絕不插嘴打擾！
      if (isGroup) {
        return ContentService.createTextOutput('IGNORE_CHAT');
      }

      // 若為 1 對 1 私聊，才提供提示與範例指引
      replyLineMessage(replyToken, '收到訊息，但未能清楚辨識意圖。\n建議範例：\n【精神加分】「綠隊精神加分2分，六年級主動牽一年級去檢錄」\n【賽事報分】「拔河低年級綠隊8分、藍隊6分、紅隊4分」\n【慢跑人數】「慢跑低年級藍隊7人、綠隊6人、紅隊5人」\n【短跑名次】「高年級男生100m第一名陳小明、第二名林大華、第三名張小傑」');
      return ContentService.createTextOutput('OK');
    }

    // 執行寫入動作
    let replyText = '';
    if (parsed.intent === 'INVALID_TEAM') {
      replyText = parsed.message;
    } else if (parsed.intent === 'ADD_SPIRIT') {
      replyText = recordSpiritPoint(parsed);
    } else if (parsed.intent === 'UPDATE_SPRINT') {
      replyText = recordSprintWinner(parsed);
    } else if (parsed.intent === 'UPDATE_EVENT') {
      replyText = recordEventScores(parsed);
    } else {
      if (isGroup) return ContentService.createTextOutput('IGNORE_CHAT');
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
 * 3. 快速本機規則引擎 (0 毫秒極速辨識，防網路/模型抖動)
 * ----------------------------------------------------------------------
 */
function fastLocalParse(text) {
  if (!text) return null;
  text = text.trim();

  // 1. 判斷是否為明確的「賽事成績/名次登記」
  // 必須明確具備「第一名/第1名/冠軍/亞軍/季軍」或者特定賽事比分/人數報分格式
  const hasRankWords = /(?:第一名|第1名|第一|第1|第二名|第2名|第二|第2|第三名|第3名|第三|第3|冠軍|亞軍|季軍)/.test(text);
  const isScoreReport = /(?:拔河|躲避球|接力|趣味|蟲蟲|彩虹|慢跑).*(?:\d+\s*分|\d+\s*人)/.test(text);
  const isSprint = /(?:100m|60m|80m|短跑).*(?:第一|第1|第二|第2|第三|第3|冠軍|金牌)/.test(text);

  // 判斷是否具備「精神加分」特徵
  const isSpiritExplicit = /(?:精神|暖心|事蹟|好品德|大帶小|運動家|歡呼|隊呼|自律|整齊|互助|清潔|環境|勉勵|照顧|幫忙|禮貌|聽講|規則|表現|秩序|熱身|充分熱身|加油)/.test(text) ||
                           /(?:各加|各得|都加|加分|精神加|精神各加)/.test(text) ||
                           /(?:[+＋加]\s*[1-9一二兩三\d]\s*分?|[1-9一二兩三\d]\s*分)/.test(text);

  // 如果帶有明確的精神加分語意，且「不是」在登記賽事名次或短跑名次，則優先走精神加分
  const isSpirit = isSpiritExplicit && (!hasRankWords && !isSprint);

  if (isSpirit) {
    // 🛡️ 戰隊名稱防呆檢核（防止打錯為黃隊、白隊、黑隊等非本校戰隊）
    const wrongTeamMatch = text.match(/([黃白黑紫橘粉灰棕橙])(?:隊)?/);
    if (wrongTeamMatch) {
      return {
        intent: 'INVALID_TEAM',
        message: `⚠️【隊伍名稱輸入有誤：${wrongTeamMatch[0]}】\n南興國小僅有「🔴 紅隊」、「🟢 綠隊」、「🔵 藍隊」三大戰隊。\n本次精神加分未登記，請確認隊伍後重新輸入！`
      };
    }

    let teams = [];
    const isAllTeams = /(?:紅\s*綠\s*藍|三隊|三大戰隊|各隊|每隊|全隊|全校|大家|全部戰隊|三戰隊)/.test(text) ||
                       (/(?:紅|烈火|金獅)/.test(text) && /(?:綠|翠玉|蜈蚣)/.test(text) && /(?:藍|跳鼓|雷鳴|碧海)/.test(text));

    if (isAllTeams) {
      teams = ['red', 'green', 'blue'];
    } else {
      if (/(?:紅|烈火|金獅)/.test(text)) teams.push('red');
      if (/(?:綠|翠玉|蜈蚣)/.test(text)) teams.push('green');
      if (/(?:藍|跳鼓|雷鳴|碧海)/.test(text)) teams.push('blue');
    }

    if (teams.length === 0) {
      return {
        intent: 'INVALID_TEAM',
        message: `⚠️【未指定有效戰隊】\n南興國小僅有「🔴 紅隊」、「🟢 綠隊」、「🔵 藍隊」三大戰隊。\n請於訊息中指明為哪一隊（紅/綠/藍隊，或「紅綠藍三隊各加...」）加分喔！`
      };
    }

    let pts = 1;
    const ptsMatch = text.match(/(?:(?:加|得|[+＋])\s*([1-9一二兩三\d])\s*分?|([1-9一二兩三\d])\s*分)/);
    if (ptsMatch) {
      const v = ptsMatch[1] || ptsMatch[2];
      if (v === '一' || v === '1') pts = 1;
      else if (v === '二' || v === '兩' || v === '2') pts = 2;
      else if (v === '三' || v === '3') pts = 3;
      else pts = parseInt(v, 10) || 1;
    } else {
      pts = 2; // 預設加2分
    }

    // 1. 精準提取登記人 / 見證人（支援冒號或空格隔開，如「見證人 校長」）
    let author = '';
    const authorMatch = text.match(/(?:登記人|見證人|見證者|通報人|認證老師|裁判|紀錄人|記錄人|見證|通報)[：:\s]+([^\s,，、。\n\r]+)/);
    if (authorMatch) {
      author = authorMatch[1].trim();
      if (author === '校長') author = '吳宜蓁校長';
    } else if (/吳宜蓁/.test(text)) {
      author = '吳宜蓁校長';
    }

    // 2. 提取地點位置（支援「地點 小禮堂」、「地點在小禮堂」、「見證地點在小禮堂」、「位置：操場」、「在小禮堂」等自然語意）
    let place = '';
    const explicitPlaceMatch = text.match(/(?:見證)?(?:地點|位置)(?:在|[：:\s]+)+([^\s,，、。\n\r]+)/);
    if (explicitPlaceMatch) {
      place = explicitPlaceMatch[1].trim();
    } else {
      const atMatch = text.match(/在\s*([^\s,，、。\n\r]+)/);
      if (atMatch) {
        const cand = atMatch[1].trim();
        const sub = cand.match(/(?:B1\s*)?(?:小?禮堂|操場(?:草地)?|風雨球場|球場|司令台|跑道|穿堂|走廊|教室|樹下|休息區|看台|大門|校門口?|活動中心|辦公室|健康中心)/);
        if (sub) place = sub[0];
        else if (/禮堂|操場|球場|司令台|跑道|穿堂|走廊|教室|草地|樹下|看台|大門|活動中心|辦公室|健康中心|B1/.test(cand)) {
          place = cand;
        }
      }
    }
    if (!place) {
      const directMatch = text.match(/(?:B1\s*)?(?:小?禮堂|操場(?:草地)?|風雨球場|司令台|穿堂|看台|活動中心)/);
      if (directMatch) {
        place = directMatch[0];
      }
    }

    // 3. 智慧權重計分法決定向度（全面防禦副詞干擾）
    let scoreCheer = 0, scoreSports = 0, scoreEco = 0, scoreCross = 0;

    // sports（運動家風度）
    if (/運動家|風度|勝不驕|敗不餒|握手|鞠躬|致敬|服從|擊掌|拍肩|道歉|輸了|落後|不氣餒|勉勵|彼此勉勵|互勉|安慰|再接再厲|熱身|充分熱身|規則|聽講/.test(text)) scoreSports += 4;
    if (/禮貌|友愛|謙虛|誠實|運動員|鼓舞|參與|熱忱/.test(text)) scoreSports += 2;

    // cheer（團隊歡呼）
    if (/歡呼|隊呼|加油|吶喊|口號|鼓掌|喝采|鼓舞|精神抖擻|士氣/.test(text)) scoreCheer += 4;
    if (/熱情|活力/.test(text)) scoreCheer += 1;

    // cross（跨齡扶持）
    if (/大帶小|跨齡|小隊長|綁鞋帶|牽手|拿外套|遞外套|學弟|學妹/.test(text)) scoreCross += 4;
    if (/照顧|幫忙|協助|牽|扶|陪|低年級|一年級|二年級/.test(text)) scoreCross += 2;

    // eco（環保自律）
    if (/垃圾|撿垃圾|清潔環境|整理場地|整理環境|整理器材|收拾器材|收拾|場地清潔|場地復原|落葉|掃地|做整潔|資源回收|寶特瓶|撿拾|打掃|場地整理/.test(text)) scoreEco += 4;
    if (/乾淨|自律|整齊|排隊|環境|清潔|水壺|整理|復原|衛生|場地/.test(text)) scoreEco += 2;

    let category = 'sports';
    const maxScore = Math.max(scoreSports, scoreCheer, scoreEco, scoreCross);
    if (maxScore > 0) {
      if (scoreEco === maxScore) category = 'eco';
      else if (scoreCross === maxScore) category = 'cross';
      else if (scoreCheer === maxScore) category = 'cheer';
      else if (scoreSports === maxScore) category = 'sports';
    } else {
      if (/整理|場地|乾淨|掃|清潔|垃圾|復原|收拾/.test(text)) category = 'eco';
      else if (/牽|扶|陪|大帶小|學弟|學妹|弟|妹|低年級/.test(text)) category = 'cross';
      else if (/喊|叫|呼|油|鼓掌|掌聲/.test(text)) category = 'cheer';
      else category = 'sports';
    }

    // 4. 清洗事蹟文字（徹底剔除隊名前綴、加分指令、見證人、地點等贅字）
    let story = text;
    story = story.replace(/(?:登記人|見證人|見證者|通報人|認證老師|裁判|紀錄人|記錄人|見證|通報)[：:\s]+[^\s,，、。\n\r]+/g, '');
    story = story.replace(/(?:見證)?(?:地點|位置)(?:在|[：:\s]+)+[^\s,，、。\n\r]+/g, '');
    if (author) {
      story = story.replace(new RegExp(author.replace('校長', '').trim(), 'g'), '');
    }

    // 循環剔除開頭的「各隊/各隊精神加分/紅隊精神加一分/紅綠藍隊各加1分」等前綴
    const teamPrefixRegex = /^(?:[🔴🟢🔵]\s*)?(?:(?:紅|綠|藍|金獅|烈火|翠玉|蜈蚣|跳鼓|雷鳴|碧海)(?:隊)?\s*(?:精神)?\s*(?:各|都)?\s*(?:加|得|[+＋])\s*[1-9一二兩三\d]\s*分?[，,、。\s]*)+/i;
    const allTeamPrefixRegex = /^(?:[🔴🟢🔵]\s*)?(?:紅\s*綠\s*藍|三隊|各隊|每隊|全隊|三大戰隊|全部戰隊)(?:隊)?\s*(?:都|各)?\s*(?:精神(?:加分|分|加)?|各加|各得|都加|加分|加|得|[+＋加]\s*[1-9一二兩三\d]\s*分?|[1-9一二兩三\d]\s*分|[，,、。\s\+＋])+/i;

    let prevStory = '';
    while (story !== prevStory) {
      prevStory = story;
      story = story.replace(teamPrefixRegex, '').trim();
      story = story.replace(allTeamPrefixRegex, '').trim();
      story = story.replace(/^[，,、。\s\+＋\-=]+|[，,、。\s]+$/g, '').trim();
    }
    if (!story) story = text;

    // 5. 智慧下標題（自動提煉 8~16 字精美大標題）
    const title = generateStoryTitle(story, category);

    return {
      intent: 'ADD_SPIRIT',
      teams: teams,
      team: teams[0],
      category: category,
      points: pts,
      title: title,
      story: story,
      author: author || '大會裁判',
      place: place
    };
  }

  // B. 團體賽事與慢跑項目匹配（支援直接報分/人數 或 名次登記，優先於個人短跑）
  if (/躲避球|拔河|接力|趣味競賽|蟲蟲危機|彩虹危機|慢跑|大隊接力/.test(text)) {
    let event = 'dodge';
    if (/慢跑/.test(text)) {
      if (/低|[一二12]年級/.test(text)) event = 'runLow';
      else if (/中|[三四34]年級/.test(text)) event = 'runMid';
      else if (/高|[五六56]年級/.test(text)) event = 'runHigh';
      else event = 'runLow';
    } else if (/拔河/.test(text)) {
      if (/高|[四五六456]年級/.test(text)) event = 'tugHigh';
      else event = 'tugLow';
    } else if (/接力/.test(text)) {
      if (/高|[四五六456]年級/.test(text)) event = 'relayHigh';
      else event = 'relayLow';
    } else if (/蟲蟲|彩虹|趣味/.test(text)) {
      event = 'rainbow';
    } else if (/躲避球/.test(text)) {
      event = 'dodge';
    }

    // 🛡️ 戰隊名稱防呆檢核（防止打錯為黃隊、白隊、黑隊等非本校戰隊）
    const wrongTeamMatch = text.match(/([黃白黑紫橘粉灰棕橙])(?:隊)?/);
    if (wrongTeamMatch) {
      return {
        intent: 'INVALID_TEAM',
        message: `⚠️【隊伍名稱輸入有誤：${wrongTeamMatch[0]}】\n南興國小僅有「🔴 紅隊」、「🟢 綠隊」、「🔵 藍隊」三大戰隊。\n您輸入的訊息包含非本校戰隊，本次成績未登錄，請修正後重新輸入！`
      };
    }

    const isHeadcount = /人/.test(text) || event.startsWith('run');

    const parseTeam = (str) => {
      if (!str) return null;
      if (/綠|翠玉|蜈蚣/.test(str)) return 'green';
      if (/藍|跳鼓|雷鳴|碧海/.test(str)) return 'blue';
      if (/紅|烈火|金獅/.test(str)) return 'red';
      return 'INVALID';
    };

    // 1. 優先檢查名次模式（支援「名次在前」與「隊伍在前」雙向自然語序）
    let t1 = null, t2 = null, t3 = null;

    // 語序 A：名次在前（例如：第一名藍隊、第一名：藍隊、冠軍藍隊、第1名藍隊）
    const m1Prefix = text.match(/(?:第一名|第1名|第一|第1|冠軍)[：:\s]*(?:為|是)?\s*([^\s,，、。\n\r]+)/);
    const m2Prefix = text.match(/(?:第二名|第2名|第二|第2|亞軍)[：:\s]*(?:為|是)?\s*([^\s,，、。\n\r]+)/);
    const m3Prefix = text.match(/(?:第三名|第3名|第三|第3|季軍)[：:\s]*(?:為|是)?\s*([^\s,，、。\n\r]+)/);

    // 只有當 m1 抓取出的字詞包含隊伍特徵（如包含「隊」或隊名關鍵字）時才判定為語序 A
    if (m1Prefix && /隊|紅|綠|藍|翠|碧|烈|蜈蚣|金獅|跳鼓|雷鳴/.test(m1Prefix[1])) {
      t1 = parseTeam(m1Prefix[1]);
      if (m2Prefix) t2 = parseTeam(m2Prefix[1]);
      if (m3Prefix) t3 = parseTeam(m3Prefix[1]);

      if (t1 === 'INVALID' || t2 === 'INVALID' || t3 === 'INVALID') {
        return {
          intent: 'INVALID_TEAM',
          message: `⚠️【名次隊伍名稱有誤】\n南興國小僅有「🔴 紅隊」、「🟢 綠隊」、「🔵 藍隊」三大戰隊。\n您輸入的名次包含非本校戰隊，本次成績未登錄，請修正後重新輸入！`
        };
      }
    }

    // 語序 B：隊伍在前（例如：紅隊第一、紅隊第一名、紅隊第1、紅隊冠軍、紅第一）
    if (!t1 || t1 === 'INVALID') {
      const redM = text.match(/(?:紅|烈火|金獅)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[一1]名?|冠軍)/);
      const greenM = text.match(/(?:綠|翠玉|蜈蚣)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[一1]名?|冠軍)/);
      const blueM = text.match(/(?:藍|跳鼓|雷鳴|碧海)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[一1]名?|冠軍)/);

      if (redM) t1 = 'red';
      else if (greenM) t1 = 'green';
      else if (blueM) t1 = 'blue';

      const red2 = text.match(/(?:紅|烈火|金獅)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[二2兩]名?|亞軍)/);
      const green2 = text.match(/(?:綠|翠玉|蜈蚣)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[二2兩]名?|亞軍)/);
      const blue2 = text.match(/(?:藍|跳鼓|雷鳴|碧海)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[二2兩]名?|亞軍)/);

      if (red2) t2 = 'red';
      else if (green2) t2 = 'green';
      else if (blue2) t2 = 'blue';

      const red3 = text.match(/(?:紅|烈火|金獅)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[三3]名?|季軍)/);
      const green3 = text.match(/(?:綠|翠玉|蜈蚣)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[三3]名?|季軍)/);
      const blue3 = text.match(/(?:藍|跳鼓|雷鳴|碧海)(?:隊)?\s*(?:是|為|得|拿)?[：:\s]*(?:第[三3]名?|季軍)/);

      if (red3) t3 = 'red';
      else if (green3) t3 = 'green';
      else if (blue3) t3 = 'blue';
    }

    // 🛡️ 名次重複檢核防呆
    if ((t1 && t2 && t1 === t2) || (t1 && t3 && t1 === t3) || (t2 && t3 && t2 === t3)) {
      return {
        intent: 'INVALID_TEAM',
        message: `⚠️【名次隊伍重複】\n各名次戰隊不可重複！\n本次成績未登錄，請確認第一名、第二名、第三名之戰隊後重新輸入！`
      };
    }

    if (t1) {
      const allTeams = ['green', 'blue', 'red'];
      if (!t2) {
        const left = allTeams.filter(t => t !== t1);
        t2 = left[0];
      }
      if (!t3) {
        const left = allTeams.filter(t => t !== t1 && t !== t2);
        t3 = left[0] || 'red';
      }

      return {
        intent: 'UPDATE_EVENT',
        event: event,
        rank1: t1,
        rank2: t2,
        rank3: t3
      };
    }

    // 2. 檢查是否有直接報得分或人數（例如：綠隊8分、藍隊6分、紅隊4分 或 藍7人、綠6人、紅5人）
    // 嚴格強制必須包含「分」或「人」單位，避免將條列序號「2. 第二名」誤判為「2分」
    const greenM = text.match(/(?:綠|翠玉|蜈蚣)(?:隊)?\s*(?:得|為|有)?\s*(\d+)\s*(?:分|人)/);
    const blueM = text.match(/(?:藍|跳鼓|雷鳴|碧海)(?:隊)?\s*(?:得|為|有)?\s*(\d+)\s*(?:分|人)/);
    const redM = text.match(/(?:紅|烈火|金獅)(?:隊)?\s*(?:得|為|有)?\s*(\d+)\s*(?:分|人)/);

    if (greenM || blueM || redM) {
      return {
        intent: 'UPDATE_EVENT',
        event: event,
        scores: {
          green: greenM ? parseInt(greenM[1], 10) : null,
          blue: blueM ? parseInt(blueM[1], 10) : null,
          red: redM ? parseInt(redM[1], 10) : null
        },
        isHeadcount: isHeadcount
      };
    }
  }

  // C. 個人短跑決賽名次匹配
  if (/短跑|60m|80m|100m|第一名|金牌|決賽|男生組|女生組/.test(text)) {
    let group = 'highBoy';
    if (/低.*男|1.*男|2.*男/.test(text)) group = 'lowBoy';
    else if (/低.*女|1.*女|2.*女/.test(text)) group = 'lowGirl';
    else if (/中.*男|3.*男|4.*男/.test(text)) group = 'midBoy';
    else if (/中.*女|3.*女|4.*女/.test(text)) group = 'midGirl';
    else if (/高.*男|5.*男|6.*男/.test(text)) group = 'highBoy';
    else if (/高.*女|5.*女|6.*女/.test(text)) group = 'highGirl';

    const m1 = text.match(/(?:第一名|金牌)[：:\s]*([^\s,，、]+)/);
    const m2 = text.match(/(?:第二名|銀牌)[：:\s]*([^\s,，、]+)/);
    const m3 = text.match(/(?:第三名|銅牌)[：:\s]*([^\s,，、]+)/);

    if (m1 || m2 || m3) {
      return {
        intent: 'UPDATE_SPRINT',
        group: group,
        gold: m1 ? m1[1] : '待定',
        silver: m2 ? m2[1] : '待定',
        bronze: m3 ? m3[1] : '待定'
      };
    }
  }

  return null;
}

/**
 * ----------------------------------------------------------------------
 * 4. 呼叫 Gemini 進行意圖分析與結構化 JSON 輸出（雙引擎自動降級備援）
 * ----------------------------------------------------------------------
 */
function callGeminiParser(userText, userId) {
  // ⚡ 優先級 1：使用本地超高速規則引擎（0 毫秒、零 API 成本、100% 精準）
  const localRes = fastLocalParse(userText);
  if (localRes) {
    return localRes;
  }

  // 🧠 優先級 2：口語隨性表達交給 Gemini Flash 進行多模態與語意解析
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;

  const systemInstruction = `
你是一位專門為「臺南市安南區南興國小 115 學年度校慶運動大會」服務的 AI 賽事秘書。
你的任務是將校長或老師輸入的隨性文字或語音轉錄內容，解析為嚴格的結構化 JSON。

學校賽事基本定義：
1. 戰隊：'green' (綠隊/翠玉蜈蚣), 'blue' (藍隊/雷鳴跳鼓), 'red' (紅隊/烈火金獅)。
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
- intent: "ADD_SPIRIT" (例如：綠隊精神加分2分，六年級幫學弟綁鞋帶；或 紅綠藍三隊各加1分，認真聽講規則)
  輸出欄位: { "intent": "ADD_SPIRIT", "teams": ["green"] 或 ["red","green","blue"], "category": "cross"|"sports"|"eco"|"cheer", "points": 1~3, "story": "完整事蹟全文", "author": "老師名稱若有提及否則填寫裁判", "place": "操場/風雨球場/小禮堂等" }

- intent: "UPDATE_SPRINT" (例如：高年級男生組第一名六年級陳小明、第二名林大華、第三名張小傑)
  輸出欄位: { "intent": "UPDATE_SPRINT", "group": "lowBoy"|"lowGirl"|"midBoy"|"midGirl"|"highBoy"|"highGirl", "gold": "姓名", "silver": "姓名", "bronze": "姓名" }

- intent: "UPDATE_EVENT" (例如：拔河低年級組綠隊8分、藍隊6分、紅隊4分 或 慢跑低年級藍隊7人、綠隊6人、紅隊5人 或 躲避球第一名綠隊、第二名紅隊、第三名藍隊)
  賽事代碼 event 必須為下列之一：
  - 'dodge' (熱血美式躲避球)
  - 'runLow' (校園慢跑低年級組)
  - 'runMid' (校園慢跑中年級組)
  - 'runHigh' (校園慢跑高年級組)
  - 'tugLow' (拔河賽一二三年級組)
  - 'tugHigh' (拔河賽四五六年級組)
  - 'rainbow' (趣味競賽蟲蟲危機)
  - 'relayLow' (大隊接力一二三年級組)
  - 'relayHigh' (大隊接力四五六年級組)
  輸出欄位可為：
  - 直接報得分或人數：{ "intent": "UPDATE_EVENT", "event": "tugLow", "scores": { "green": 8, "blue": 6, "red": 4 }, "isHeadcount": false }
  - 報名次：{ "intent": "UPDATE_EVENT", "event": "dodge", "rank1": "green", "rank2": "red", "rank3": "blue" }

必須只回傳合法的純 JSON，絕不加 markdown 標籤（如 \`\`\`json ）。
`;

  const modelsToTry = [CONFIG.GEMINI_MODEL, 'gemini-1.5-flash', 'gemini-2.0-flash'];

  for (const model of modelsToTry) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const payload = {
        contents: [{ role: 'user', parts: [{ text: userText }] }],
        systemInstruction: { parts: [{ text: systemInstruction }] },
        generationConfig: { temperature: 0.1, responseMimeType: "application/json" }
      };
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts[0].text) {
          let raw = json.candidates[0].content.parts[0].text.trim();
          raw = raw.replace(/^```json/i, '').replace(/```$/i, '').trim();
          const parsedGemini = JSON.parse(raw);
          if (parsedGemini && parsedGemini.intent === 'ADD_SPIRIT') {
            if (!parsedGemini.teams && parsedGemini.team) {
              parsedGemini.teams = [parsedGemini.team];
            }
          }
          return parsedGemini;
        }
      }
    } catch (e) {
      console.warn(`Gemini model ${model} failed:`, e);
    }
  }

  return null;
}

/**
 * 試算表公式注入與剖析防禦（徹底根除 Google Sheets #ERROR! 公式剖析錯誤）
 * 當文字開頭為 = + - @ 等符號時，主動補上單引號轉為強制純文字
 */
function safeSheetText(str) {
  if (str === null || str === undefined) return '';
  str = String(str).trim();
  if (/^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

/**
 * ----------------------------------------------------------------------
 * 5. 寫入資料庫：精神加分記錄
 * ----------------------------------------------------------------------
 */
function recordSpiritPoint(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders(ss);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);

  const teamNameMap = { green: '🟢 綠隊', blue: '🔵 藍隊', red: '🔴 紅隊' };
  const catNameMap = { cross: '🤝 跨齡扶持', sports: '🏃 運動家風度', eco: '🌿 環保自律', cheer: '📣 團隊歡呼' };

  const now = new Date();
  const cleanStory = safeSheetText(p.story);

  // 支援單隊或多隊（紅綠藍同時各加分）
  const targetTeams = (p.teams && p.teams.length > 0) ? p.teams : [p.team || 'green'];
  const appendedRows = [];

  for (const t of targetTeams) {
    const row = [
      now,
      teamNameMap[t] || t,
      t,
      catNameMap[p.category] || p.category,
      p.points || 1,
      cleanStory,
      safeSheetText(p.author || '大會裁判'),
      safeSheetText(p.place || '')
    ];
    sheet.appendRow(row);
    appendedRows.push(sheet.getLastRow());
  }

  // 暫存本次新增的列數（逗號隔開），供 /undo 一次全部撤銷
  PropertiesService.getScriptProperties().setProperty('LAST_SPIRIT_ROW', appendedRows.join(','));

  const titleText = p.title || generateStoryTitle(p.story, p.category);
  const placeTag = p.place ? ` (${p.place})` : '';

  const teamDisplay = targetTeams.map(t => teamNameMap[t] || t).join('、');
  const pointsTag = targetTeams.length > 1 ? `各 +${p.points || 1}分` : `+${p.points || 1}分`;

  return `🎉【精神加分登記成功！】\n` +
         `━━━━━━━━━━━━━━\n` +
         `戰隊：${teamDisplay} (${pointsTag})\n` +
         `向度：${catNameMap[p.category] || p.category}\n` +
         `標題：${titleText}\n` +
         `事蹟：${p.story}\n` +
         `見證：${p.author || '大會裁判'}${placeTag}\n` +
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

  const rows = lastRowStr.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n)).sort((a, b) => b - a);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPIRIT);

  let deletedCount = 0;
  let sampleData = null;

  for (const r of rows) {
    if (r > 1 && r <= sheet.getLastRow()) {
      if (!sampleData) {
        sampleData = sheet.getRange(r, 1, 1, 8).getValues()[0];
      }
      sheet.deleteRow(r);
      deletedCount++;
    }
  }

  scriptProps.deleteProperty('LAST_SPIRIT_ROW');

  if (deletedCount > 0) {
    const multiTag = deletedCount > 1 ? `（共撤銷 ${deletedCount} 筆戰隊紀錄）` : '';
    return `↩️【已成功撤銷精神加分】${multiTag}\n事蹟：${sampleData ? sampleData[5] : ''}\n官網資料已即時更正！`;
  }
  return '撤銷失敗，該列已不存在或已經撤銷過。';
}

/**
 * ----------------------------------------------------------------------
 * 6. 寫入資料庫：個人短跑決賽名次 (recordSprintWinner)
 * ----------------------------------------------------------------------
 */
function recordSprintWinner(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders(ss);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SPRINT);
  const data = sheet.getDataRange().getValues();

  const groupNameMap = {
    lowBoy: '低年級男生組 (60m)',
    lowGirl: '低年級女生組 (60m)',
    midBoy: '中年級男生組 (80m)',
    midGirl: '中年級女生組 (80m)',
    highBoy: '高年級男生組 (100m)',
    highGirl: '高年級女生組 (100m)'
  };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === p.group) {
      if (p.gold && p.gold !== '待定') sheet.getRange(i + 1, 4).setValue(p.gold);
      if (p.silver && p.silver !== '待定') sheet.getRange(i + 1, 5).setValue(p.silver);
      if (p.bronze && p.bronze !== '待定') sheet.getRange(i + 1, 6).setValue(p.bronze);
      sheet.getRange(i + 1, 7).setValue('裁判登記');
      sheet.getRange(i + 1, 8).setValue(new Date());
      break;
    }
  }

  const grpTitle = groupNameMap[p.group] || p.group;
  return `🏃【個人短跑決賽名次已更新！】\n` +
         `━━━━━━━━━━━━━━\n` +
         `組別：${grpTitle}\n` +
         `🥇 金牌：${p.gold || '尚未決賽'}\n` +
         `🥈 銀牌：${p.silver || '尚未決賽'}\n` +
         `🥉 銅牌：${p.bronze || '尚未決賽'}\n` +
         `━━━━━━━━━━━━━━\n` +
         `💡 官網各項賽事成績公告榜已即時同步！`;
}

/**
 * ----------------------------------------------------------------------
 * 7. 寫入資料庫：團體賽事得分 (recordEventScores)
 * ----------------------------------------------------------------------
 */
function recordEventScores(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders(ss);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVENTS);
  const data = sheet.getDataRange().getValues();

  const teamNameMap = { green: '綠隊', blue: '藍隊', red: '紅隊' };
  let eventName = p.event;
  let targetRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === p.event) {
      targetRow = i + 1;
      eventName = data[i][1];
      break;
    }
  }

  if (targetRow === -1) {
    return `⚠️ 找不到代碼為 ${p.event} 的賽事項目。請確認項目名稱。`;
  }

  // 取得原列資料
  const currentRowData = sheet.getRange(targetRow, 1, 1, 9).getValues()[0];
  let curGreen = Number(currentRowData[3]) || 0;
  let curBlue = Number(currentRowData[4]) || 0;
  let curRed = Number(currentRowData[5]) || 0;

  let newGreen = curGreen;
  let newBlue = curBlue;
  let newRed = curRed;
  let memo = '';
  let detailsText = '';

  // 模式 A：直接指定三隊得分或達標人數 (p.scores)
  if (p.scores) {
    if (p.scores.green !== undefined && p.scores.green !== null) newGreen = Number(p.scores.green);
    if (p.scores.blue !== undefined && p.scores.blue !== null) newBlue = Number(p.scores.blue);
    if (p.scores.red !== undefined && p.scores.red !== null) newRed = Number(p.scores.red);

    const unit = p.isHeadcount ? '人' : '分';
    memo = `綠隊 ${newGreen} ${unit}、藍隊 ${newBlue} ${unit}、紅隊 ${newRed} ${unit}`;
    detailsText = `🟢 綠隊：${newGreen} ${unit}\n🔵 藍隊：${newBlue} ${unit}\n🔴 紅隊：${newRed} ${unit}`;
  } 
  // 模式 B：以第一名、第二名、第三名名次登記 (p.rank1)
  else if (p.rank1) {
    // 依據 115 學年度校慶實施計畫暨工作手冊（最新規範）競賽積分規範：
    // 1. 全員趣味賽（如：蟲蟲危機 rainbow 等）及 熱血躲避球賽（dodge）：
    //    採【第一名 16 分、第二名 12 分、第三名 8 分】
    // 2. 分一二三年級組、四五六年級組各自對抗的積分賽（分組拔河 tugLow/tugHigh、分組大隊接力 relayLow/relayHigh 等）：
    //    採【第一名 8 分、第二名 6 分、第三名 4 分】
    const isHighPoints = (p.event === 'rainbow' || p.event === 'dodge' || p.event === 'allTogether');

    const pts1 = isHighPoints ? 16 : 8;
    const pts2 = isHighPoints ? 12 : 6;
    const pts3 = isHighPoints ? 8 : 4;

    const scores = { green: 0, blue: 0, red: 0 };
    if (p.rank1) scores[p.rank1] = pts1;
    if (p.rank2) scores[p.rank2] = pts2;
    if (p.rank3) scores[p.rank3] = pts3;

    newGreen = scores.green;
    newBlue = scores.blue;
    newRed = scores.red;
    memo = `冠軍${teamNameMap[p.rank1]}(+${pts1})、亞軍${teamNameMap[p.rank2] || '藍隊'}(+${pts2})、季軍${teamNameMap[p.rank3] || '紅隊'}(+${pts3})`;
    detailsText = `🥇 冠軍：${teamNameMap[p.rank1]} (+${pts1}分)\n🥈 亞軍：${teamNameMap[p.rank2] || '藍隊'} (+${pts2}分)\n🥉 季軍：${teamNameMap[p.rank3] || '紅隊'} (+${pts3}分)`;
  }

  // 寫入 Google 試算表 D, E, F, G, H, I 欄
  sheet.getRange(targetRow, 4).setValue(newGreen);
  sheet.getRange(targetRow, 5).setValue(newBlue);
  sheet.getRange(targetRow, 6).setValue(newRed);
  sheet.getRange(targetRow, 7).setValue(memo);
  sheet.getRange(targetRow, 8).setValue(p.author || '裁判登記 (LINE)');
  sheet.getRange(targetRow, 9).setValue(new Date());

  return `🏆【賽事項目成績已成功登錄！】\n` +
         `━━━━━━━━━━━━━━\n` +
         `賽事：${eventName}\n` +
         `${detailsText}\n` +
         `備註：${memo}\n` +
         `━━━━━━━━━━━━━━\n` +
         `💡 雲端試算表與官網戰隊總積分榜已即時同步更新！`;
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

    // 寫入預設賽事架構（依 115 校慶工作手冊定稿標準計分：全員趣味 16/12/8；分組對抗 8/6/4）
    const defaultEvents = [
      ['11/26 (四)', '熱血美式躲避球循環賽', 'dodge', 16, 8, 12, '綠隊冠軍(+16)、紅隊亞軍(+12)、藍隊季軍(+8)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-低年級3圈(5分內)', 'runLow', 6, 7, 5, '藍7人、綠6人、紅5人達標(每人1分)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-中年級800m(7分內)', 'runMid', 6, 4, 5, '綠6人、紅5人、藍4人達標(每人1分)', '大會裁判組', new Date()],
      ['12/01 (二)', '校園慢跑-高年級800m(6分內)', 'runHigh', 5, 4, 6, '紅6人、綠5人、藍4人達標(每人1分)', '大會裁判組', new Date()],
      ['12/05 (六)', '拔河賽-一二三年級組', 'tugLow', 8, 6, 4, '綠隊冠軍(+8)、藍隊亞軍(+6)、紅隊季軍(+4)', '大會裁判組', new Date()],
      ['12/05 (六)', '拔河賽-四五六年級組', 'tugHigh', 6, 4, 8, '紅隊冠軍(+8)、綠隊亞軍(+6)、藍隊季軍(+4)', '大會裁判組', new Date()],
      ['12/05 (六)', '趣味競賽-蟲蟲危機', 'rainbow', 8, 16, 12, '藍隊冠軍(+16)、紅隊亞軍(+12)、綠隊季軍(+8)', '大會裁判組', new Date()],
      ['12/05 (六)', '大隊接力-一二三年級組', 'relayLow', 6, 8, 4, '藍隊冠軍(+8)、綠隊亞軍(+6)、紅隊季軍(+4)', '大會裁判組', new Date()],
      ['12/05 (六)', '大隊接力-四五六年級組', 'relayHigh', 8, 4, 6, '綠隊冠軍(+8)、紅隊亞軍(+6)、藍隊季軍(+4)', '大會裁判組', new Date()]
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
  if (/綠|翠玉|蜈蚣/.test(t)) return 'green';
  if (/藍|跳鼓|雷鳴|碧海/.test(t)) return 'blue';
  if (/紅|烈火|金獅/.test(t)) return 'red';
  return 'green';
}

/**
 * ----------------------------------------------------------------------
 * 智慧精神事蹟標題提煉器 (自動為事蹟下優雅動人的精準大標題)
 * ----------------------------------------------------------------------
 */
function generateStoryTitle(story, category) {
  if (!story) return '溫馨校園互助事蹟';

  // 1. 環保自律向度
  if (category === 'eco' || /清潔|環境|垃圾|落葉|掃地|整齊|收拾|排隊|水壺|自律|整潔|回收|紙屑|整理|場地|復原|打掃/.test(story)) {
    if (/整理.*場地|場地.*整理|場地.*復原|復原.*場地|整理場地/.test(story)) return '合力整理場地，展現高度自律品德';
    if (/清潔|環境|打掃|掃地|做整潔/.test(story)) return '主動維護整潔，守護美麗校園';
    if (/垃圾|落葉|撿|紙屑|寶特瓶/.test(story)) return '隨手彎腰拾起，美化綠活校園';
    if (/排隊|水壺|整齊|收拾|常規|器材/.test(story)) return '自律守規井然，落實生活常規';
    return '自主維護環境，實踐綠活校園';
  }

  // 2. 跨齡扶持向度
  if (category === 'cross' || /大帶小|小隊長|一年級|二年級|學弟|學妹|低年級|牽|陪|抱|外套|跌倒|鞋帶/.test(story)) {
    if (/外套|衣服|穿|保暖/.test(story)) return '大手貼心遞外套，溫暖學弟學妹';
    if (/鞋帶|綁/.test(story)) return '蹲下身細心綁鞋帶，展現大帶小溫情';
    if (/跌倒|扶|牽手|牽|抱/.test(story)) return '大手牽小手，暖心扶持齊步走';
    if (/檢錄|排隊|帶隊|帶領/.test(story)) return '熱心帶領檢錄，小隊長盡責典範';
    return '跨齡互助扶持，大手牽小手展溫馨';
  }

  // 3. 運動家風度向度
  if (category === 'sports' || /風度|運動家|運動員|禮貌|敬禮|握手|擁抱|安慰|服從|擊掌|拍肩|勝不驕|敗不餒|輸了|落後|勉勵|彼此勉勵|互勉/.test(story)) {
    if (/輸了|落後|不氣餒|勉勵|彼此勉勵|互勉/.test(story)) return '落後不氣餒，彼此勉勵展風度';
    if (/握手|鞠躬|敬禮|致意/.test(story)) return '賽後致敬對手，展現運動家風度';
    if (/拍肩|安慰|鼓勵|打氣/.test(story)) return '拍肩暖心打氣，超越勝負的真情';
    if (/跌倒|扶對手|受傷/.test(story)) return '扶起競賽對手，發揚運動家品格';
    return '勝不驕敗不餒，盡顯運動家典範';
  }

  // 4. 團隊歡呼向度
  if (category === 'cheer' || /歡呼|隊呼|加油|吶喊|士氣|掌聲|口號|鼓掌|喝采|熱情/.test(story)) {
    if (/隊呼|口號/.test(story)) return '戰隊呼響徹雲霄，凝聚最強向心力';
    if (/加油|吶喊|掌聲|喝采/.test(story)) return '熱情吶喊加油，點燃全場澎湃士氣';
    return '齊心歡呼鼓舞，展現團隊最高士氣';
  }

  // 預設擷取精華
  return story.length > 16 ? story.substring(0, 16) + '...' : story;
}
