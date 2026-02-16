// =================================================================
//                        全域設定與常數
// =================================================================
const FORM_RESPONSES_SHEET_NAME = '表單回應 1'; 
const STUDENTS_SHEET_NAME = 'Students';
const SEAT_VIEW_SHEET_NAME = 'Seat_View';
const CHOICES_SHEET_NAME = 'Round_now_Choices';
const LOTTERY_RESULT_SHEET_NAME = 'Lottery_Result';
const HISTORY_LOG_SHEET_NAME = 'History_Log'; // 歷史存檔工作表
const CONFIG_SHEET_NAME = 'Config';           // 系統設定工作表

// =================================================================
//                        Web App 路由與入口
// =================================================================

function doGet(e) {
  const page = e.parameter.page;
  
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('管理後台 - 座位選位系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('學生登入 - 座位選位系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// =================================================================
//                        系統設定與狀態管理
// =================================================================

function getSystemStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.getRange('A1:B1').setValues([['Key', 'Value']]);
    configSheet.getRange('A2:B2').setValues([['isOpen', 'FALSE']]);
    SpreadsheetApp.flush();
  }
  
  const data = configSheet.getRange('A2:B2').getValues();
  const isOpen = String(data[0][1]).toUpperCase() === 'TRUE'; 
  
  return {
    isOpen: isOpen,
    webAppUrl: ScriptApp.getService().getUrl()
  };
}

function setSystemStatus(isOpen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.getRange('A1:B1').setValues([['Key', 'Value']]);
  }
  configSheet.getRange('A2:B2').setValues([['isOpen', isOpen ? 'TRUE' : 'FALSE']]);
  SpreadsheetApp.flush(); 
  return { success: true, isOpen: isOpen };
}

// =================================================================
//                        身分驗證與登入
// =================================================================

function loginStudent(seatNo, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    const lastRow = studentSheet.getLastRow();
    const data = studentSheet.getRange('A2:F' + lastRow).getValues();
    
    let foundStudent = null;
    for (let i = 0; i < data.length; i++) {
      const rowSeat = String(data[i][0]).trim();
      const rowName = data[i][1];
      const rowPass = String(data[i][5]).trim();

      if (rowSeat === String(seatNo).trim() && rowPass.toUpperCase() === String(password).trim().toUpperCase()) {
        foundStudent = { seatNo: rowSeat, name: rowName };
        break;
      }
    }

    if (foundStudent) {
      return { success: true, student: foundStudent };
    } else {
      return { success: false, message: '座號或密碼錯誤' };
    }
  } catch (e) {
    return { success: false, message: '系統錯誤: ' + e.message };
  }
}

// =================================================================
//                        座位資料讀取與提交
// =================================================================

function getSeatMapData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatViewSheet = ss.getSheetByName(SEAT_VIEW_SHEET_NAME);
    const range = seatViewSheet.getRange('B4:L20'); 
    const rawData = range.getValues(); 
    const seatMap = {};

    for (let i = 0; i < rawData.length; i += 3) {
      if (i + 1 >= rawData.length) continue; 
      const rowName = rawData[i];     
      const rowStatus = rawData[i + 1]; 
      
      for (let j = 0; j < rowName.length; j++) {
        const seatId = String(rowName[j]).trim();
        const status = rowStatus[j]; 

        if (seatId !== '' && seatId.toUpperCase() !== 'FALSE' && !seatId.includes('學年')) {
          let isOccupied = false;
          let occupantSeatId = null;
          let isAvailable = true;

          if (typeof status === 'number') {
            isOccupied = false;
            isAvailable = true; 
          } else if (String(status).length > 0) {
            isOccupied = true;
            isAvailable = false;
            occupantSeatId = String(status);
          }

          seatMap[seatId] = {
            isOccupied: isOccupied,
            isAvailable: isAvailable,
            occupantSeatId: occupantSeatId,
            rawStatus: status
          };
        }
      }
    }
    return { seatMap: seatMap };
  } catch (e) {
    return { error: e.message };
  }
}

function recordChoice(formData) {
  const { seatId, choice } = formData;
  if (!getSystemStatus().isOpen) return { success: false, message: '選位已截止或尚未開放' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responseSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME); 
    const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    
    const rowData = [new Date(), seatId, choice, choice];
    responseSheet.appendRow(rowData);

    if (stuSheet) {
      const lastRow = stuSheet.getLastRow();
      const range = stuSheet.getRange('A2:G' + lastRow);
      const data = range.getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(seatId).trim()) {
          data[i][6] = ''; 
          break;
        }
      }
      range.setValues(data);
    }
    return { success: true, seatId: seatId, choice: choice };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =================================================================
//                        核心抽籤邏輯
// =================================================================

function allocateSeatsByScore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const choicesSheet = ss.getSheetByName(CHOICES_SHEET_NAME);
  const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
  
  const stuDataRange = stuSheet.getRange('A2:G' + stuSheet.getLastRow());
  const stuData = stuDataRange.getValues();
  const scoreMap = {};
  const seatRowIndex = {};

  stuData.forEach((row, idx) => {
    const sNo = Number(row[0]);
    if (sNo) {
      scoreMap[sNo] = Number(row[4]) || 0;
      seatRowIndex[sNo] = idx;
    }
  });

  const choiceData = choicesSheet.getRange('A2:C' + choicesSheet.getLastRow()).getValues();
  const participants = new Set();
  const choiceMap = {};

  choiceData.forEach(row => {
    const sNo = Number(row[0]);
    const seatId = String(row[2]).trim().toUpperCase();
    if (sNo && seatId) {
      participants.add(sNo);
      if (!choiceMap[seatId]) choiceMap[seatId] = [];
      choiceMap[seatId].push({ seatNo: sNo, score: scoreMap[sNo] || 0 });
    }
  });

  const winnerMap = {};
  const winners = [];
  for (const seatId in choiceMap) {
    const contenders = choiceMap[seatId];
    const winnerSeatNo = determineWinner(contenders);
    winnerMap[winnerSeatNo] = seatId;
    winners.push([winnerSeatNo, seatId, contenders.length > 1 ? '競爭勝出' : '直接分配', new Date()]);
  }

  stuData.forEach((row, idx) => {
    const sNo = Number(row[0]);
    if (participants.has(sNo)) {
      if (winnerMap[sNo]) {
        stuData[idx][3] = winnerMap[sNo]; 
        stuData[idx][6] = 'SUCCESS';      
      } else {
        stuData[idx][6] = 'FAIL';
      }
    }
  });

  stuDataRange.setValues(stuData);
  
  const lotSheet = ss.getSheetByName(LOTTERY_RESULT_SHEET_NAME);
  if (lotSheet && winners.length > 0) {
    lotSheet.getRange(lotSheet.getLastRow() + 1, 1, winners.length, 4).setValues(winners);
  }

  return { success: true, count: winners.length };
}

function determineWinner(contenders) {
  let totalWeight = 0;
  contenders.forEach(c => totalWeight += Math.max(0, c.score));
  
  if (totalWeight <= 0) {
    return contenders[Math.floor(Math.random() * contenders.length)].seatNo;
  }

  let r = Math.random() * totalWeight;
  for (let c of contenders) {
    r -= c.score;
    if (r <= 0) return c.seatNo;
  }
  return contenders[contenders.length - 1].seatNo;
}

// =================================================================
//                        永久性存檔與重置
// =================================================================

function archiveOnly(roundName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
  const historySheet = ss.getSheetByName(HISTORY_LOG_SHEET_NAME) || ss.insertSheet(HISTORY_LOG_SHEET_NAME);
  
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(['時間', '輪次名稱', '座號', '姓名', '分配座位']);
  }

  const data = stuSheet.getRange('A2:D' + stuSheet.getLastRow()).getValues();
  const archiveData = [];
  const now = new Date();

  data.forEach(row => {
    if (row[3]) { 
      archiveData.push([now, roundName, row[0], row[1], row[3]]);
    }
  });

  if (archiveData.length > 0) {
    historySheet.getRange(historySheet.getLastRow() + 1, 1, archiveData.length, 5).setValues(archiveData);
    return { success: true, message: `成功存檔 ${archiveData.length} 筆紀錄！` };
  } else {
    return { success: false, message: '目前沒有分配完成的座位可供存檔' };
  }
}

function simulateMissingChoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
  const responseSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME);
  
  // 1. 抓取目前最新的座位圖，確保知道哪些位置「真的空著」
  const seatViewData = getSeatMapData().seatMap;
  const availableSeats = Object.keys(seatViewData).filter(id => !seatViewData[id].isOccupied);
  
  if (availableSeats.length === 0) return { success: false, message: '已無可用空位' };

  // 2. 抓取目前已經提交志願的名單，避免重複產生
  const currentResponses = responseSheet.getLastRow() > 1 ? 
    responseSheet.getRange('B2:B' + responseSheet.getLastRow()).getValues().map(r => String(r[0]).trim()) : [];
  const respondedSet = new Set(currentResponses);

  const stuData = stuSheet.getRange('A2:G' + stuSheet.getLastRow()).getValues();
  const missingStudents = [];
  
  stuData.forEach(row => {
    const seatNo = String(row[0]).trim();
    const finalSeat = String(row[3] || '').trim();
    
    // 條件：只要 A 欄有座號，且 D 欄是空的 (代表還沒拿到位置)
    // 且 B 欄沒有出現在表單回應中 (代表目前沒志願)
    if (seatNo && seatNo !== "座號" && (!finalSeat || finalSeat === "") && !respondedSet.has(seatNo)) {
      missingStudents.push(seatNo);
    }
  });

  if (missingStudents.length === 0) return { success: false, message: '目前沒有需要補全志願的學生' };

  const newRows = [];
  const now = new Date();
  missingStudents.forEach(sNo => {
    const randomSeat = availableSeats[Math.floor(Math.random() * availableSeats.length)];
    newRows.push([now, sNo, randomSeat, randomSeat]);
  });

  if (newRows.length > 0) {
    responseSheet.getRange(responseSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }

  return { success: true, message: `已幫 ${missingStudents.length} 位同學產生新志願 (包含上一輪未中籤者)` };
}

function resetSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
  const formSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME);
  
  if (stuSheet) {
    const lastRow = stuSheet.getLastRow();
    if (lastRow >= 2) {
      stuSheet.getRange('D2:D' + lastRow).clearContent();
      stuSheet.getRange('G2:G' + lastRow).clearContent();
    }
  }
  
  if (formSheet && formSheet.getLastRow() > 1) {
    formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).clearContent();
  }
  
  setSystemStatus(false);
  return { success: true, message: '系統已重置，所有人狀態已清空' };
}

function getMyRoundResult(seatNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
  const data = stuSheet.getRange('A2:G' + stuSheet.getLastRow()).getValues();
  const target = String(seatNo).trim();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === target) {
      const result = String(data[i][6] || '').toUpperCase();
      const finalSeat = data[i][3] || '';
      return { status: result || 'NONE', seat: finalSeat };
    }
  }
  return { status: 'NONE' };
}
