/**
 * 包裹異常追蹤表修復腳本
 * 解決問題：員工分頁結案後直接消失，無法同步到主表
 * 
 * 修復邏輯：
 * 1. 結案時先同步到主表，再隱藏/歸檔該列
 * 2. 不是直接刪除，而是標記為已處理
 */

// Google Sheets ID (請替換為你的實際ID)
const SHEET_ID = '1ZBF9LDxCQ28mCLFkXcCNoGu5M_ybTWKLjLgn6YSaaI0';
const MAIN_SHEET_NAME = '包裹異常-追蹤表';

/**
 * 修復結案流程的核心函數
 * @param {string} deptSheetName - 部門分頁名稱
 * @param {number} row - 要結案的行號
 */
function fixedCloseCase(deptSheetName, row) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const deptSheet = ss.getSheetByName(deptSheetName);
    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    
    if (!deptSheet || !mainSheet) {
      throw new Error(`找不到工作表: ${deptSheetName} 或 ${MAIN_SHEET_NAME}`);
    }
    
    // 1. 取得要結案的資料
    const caseData = deptSheet.getRange(row, 1, 1, deptSheet.getLastColumn()).getValues()[0];
    const caseId = caseData[0]; // 假設第一欄是案件ID
    
    if (!caseId) {
      throw new Error('案件ID不能為空');
    }
    
    // 2. 更新結案狀態和時間
    const now = new Date();
    const statusColumnIndex = getColumnIndex(deptSheet, '處理狀態'); // 假設有處理狀態欄
    const remarkColumnIndex = getColumnIndex(deptSheet, '備註');
    
    if (statusColumnIndex > 0) {
      deptSheet.getRange(row, statusColumnIndex).setValue('已結案');
    }
    
    // 3. 同步到主表 (關鍵修復)
    syncToMainSheet(caseId, caseData, mainSheet, now);
    
    // 4. 不直接刪除，而是隱藏該列或移到歸檔區
    hideOrArchiveRow(deptSheet, row, caseData);
    
    // 5. 記錄操作日誌
    console.log(`[修復] 案件 ${caseId} 已正確結案並同步到主表`);
    
    // 6. 發送通知郵件 (如果需要)
    sendCloseNotification(caseId, deptSheetName, caseData);
    
    return true;
    
  } catch (error) {
    console.error(`[錯誤] 結案處理失敗: ${error.message}`);
    // 發送錯誤通知給管理員
    sendErrorNotification(deptSheetName, row, error.message);
    return false;
  }
}

/**
 * 同步案件到主表
 */
function syncToMainSheet(caseId, caseData, mainSheet, closeTime) {
  try {
    // 尋找主表中的對應記錄
    const mainData = mainSheet.getDataRange().getValues();
    let foundRow = -1;
    
    for (let i = 1; i < mainData.length; i++) { // 從第2行開始 (跳過標題)
      if (mainData[i][0] == caseId) { // 假設第一欄是案件ID
        foundRow = i + 1; // 轉換為1-based行號
        break;
      }
    }
    
    if (foundRow > 0) {
      // 更新現有記錄
      const statusCol = getColumnIndex(mainSheet, '處理狀態');
      const closeTimeCol = getColumnIndex(mainSheet, '結案時間');
      const handlerCol = getColumnIndex(mainSheet, '處理人員');
      
      if (statusCol > 0) mainSheet.getRange(foundRow, statusCol).setValue('已結案');
      if (closeTimeCol > 0) mainSheet.getRange(foundRow, closeTimeCol).setValue(closeTime);
      if (handlerCol > 0) mainSheet.getRange(foundRow, handlerCol).setValue(Session.getActiveUser().getEmail());
      
      console.log(`[同步] 已更新主表第 ${foundRow} 行`);
    } else {
      // 新增記錄到主表 (理論上不應該發生，但作為備用)
      const newRowData = [...caseData];
      newRowData.push(closeTime, '已結案', Session.getActiveUser().getEmail());
      mainSheet.appendRow(newRowData);
      console.log(`[同步] 已新增記錄到主表`);
    }
    
  } catch (error) {
    throw new Error(`同步到主表失敗: ${error.message}`);
  }
}

/**
 * 隱藏或歸檔已結案的列，而不是直接刪除
 */
function hideOrArchiveRow(sheet, row, caseData) {
  try {
    // 方案1: 隱藏該列
    sheet.hideRows(row);
    
    // 方案2: 移動到歸檔工作表 (可選)
    // moveToArchiveSheet(sheet.getParent(), caseData);
    
    console.log(`[歸檔] 第 ${row} 行已隱藏`);
  } catch (error) {
    console.error(`[警告] 歸檔失敗，但主要流程已完成: ${error.message}`);
  }
}

/**
 * 取得欄位索引 (1-based)
 */
function getColumnIndex(sheet, columnName) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const index = headers.indexOf(columnName);
    return index >= 0 ? index + 1 : -1;
  } catch (error) {
    console.error(`取得欄位索引失敗: ${error.message}`);
    return -1;
  }
}

/**
 * 發送結案通知
 */
function sendCloseNotification(caseId, deptName, caseData) {
  try {
    const recipients = [
      'mona@company.com',    // mona
      'supervisor@company.com' // 主管
    ];
    
    const subject = `包裹異常案件已結案 - ${caseId}`;
    const body = `
親愛的同事：

案件編號：${caseId}
處理部門：${deptName}
結案時間：${new Date().toLocaleString('zh-TW')}
處理人員：${Session.getActiveUser().getEmail()}

案件已成功結案並同步到主追蹤表。

此為系統自動通知，請勿回覆。
    `;
    
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });
    
    console.log(`[通知] 結案通知已發送`);
  } catch (error) {
    console.error(`[警告] 通知發送失敗: ${error.message}`);
  }
}

/**
 * 發送錯誤通知給管理員
 */
function sendErrorNotification(deptName, row, errorMsg) {
  try {
    MailApp.sendEmail({
      to: 'admin@company.com', // 管理員郵箱
      subject: `包裹異常系統錯誤 - ${deptName}`,
      body: `
部門工作表：${deptName}
錯誤行號：${row}
錯誤訊息：${errorMsg}
發生時間：${new Date().toLocaleString('zh-TW')}

請檢查系統狀態。
      `
    });
  } catch (error) {
    console.error(`發送錯誤通知失敗: ${error.message}`);
  }
}

/**
 * 批次修復現有的問題案件
 */
function batchFixExistingCases() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    
    // 取得所有部門分頁
    const deptSheets = ['客服', '倉庫', '採購', '其他']; // 請根據實際部門名稱調整
    
    let fixedCount = 0;
    
    for (const deptName of deptSheets) {
      const deptSheet = ss.getSheetByName(deptName);
      if (!deptSheet) continue;
      
      const data = deptSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) { // 從第2行開始
        const caseId = data[i][0];
        if (!caseId) continue;
        
        // 檢查是否為未同步的已結案件
        if (data[i].some(cell => cell && cell.toString().includes('已結案'))) {
          // 重新同步到主表
          syncToMainSheet(caseId, data[i], mainSheet, new Date());
          fixedCount++;
        }
      }
    }
    
    console.log(`[批次修復] 已修復 ${fixedCount} 個案件`);
    return fixedCount;
    
  } catch (error) {
    console.error(`批次修復失敗: ${error.message}`);
    return 0;
  }
}

/**
 * 安裝修復後的觸發器
 */
function installFixedTriggers() {
  try {
    // 刪除舊的觸發器
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction().includes('onEdit')) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 安裝新的觸發器
    ScriptApp.newTrigger('onEditInstallable')
      .onEdit()
      .create();
    
    console.log('✅ 修復後的觸發器已安裝');
  } catch (error) {
    console.error(`觸發器安裝失敗: ${error.message}`);
  }
}

/**
 * 修復後的編輯觸發器
 */
function onEditInstallable(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const value = e.value;
    
    // 只處理部門分頁的結案操作
    if (['客服', '倉庫', '採購', '其他'].includes(sheetName)) {
      // 檢查是否為結案操作
      if (value && (value.includes('結案') || value.includes('完成'))) {
        const row = range.getRow();
        
        // 使用修復後的結案函數
        const success = fixedCloseCase(sheetName, row);
        
        if (success) {
          // 可以在這裡添加 UI 提示
          SpreadsheetApp.getUi().alert(`案件已成功結案並同步到主表`);
        } else {
          SpreadsheetApp.getUi().alert(`結案處理失敗，請聯絡管理員`);
        }
      }
    }
    
  } catch (error) {
    console.error(`編輯觸發器錯誤: ${error.message}`);
  }
}
