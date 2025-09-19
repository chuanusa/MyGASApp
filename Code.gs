
// ===============================================================
// 全域設定 (Global Settings)
// ===============================================================
const SPREADSHEET_ID = '1LnCYteh9p8IqtXPoidFl4eMPCKI922WbQqX5OLzx8tw';
const DRIVE_FOLDER_ID = '1E5h2fx361IWHm1br8sKQWzGulLDCcmD2';

const SHEET_MENU = '下拉選單';
const SHEET_TRAINED = '已受訓清單';
const SHEET_RECORD = '參加記錄';
const SHEET_FILE_LOG = '檔案上傳紀錄';
const SHEET_EDIT_LOG = '修改紀錄';
const SHEET_DELETE_LOG = '刪除紀錄';

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// ===============================================================
// 主要網頁服務 (Main Web App Service)
// ===============================================================

/**
 * 當使用者透過 GET 請求訪問網頁時執行
 * @param {object} e - 事件參數
 * @returns {HtmlOutput} - 渲染後的網頁
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.data = {
    // 如果需要，可以在此傳遞初始數據到前端
  };
  return template.evaluate()
    .setTitle('員工工安教育訓練報名系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 引入其他 HTML 檔案內容 (CSS, JavaScript)
 * @param {string} filename - 檔案名稱
 * @returns {string} - 檔案內容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===============================================================
// 資料讀取功能 (Data Reading Functions)
// ===============================================================

/**
 * 獲取所有唯一的部門列表
 * @returns {string[]} - 部門名稱陣列
 */
function getDepartments() {
  try {
    const sheet = ss.getSheetByName(SHEET_MENU);
    const range = sheet.getRange('B2:B' + sheet.getLastRow());
    const values = range.getValues().flat().filter(String);
    return [...new Set(values)]; // 使用 Set 去除重複項
  } catch (e) {
    console.error('getDepartments 失敗:', e);
    return [];
  }
}

/**
 * 根據部門獲取姓名列表
 * @param {string} department - 部門名稱
 * @returns {string[]} - 姓名陣列
 */
function getNamesByDepartment(department) {
  try {
    const sheet = ss.getSheetByName(SHEET_MENU);
    const data = sheet.getRange('B2:C' + sheet.getLastRow()).getValues();
    const names = data.filter(row => row[0] === department).map(row => row[1]);
    return names;
  } catch (e) {
    console.error('getNamesByDepartment 失敗:', e);
    return [];
  }
}

/**
 * 獲取指定員工的課程資訊 (可報名/已報名/已完成)
 * @param {string} name - 員工姓名
 * @returns {object} - 包含可報名與已參加課程的物件
 */
function getCoursesForUser(name) {
  try {
    const allCoursesSheet = ss.getSheetByName(SHEET_MENU);
    const trainedSheet = ss.getSheetByName(SHEET_TRAINED);
    const recordSheet = ss.getSheetByName(SHEET_RECORD);

    // 1. 獲取所有課程 (包含名稱與簡稱)
    const allCoursesData = allCoursesSheet.getRange('H2:I' + allCoursesSheet.getLastRow()).getValues();
    const allCourses = allCoursesData.filter(row => row[0]).map(row => ({ name: row[0], shortName: row[1] }));

    // 2. 獲取已受訓清單中的課程
    const trainedData = trainedSheet.getRange('B2:C' + trainedSheet.getLastRow()).getValues();
    const trainedCourses = trainedData.filter(row => row[1] === name).map(row => row[0]);

    // 3. 獲取已在 "參加記錄" 中報名的課程
    const recordData = recordSheet.getRange('C2:D' + recordSheet.getLastRow()).getValues();
    const registeredCourses = recordData.filter(row => row[0] === name).map(row => row[1]);
    
    const attendedCourseNames = [...new Set([...trainedCourses, ...registeredCourses])];
    
    // 從所有課程中篩選出可報名的
    const availableCourses = allCourses.filter(course => !attendedCourseNames.includes(course.name));

    return {
      available: availableCourses, // 回傳物件陣列 {name, shortName}
      attended: attendedCourseNames // 回傳名稱字串陣列
    };
  } catch (e) {
    console.error('getCoursesForUser 失敗:', e);
    return { available: [], attended: [] };
  }
}

// ===============================================================
// 資料寫入/修改功能 (Data Writing/Modification Functions)
// ===============================================================

/**
 * 提交報名資料
 * @param {object} data - 前端傳來的報名資料，包含 name, courses, modifier, reason
 * @returns {object} - 操作結果
 */
function submitRegistration(data) {
  try {
    const { name, courses, modifier, reason } = data;
    const recordSheet = ss.getSheetByName(SHEET_RECORD);
    const editLogSheet = ss.getSheetByName(SHEET_EDIT_LOG);
    const timestamp = new Date();

    // 找出此人舊的報名資料
    const existingRecords = findUserRecords(name);

    // 如果有新報名的課程
    if (courses && courses.length > 0) {
        const newRecords = [];
        courses.forEach(course => {
            const recordId = Utilities.getUuid();
            newRecords.push([
                recordId, '', name, course, '', timestamp, timestamp, modifier, reason
            ]);
        });

        // 寫入新的報名資料
        if (newRecords.length > 0) {
            recordSheet.getRange(recordSheet.getLastRow() + 1, 1, newRecords.length, newRecords[0].length).setValues(newRecords);
        }
    }
    
    // 處理修改與刪除邏輯
    const oldCourses = existingRecords.map(r => r.courseName);
    const newCourses = courses || [];

    // 找出被取消的課程
    const cancelledCourses = oldCourses.filter(c => !newCourses.includes(c));
    if (cancelledCourses.length > 0) {
      cancelCourses(name, cancelledCourses, modifier, reason);
    }

    // 記錄修改日誌
    if (existingRecords.length > 0 || newCourses.length > 0) {
       logEdit(name, oldCourses.join(', '), newCourses.join(', '), modifier, reason);
    }

    return { success: true, message: '報名資料已成功更新！' };
  } catch (e) {
    console.error('submitRegistration 失敗:', e);
    return { success: false, message: '伺服器發生錯誤，請稍後再試。' };
  }
}

/**
 * 輔助函數：尋找使用者現有的報名記錄
 */
function findUserRecords(name) {
    const recordSheet = ss.getSheetByName(SHEET_RECORD);
    if (recordSheet.getLastRow() < 2) return [];
    const data = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, 9).getValues();
    const records = [];
    data.forEach((row, index) => {
        if (row[2] === name) { // C欄是姓名
            records.push({
                rowIndex: index + 2,
                recordId: row[0],
                courseName: row[3]
            });
        }
    });
    return records;
}

/**
 * 輔助函數：取消課程 (刪除舊紀錄並寫入Log)
 */
function cancelCourses(name, coursesToCancel, modifier, reason) {
    const recordSheet = ss.getSheetByName(SHEET_RECORD);
    const deleteLogSheet = ss.getSheetByName(SHEET_DELETE_LOG);
    const data = recordSheet.getRange(1, 1, recordSheet.getLastRow(), recordSheet.getLastColumn()).getValues();
    const rowsToDelete = [];
    const logs = [];

    for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        if (row[2] === name && coursesToCancel.includes(row[3])) {
            rowsToDelete.push(i + 1);
            logs.push([
                new Date(), modifier, reason, ...row
            ]);
        }
    }

    if (logs.length > 0) {
        deleteLogSheet.getRange(deleteLogSheet.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);
    }

    // 從後往前刪除，避免 index 錯亂
    rowsToDelete.forEach(rowIndex => {
        recordSheet.deleteRow(rowIndex);
    });
}

/**
 * 輔助函數：記錄修改日誌
 */
function logEdit(name, before, after, modifier, reason) {
    if (before === after) return; // 內容沒變，不記錄
    const editLogSheet = ss.getSheetByName(SHEET_EDIT_LOG);
    editLogSheet.appendRow([
        new Date(),
        Utilities.getUuid(), // 每次修改都是一個新事件
        modifier,
        reason,
        '參加課程',
        before,
        after,
        name,
        '' // 課程簡稱欄位，此處合併修改，故留空
    ]);
}


// ===============================================================
// 檔案上傳功能 (File Upload Function)
// ===============================================================

/**
 * 處理檔案上傳
 * @param {object} formObject - 包含檔案資料和表單資訊的物件
 * @returns {object} - 操作結果
 */
function uploadFile(formObject) {
  try {
    const { fileData, fileName, department, uploader, reason } = formObject;
    
    const decodedData = Utilities.base64Decode(fileData.split(',')[1]);
    const blob = Utilities.newBlob(decodedData, MimeType.PDF, fileName); // 假設都是PDF

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileId = file.getId();
    const fileUrl = file.getUrl();

    // 記錄到 '檔案上傳紀錄'
    const logSheet = ss.getSheetByName(SHEET_FILE_LOG);
    const version = getNextVersion(department);
    logSheet.appendRow([
      department,
      fileName,
      fileId,
      fileUrl,
      version,
      uploader,
      reason,
      new Date()
    ]);

    return { success: true, message: '檔案已成功上傳！' };
  } catch (e) {
    console.error('uploadFile 失敗:', e);
    return { success: false, message: '檔案上傳失敗: ' + e.toString() };
  }
}

/**
 * 獲取指定部門的下一個版本號
 * @param {string} department - 部門名稱
 * @returns {number} - 新的版本號
 */
function getNextVersion(department) {
  const logSheet = ss.getSheetByName(SHEET_FILE_LOG);
  if (logSheet.getLastRow() < 2) return 1;
  const data = logSheet.getRange('A2:E' + logSheet.getLastRow()).getValues();
  const versions = data.filter(row => row[0] === department).map(row => row[4]);
  return versions.length > 0 ? Math.max(...versions) + 1 : 1;
}

// ===============================================================
// 總覽與報表功能 (Dashboard & Reporting Functions)
// ===============================================================

/**
 * 獲取各部門的填報狀態
 * @returns {object[]} - 各部門狀態物件的陣列
 */
function getDepartmentStatus() {
  try {
    const menuSheet = ss.getSheetByName(SHEET_MENU);
    const recordSheet = ss.getSheetByName(SHEET_RECORD);

    // 1. 統計各部門總人數
    const deptData = menuSheet.getRange('B2:C' + menuSheet.getLastRow()).getValues();
    const deptCounts = deptData.reduce((acc, row) => {
      const dept = row[0];
      if (dept) {
        if (!acc[dept]) acc[dept] = 0;
        acc[dept]++;
      }
      return acc;
    }, {});

    // 2. 統計已填報人數 (依據 "參加記錄" 中的唯一姓名)
    if (recordSheet.getLastRow() < 2) { // 如果參加記錄是空的
        return Object.keys(deptCounts).map(dept => ({
            department: dept,
            completed: 0,
            total: deptCounts[dept],
            status: '未鎖定' // 預設狀態
        }));
    }
    const recordData = recordSheet.getRange('C2:C' + recordSheet.getLastRow()).getValues();
    const completedCounts = recordData.flat().reduce((acc, name) => {
      const userDept = findDepartmentByName(deptData, name);
      if (userDept) {
        if (!acc[userDept]) acc[userDept] = new Set();
        acc[userDept].add(name);
      }
      return acc;
    }, {});

    // 3. 組合結果
    const status = Object.keys(deptCounts).map(dept => {
      const completed = completedCounts[dept] ? completedCounts[dept].size : 0;
      return {
        department: dept,
        completed: completed,
        total: deptCounts[dept],
        status: '未鎖定' // TODO: 增加鎖定狀態邏輯
      };
    });

    return status;
  } catch (e) {
    console.error('getDepartmentStatus 失敗:', e);
    return [];
  }
}

/**
 * 輔助函數：根據姓名查找部門
 */
function findDepartmentByName(data, name) {
  const userRow = data.find(row => row[1] === name);
  return userRow ? userRow[0] : null;
}

/**
 * 獲取指定部門用於列印的報名資料
 * @param {string} department - 部門名稱
 * @returns {object} - 整理好的報名資料
 */
function getPrintData(department) {
    try {
        const recordSheet = ss.getSheetByName(SHEET_RECORD);
        if (recordSheet.getLastRow() < 2) return {};
        const data = recordSheet.getRange('C2:D' + recordSheet.getLastRow()).getValues();
        
        const menuSheet = ss.getSheetByName(SHEET_MENU);
        const deptData = menuSheet.getRange('B2:C' + menuSheet.getLastRow()).getValues();

        const departmentMembers = deptData.filter(row => row[0] === department).map(row => row[1]);
        
        const printData = {};

        departmentMembers.forEach(name => {
            const courses = data.filter(row => row[0] === name).map(row => row[1]);
            if (courses.length > 0) {
                printData[name] = courses;
            } else {
                // 即使沒報名課程，也顯示出來，可能標示為 "無需參加" 或 "未填報"
                printData[name] = ["未報名任何課程"];
            }
        });

        return printData;
    } catch (e) {
        console.error('getPrintData 失敗:', e);
        return {};
    }
}
