/**
 * Google Apps Script for Golf Reservation System
 * This script handles data operations between the web app and Google Sheets.
 * It manages four sheets: Data (bookings), Quota (user limits), Log (activity tracking), and Password (user credentials).
 */

const SHEET_NAME_DATA = "Data";
const SHEET_NAME_QUOTA = "Quota";
const SHEET_NAME_LOG = "Log";
const SHEET_NAME_PASSWORD = "Password";

/**
 * HTTP GET 요청을 처리합니다.
 * 'read' 액션이 들어오면 모든 필요한 데이터(예약, 한도, 사용자 정보)를 함께 전송합니다.
 */
function doGet(e) {
  if (e.parameter.action === 'read') {
    const bookings = getBookings();
    const quotas = getQuotas();
    const users = getUsers();
    
    const responseData = { 
      bookings: bookings, 
      quotas: quotas,
      users: users 
    };
    
    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * HTTP POST 요청을 처리합니다.
 */
function doPost(e) {
  const request = JSON.parse(e.postData.contents);
  let response = {};
  try {
    switch (request.action) {
      case 'add':
        addBooking(request.booking, request.user);
        response = { status: 'success', message: '예약이 추가되었습니다.' };
        break;
      case 'update':
        updateBooking(request.booking, request.user, request.logMessage);
        response = { status: 'success', message: '예약이 수정되었습니다.' };
        break;
      case 'addMultiple':
        addMultipleBookings(request.bookings, request.user);
        response = { status: 'success', message: '데이터가 성공적으로 업로드되었습니다.' };
        break;
      case 'cancel':
        cancelBookings(request.ids, request.user);
        response = { status: 'success', message: '선택한 예약이 취소 처리되었습니다.' };
        break;
      default:
        throw new Error("알 수 없는 요청입니다.");
    }
  } catch (error) {
    response = { status: 'error', message: error.message };
  }
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName(sheetName);
}

function logAction(user, action, bookingDate, golfCourse, team) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = getSheet(SHEET_NAME_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOG);
    logSheet.appendRow(["Timestamp", "User", "Action", "Target Date", "Target Golf Course", "Target Team"]);
  }
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([timestamp, user, action, bookingDate || '', golfCourse || '', team || '']);
}

function getUsers() {
    const sheet = getSheet(SHEET_NAME_PASSWORD);
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];
    const headers = values[0];
    const idIndex = headers.indexOf('id');
    const passwordIndex = headers.indexOf('password');

    if (idIndex === -1 || passwordIndex === -1) return [];

    return values.slice(1).map(row => ({
        name: row[idIndex],
        phone: String(row[passwordIndex]) 
    }));
}

function getQuotas() {
    const quotaSheet = getSheet(SHEET_NAME_QUOTA);
    if (!quotaSheet) return [];
    const values = quotaSheet.getDataRange().getValues();
    if (values.length < 2) return [];
    const headers = values[0];
    return values.slice(1).map(row => {
        let obj = {};
        headers.forEach((header, i) => { obj[header] = row[i]; });
        return obj;
    });
}

function getBookings() {
  const sheet = getSheet(SHEET_NAME_DATA);
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const headers = values[0];
  return values.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      let cellValue = row[i];
      if ((header === 'Date' || header === 'Time') && cellValue) {
        try {
          const dateObj = new Date(cellValue);
          if (!isNaN(dateObj.getTime())) {
            if (header === 'Date') obj[header] = Utilities.formatDate(dateObj, spreadsheetTimeZone, "yyyy-MM-dd");
            else if (header === 'Time') obj[header] = Utilities.formatDate(dateObj, spreadsheetTimeZone, "HH:mm");
          } else { obj[header] = cellValue; }
        } catch (e) { obj[header] = cellValue; }
      } else { obj[header] = cellValue; }
    });
    return obj;
  });
}

function addBooking(booking, user) {
  const sheet = getSheet(SHEET_NAME_DATA);
  const newId = Utilities.getUuid();
  const headers = sheet.getRange(1, 1, 1, 8).getValues()[0]; // A-H 열 헤더만 사용
  const newRow = headers.map(header => {
    if (header === 'id') return newId;
    return booking[header] || "";
  });
  sheet.appendRow(newRow);
  logAction(user, "추가", booking['Date'], booking['Golf Course'], booking['Team']);
}

// [수정됨] 엑셀 업로드 시 A~H 열만 덮어쓰도록 로직 변경
function addMultipleBookings(bookings, user) {
  const sheet = getSheet(SHEET_NAME_DATA);
  // A열부터 H열까지의 헤더만 명시적으로 사용
  const headers = sheet.getRange(1, 1, 1, 8).getValues()[0];
  
  // 1. 기존 데이터 삭제 (A2부터 H열의 데이터가 있는 마지막 행까지)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).clearContent();
  }

  // 2. 새로운 데이터 준비
  const newRows = bookings.map(booking => {
    const newId = Utilities.getUuid();
    logAction(user || "Admin", "엑셀 업로드(추가)", booking.Date, booking['Golf Course'], booking.Team);
    return headers.map(header => {
      if (header === 'id') return newId;
      return booking[header] || "";
    });
  });

  // 3. 새로운 데이터를 A2 셀부터 H열까지만 입력
  if(newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
  }
}


function updateBooking(booking, user, logMessage) {
  const sheet = getSheet(SHEET_NAME_DATA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('id');
  if (idColumnIndex === -1) throw new Error("ID 열을 찾을 수 없습니다.");
  const rowIndex = data.findIndex(row => row[idColumnIndex] === booking.id);
  if (rowIndex === -1) throw new Error("수정할 예약을 찾을 수 없습니다.");
  const newRow = headers.map(header => booking[header] || "");
  sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
  logAction(user, logMessage || "수정", booking['Date'], booking['Golf Course'], booking['Team']);
}

function cancelBookings(ids, user) {
  const sheet = getSheet(SHEET_NAME_DATA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('id');
  const statusColumnIndex = headers.indexOf('Status');
  const dateColumnIndex = headers.indexOf('Date');
  const courseColumnIndex = headers.indexOf('Golf Course');
  const teamColumnIndex = headers.indexOf('Team');
  if (idColumnIndex === -1 || statusColumnIndex === -1) throw new Error("ID 또는 Status 열을 찾을 수 없습니다.");
  const spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  ids.forEach(id => {
      const rowIndex = data.findIndex(row => row[idColumnIndex] === id);
      if (rowIndex > -1) {
          const rowData = data[rowIndex];
          sheet.getRange(rowIndex + 1, statusColumnIndex + 1).setValue("취소중");
          const bookingDate = Utilities.formatDate(new Date(rowData[dateColumnIndex]), spreadsheetTimeZone, "yyyy-MM-dd");
          const golfCourse = rowData[courseColumnIndex];
          const team = rowData[teamColumnIndex];
          logAction(user, "취소", bookingDate, golfCourse, team);
      }
  });
}

