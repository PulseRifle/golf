// ##### 상수 설정 (CONSTANTS) #####
const DATA_SHEET_NAME = 'Data';
const USERS_SHEET_NAME = 'Users';
const QUOTA_SHEET_NAME = 'Quota';
const LOG_SHEET_NAME = 'Log';
const ENTERTAINMENT_SHEET_NAME = '접대비'; // 새로 추가

// HTTP GET 요청을 처리합니다. 주로 데이터 조회를 담당합니다.
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'read') {
      const data = getAllData();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error("알 수 없는 GET 요청입니다.");
    }
  } catch (error) {
    Logger.log(error.stack);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// HTTP POST 요청을 처리합니다. 데이터 추가, 수정, 삭제 등의 작업을 담당합니다.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30초간 락 대기

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result = {};

    switch (action) {
      case 'add':
        result = addBooking(data.user, data.booking);
        break;
      case 'update':
        result = updateBooking(data.user, data.booking, data.logMessage);
        break;
      case 'cancel':
        result = cancelBookings(data.user, data.ids);
        break;
      case 'addMultiple':
        result = addMultipleBookings(data.user, data.bookings);
        break;
      case 'saveEntertainment': // 새로 추가
        result = saveEntertainment(data.name, data.commonEntertainment, data.personalEntertainment);
        break;
      default:
        throw new Error("알 수 없는 POST 요청입니다.");
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(error.stack);
    logAction('SYSTEM', 'ERROR', error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}


/**
 * 모든 시트에서 초기 데이터를 가져옵니다.
 * @returns {object} bookings, users, quotas 데이터를 포함하는 객체
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  const quotaSheet = ss.getSheetByName(QUOTA_SHEET_NAME);

  // [수정됨] 데이터 시트에서 'EE', 'WE' 컬럼을 포함하여 11개 열을 읽어옵니다.
  // 컬럼 순서: id, Golf Course, Name, Date, Team, Time, Status, Comment, MP, EE, WE
  const bookingsData = sheetDataToObjects(dataSheet.getDataRange().getValues(), 11);
  const usersData = sheetDataToObjects(usersSheet.getDataRange().getValues());
  const quotasData = sheetDataToObjects(quotaSheet.getDataRange().getValues());

  return {
    bookings: bookingsData,
    users: usersData,
    quotas: quotasData
  };
}


/**
 * 신규 예약을 'Data' 시트에 추가합니다.
 * @param {string} user - 작업을 수행한 사용자 이름
 * @param {object} booking - 추가할 예약 정보
 * @returns {object} 성공 상태 메시지
 */
function addBooking(user, booking) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const newId = "booking_" + new Date().getTime(); // 고유 ID 생성

  // [수정됨] 신규 예약 시 MP 기본값으로 1, EE 0, WE 빈값을 설정합니다.
  dataSheet.appendRow([
    newId,
    booking['Golf Course'],
    booking['Name'],
    booking['Date'],
    booking['Team'],
    booking['Time'],
    booking['Status'],
    booking['Comment'],
    1, // MP 기본값
    booking['EE'] || 0, // EE 기본값
    booking['WE'] || '' // WE 컬럼 (빈값 또는 "주 말")
  ]);

  logAction(user, 'ADD', `신규 예약 추가: ${booking['Date']} ${booking['Golf Course']} (${booking['Name']})`);
  return { status: 'success', message: '예약이 성공적으로 추가되었습니다.' };
}


/**
 * 기존 예약을 수정합니다.
 * @param {string} user - 작업을 수행한 사용자 이름
 * @param {object} booking - 수정할 예약 정보
 * @param {string} logMessage - 로그에 남길 메시지
 * @returns {object} 성공 상태 메시지
 */
function updateBooking(user, booking, logMessage) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('id');

  const rowIndex = data.findIndex(row => row[idIndex] == booking.id);

  if (rowIndex > 0) {
    // [수정됨] 'EE', 'WE' 컬럼까지 11개 열을 업데이트합니다. 기존 값은 유지됩니다.
    const rowData = data[rowIndex];
    dataSheet.getRange(rowIndex + 1, 1, 1, 11).setValues([[
      booking.id,
      booking['Golf Course'],
      booking['Name'],
      booking['Date'],
      booking['Team'],
      booking['Time'],
      booking['Status'],
      booking['Comment'],
      rowData[8], // 기존 MP 값 유지
      rowData[9], // 기존 EE 값 유지
      rowData[10] // 기존 WE 값 유지
    ]]);

    logAction(user, 'UPDATE', `${logMessage}: ${booking.id}`);
    return { status: 'success', message: '예약이 성공적으로 수정되었습니다.' };
  } else {
    throw new Error("수정할 예약을 찾을 수 없습니다.");
  }
}

/**
 * 여러 예약을 '취소중' 상태로 변경합니다.
 * @param {string} user - 작업을 수행한 사용자 이름
 * @param {Array<string>} ids - 취소할 예약 ID 목록
 * @returns {object} 성공 상태 메시지
 */
function cancelBookings(user, ids) {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('id');
    const statusIndex = headers.indexOf('Status');
    let cancelledCount = 0;

    data.forEach((row, index) => {
        if (index > 0 && ids.includes(row[idIndex])) {
            dataSheet.getRange(index + 1, statusIndex + 1).setValue('취소중');
            cancelledCount++;
        }
    });

    if (cancelledCount > 0) {
        logAction(user, 'CANCEL', `${cancelledCount}개 항목 취소 요청: ${ids.join(', ')}`);
        return { status: 'success', message: `${cancelledCount}개의 예약이 '취소중' 상태로 변경되었습니다.` };
    } else {
        throw new Error("취소할 예약을 찾을 수 없습니다.");
    }
}


/**
 * CSV 파일로부터 여러 예약을 한 번에 추가합니다.
 * @param {string} user - 작업을 수행한 사용자 이름
 * @param {Array<object>} bookings - 추가할 예약 객체 배열
 * @returns {object} 성공 상태 메시지
 */
function addMultipleBookings(user, bookings) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);

  bookings.forEach(booking => {
    const newId = "booking_" + new Date().getTime() + Math.random();
    // [수정됨] MP, EE, WE 값이 없으면 기본값으로 설정합니다.
    dataSheet.appendRow([
      newId,
      booking['Golf Course'],
      booking['Name'],
      booking['Date'],
      booking['Team'],
      booking['Time'],
      booking['Status'],
      booking['Comment'],
      booking['MP'] || 1,
      booking['EE'] || 0,
      booking['WE'] || ''
    ]);
  });

  logAction(user, 'ADD_MULTIPLE', `${bookings.length}개의 예약을 엑셀로 추가했습니다.`);
  return { status: 'success', message: `${bookings.length}개의 예약이 성공적으로 추가되었습니다.` };
}


/**
 * [새로 추가] 접대비 정보를 '접대비' 시트에 저장합니다.
 * @param {string} name - 사용자 이름
 * @param {string} commonEntertainment - 공통접대비
 * @param {string} personalEntertainment - 개인접대비
 * @returns {object} 성공 상태 메시지
 */
function saveEntertainment(name, commonEntertainment, personalEntertainment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let entertainmentSheet = ss.getSheetByName(ENTERTAINMENT_SHEET_NAME);

  // '접대비' 시트가 없으면 생성
  if (!entertainmentSheet) {
    entertainmentSheet = ss.insertSheet(ENTERTAINMENT_SHEET_NAME);
    entertainmentSheet.appendRow(['Name', '공통접대비', '개인접대비']);
  }

  const data = entertainmentSheet.getDataRange().getValues();
  const headers = data[0];
  const nameColIndex = headers.indexOf('Name');
  const commonColIndex = headers.indexOf('공통접대비');
  const personalColIndex = headers.indexOf('개인접대비');

  // Name에 해당하는 행 찾기
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameColIndex] === name) {
      rowIndex = i + 1; // Sheet는 1-based index
      break;
    }
  }

  // 행을 찾았으면 업데이트, 없으면 새로 추가
  if (rowIndex > 0) {
    entertainmentSheet.getRange(rowIndex, commonColIndex + 1).setValue(commonEntertainment);
    entertainmentSheet.getRange(rowIndex, personalColIndex + 1).setValue(personalEntertainment);
  } else {
    // 새 행 추가
    entertainmentSheet.appendRow([name, commonEntertainment, personalEntertainment]);
  }

  logAction(name, 'SAVE_ENTERTAINMENT', `공통접대비: ${commonEntertainment}, 개인접대비: ${personalEntertainment}`);
  return { status: 'success', message: '접대비 정보가 저장되었습니다.' };
}


/**
 * 시트 데이터를 객체 배열로 변환하는 헬퍼 함수입니다.
 * @param {Array<Array<string>>} data - 시트에서 읽어온 2D 배열 데이터
 * @param {number} [numColumns=null] - 읽어올 열의 개수 (지정하지 않으면 전체)
 * @returns {Array<object>} 변환된 객체 배열
 */
function sheetDataToObjects(data, numColumns = null) {
  const headers = data[0];
  const effectiveNumColumns = numColumns ? numColumns : headers.length;

  return data.slice(1).map(row => {
    let obj = {};
    for (let i = 0; i < effectiveNumColumns; i++) {
      if (headers[i]) {
        obj[headers[i]] = row[i];
      }
    }
    return obj;
  });
}

/**
 * Log 시트에 활동 기록을 남깁니다.
 * @param {string} user - 작업을 수행한 사용자
 * @param {string} action - 수행한 작업 종류
 * @param {string} details - 작업 상세 내용
 */
function logAction(user, action, details) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  logSheet.appendRow([new Date(), user, action, details]);
}
