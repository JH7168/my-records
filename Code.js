function doGet() {
  // 💡 [수정됨] 분리된 HTML 파일들을 정상적으로 조립할 수 있도록 템플릿 방식으로 변경
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('정환❤️선영 다이어리')
      // 💡 [수정됨] 페이지 자체의 핀치 확대와 지도 내부 확대(echarts roam)가 동시에 반응해서
      // 확대/축소가 널뛰던 문제를 막기 위해 페이지 레벨 핀치 줌은 막아둔다.
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시트1');
}

// 💡 한 시트에 두 개의 표(A~F열 데이트기록 / J~N열 통장기록)가 같이 있어서
// getLastRow()만으로는 각 표의 실제 마지막 행을 알 수 없다. 지정한 열 범위 안에서
// 값이 있는 마지막 행을 뒤에서부터 찾아준다. (데이터가 없으면 헤더 행을 반환)
function findLastDataRow_(sheet, startCol, numCols) {
  const lastRow = sheet.getLastRow();
  let last = 1;
  if (lastRow >= 1) {
    const data = sheet.getRange(1, startCol, lastRow, numCols).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i].some(cell => cell !== "")) {
        last = i + 1;
        break;
      }
    }
  }
  return last;
}

// === [데이트 기록] 시트 헤더(1행) 초기화 ===
function ensureDateHeaders_(sheet) {
  if (sheet.getRange("A1").getValue() !== "") return;
  const headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setValues([['날짜', '지역1', '지역2', '장소', '분류', '내용']]);
  headerRange.setBackground("#ffe3e7").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  const catHeader = sheet.getRange("H1");
  catHeader.setValue("분류 종류");
  catHeader.setBackground("#ffe3e7").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
}

// === [통장기록] 시트 헤더(1행) 및 시작 금액 초기화 ===
function ensureAccountHeaders_(sheet) {
  if (sheet.getRange("J1").getValue() !== "") return;
  const headerRange = sheet.getRange("J1:N1");
  headerRange.setValues([['기존금액', '날짜', '구분', '분류', '내용']]);
  headerRange.setBackground("#ffe3e7").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("J2:N2").setValues([[588611, '2026-06-25', '입금', '💰 정기 모임비', '시작 금액']]);
  sheet.getRange("J1:N2").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
}

// === [데이트 기록] A~F열을 읽어 레코드 배열로 변환 ===
function readDateRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const records = [];
  data.forEach((row, index) => {
    if (row[0] !== "") {
      records.push({
        rowIdx: index + 2, // 💡 수정을 위해 줄 번호 기억
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[0],
        province: row[1],
        region: row[2],
        place: row[3],
        category: row[4],
        content: row[5]
      });
    }
  });
  return records;
}

// === [통장기록] J~N열을 읽어 레코드 배열로 변환 ===
function readAccountRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dataJN = sheet.getRange(2, 10, lastRow - 1, 5).getValues();
  const records = [];

  dataJN.forEach((row, index) => {
    if (row[0] !== "" || row[1] !== "") {
      let dateStr = row[1];
      if (row[1] instanceof Date) {
        dateStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      records.push({
        rowIdx: index + 2, // 💡 수정을 위해 줄 번호 기억
        amount: row[0],
        date: dateStr || '',
        type: row[2] || '출금',
        category: row[3] || '',
        content: row[4] || ''
      });
    }
  });
  return records;
}

// === [데이트 기록] 불러오기 (A~F열) ===
function getDateData() {
  const sheet = getSheet_();
  if (!sheet) return [];
  return readDateRecords_(sheet);
}

// [데이트 기록] 카테고리 종류 가져오기
function getCategories() {
  const sheet = getSheet_();
  if (!sheet) return ['음식', '활동', '기타'];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return ['음식', '활동', '기타'];
  const data = sheet.getRange(1, 8, lastRow, 1).getValues();
  const categories = data.flat().filter(item => item !== "" && item !== "분류 종류");
  return categories.length > 0 ? [...new Set(categories)] : ['음식', '활동', '기타'];
}

// === [통장기록] 불러오기 (J~N열) ===
function getAccountData() {
  const sheet = getSheet_();
  if (!sheet) return [];
  ensureAccountHeaders_(sheet);
  return readAccountRecords_(sheet);
}

// === [데이트 기록] 저장하기 === (성공 시 최신 전체 레코드를 바로 반환해 재조회 왕복을 없앤다)
function saveDateRecord(record) {
  const sheet = getSheet_();
  if (!sheet) return [];

  ensureDateHeaders_(sheet);

  const nextRow = findLastDataRow_(sheet, 1, 6) + 1;
  sheet.getRange(nextRow, 1, 1, 6).setValues([[record.date, record.province, record.region, record.place, record.category, record.content]]);

  // [수정됨] 데이터 영역 전체 가운데 정렬 및 테두리 적용으로 깔끔하게 통일
  const dataRange = sheet.getRange(1, 1, nextRow, 6);
  dataRange.setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

  if (nextRow >= 3) {
    sheet.getRange(2, 1, nextRow - 1, 6).sort({column: 1, ascending: true});
  }
  return readDateRecords_(sheet);
}

// === [통장기록] 저장하기 === (성공 시 최신 전체 레코드를 바로 반환해 재조회 왕복을 없앤다)
function saveAccountRecord(record) {
  const sheet = getSheet_();
  if (!sheet) return [];

  ensureAccountHeaders_(sheet);

  const nextRow = findLastDataRow_(sheet, 10, 5) + 1;
  sheet.getRange(nextRow, 10, 1, 5).setValues([[record.amount, record.date, record.type, record.category, record.content]]);

  // [수정됨] 데이터 영역 전체 가운데 정렬 및 테두리 적용으로 깔끔하게 통일
  const accountDataRange = sheet.getRange(1, 10, nextRow, 5);
  accountDataRange.setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);

  if (nextRow >= 4) {
    sheet.getRange(3, 10, nextRow - 2, 5).sort({column: 11, ascending: true});
  }
  return readAccountRecords_(sheet);
}

// === [데이트 기록] 수정하기 ===
function updateDateRecord(rowIdx, record) {
  const sheet = getSheet_();
  sheet.getRange(rowIdx, 1, 1, 6).setValues([[record.date, record.province, record.region, record.place, record.category, record.content]]);
  const lastRow = findLastDataRow_(sheet, 1, 6);
  if (lastRow >= 3) sheet.getRange(2, 1, lastRow - 1, 6).sort({column: 1, ascending: true});
  return readDateRecords_(sheet);
}

// === [통장기록] 수정하기 ===
function updateAccountRecord(rowIdx, record) {
  const sheet = getSheet_();
  sheet.getRange(rowIdx, 10, 1, 5).setValues([[record.amount, record.date, record.type, record.category, record.content]]);
  const lastRow = findLastDataRow_(sheet, 10, 5);
  if (lastRow >= 4) sheet.getRange(3, 10, lastRow - 2, 5).sort({column: 11, ascending: true});
  return readAccountRecords_(sheet);
}

// === [데이트 기록] 삭제하기 ===
// ⚠️ 같은 시트에 통장기록(J~N열)도 함께 있으므로 deleteRow()로 행 전체를 지우면 안 되고,
// A~F열 범위만 한 칸씩 당겨서 지운다.
function deleteDateRecord(rowIdx) {
  const sheet = getSheet_();
  if (!sheet) return [];

  const last = findLastDataRow_(sheet, 1, 6);
  if (rowIdx >= 2 && rowIdx <= last) {
    if (rowIdx < last) {
      const shiftValues = sheet.getRange(rowIdx + 1, 1, last - rowIdx, 6).getValues();
      sheet.getRange(rowIdx, 1, last - rowIdx, 6).setValues(shiftValues);
    }
    sheet.getRange(last, 1, 1, 6).clearContent();
  }
  return readDateRecords_(sheet);
}

// === [통장기록] 삭제하기 ===
// ⚠️ 같은 시트에 데이트기록(A~F열)도 함께 있으므로 J~N열 범위만 당겨서 지운다.
function deleteAccountRecord(rowIdx) {
  const sheet = getSheet_();
  if (!sheet) return [];

  const last = findLastDataRow_(sheet, 10, 5);
  if (rowIdx >= 2 && rowIdx <= last) {
    if (rowIdx < last) {
      const shiftValues = sheet.getRange(rowIdx + 1, 10, last - rowIdx, 5).getValues();
      sheet.getRange(rowIdx, 10, last - rowIdx, 5).setValues(shiftValues);
    }
    sheet.getRange(last, 10, 1, 5).clearContent();
  }
  return readAccountRecords_(sheet);
}

// === [데이트 추천] 제미나이 API 연동 ===
function callGeminiAPI(prompt) {
  // 💡 API 키는 코드에 직접 저장하지 않고 PropertiesService(스크립트 속성)에서 불러옵니다.
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return "앗! API 키가 설정되어 있지 않습니다. Apps Script 편집기의 [프로젝트 설정 > 스크립트 속성]에서 GEMINI_API_KEY를 등록해 주세요.";
  }
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  const payload = {
    "contents": [{
      "parts": [{ "text": prompt }]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // 1. 정상적으로 추천 답변이 도착한 경우
    if (json.candidates && json.candidates.length > 0) {
      return json.candidates[0].content.parts[0].text;
    }
    // 2. 서버 과부하 (503 에러) 인 경우
    else if (json.error && json.error.code === 503) {
      return "앗! 현재 제미나이 접속자가 너무 많아 서버가 잠시 숨을 고르고 있어요. 😥\n\n1~2분 정도만 기다리셨다가 다시 버튼을 눌러주시면 멋진 코스를 짜드릴게요!";
    }
    // 3. 단기간 요청 제한 (429 에러 - 혹시 모를 대비용)
    else if (json.error && json.error.code === 429) {
      return "앗! 짧은 시간에 너무 많은 요청이 들어왔어요. ⏳\n\n1분 정도만 기다리셨다가 다시 시도해 주세요!";
    }
    // 4. API 키 오류 등 기타 문제인 경우
    else if (json.error) {
      return "앗! 답변을 생성하는 데 문제가 발생했습니다. API 키가 정확하게 입력되었는지 다시 한번 확인해 주세요! 😥\n\n(상세 에러: " + json.error.message + ")";
    }
    // 5. 알 수 없는 구조의 에러
    else {
      return "응답을 처리할 수 없습니다. 다시 시도해 주세요.";
    }
  } catch (e) {
    return "제미나이 서버와 통신하는 중 오류가 발생했습니다: " + e.toString();
  }
}

// === 권한 강제 획득용 치트키 함수 ===
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}
