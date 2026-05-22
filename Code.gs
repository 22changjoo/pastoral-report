// ============================================================
// 목자 보고서 시스템 - Google Apps Script 백엔드
// ============================================================

// ▼▼▼ 여기에 설정값을 입력하세요 ▼▼▼
const CONFIG = {
  // 구글 스프레드시트 ID
  SPREADSHEET_ID: "1_kfHyRKf094BJkIl-OTL98AV5vhTAYhkrQk0Fs89Teg",

  // 담당 목사 이메일 (초장별)
  PASTOR_EMAILS: {
    "1초장": "21cdavidmy@igodswill.org",
    "2/3초장": "hwjcho@igodswill.org",
    "4초장": "ahavtika@igodswill.org",
    "5초장": "gwcj@igodswill.org",
    "6초장": "prophetcho@igodswill.org",
  },

  // 전체 담당 목사에게도 함께 보낼 경우 (없으면 빈 배열 [])
  CC_EMAILS: [],

  // 교회 이름
  CHURCH_NAME: "높은뜻푸른교회",
};
// ▲▲▲ 설정값 입력 끝 ▲▲▲

// 시트 헤더 정의
const HEADERS = [
  "제출일시",
  "모임날짜",
  "쉴물가",
  "장소",
  "참석인원",
  "불참인원 및 사유",
  "모임내용",
  "오늘 모임은 어땠나요",
  "기도제목",
];

/**
 * 웹앱 GET 요청 - admin.html에서 데이터 요청 시 JSON 반환, 그 외엔 HTML 반환
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getReports') {
    return getReportsJson();
  }
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("쉴물가 모임 보고서")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 웹앱 POST 요청 - 외부에서 JSON으로 전송된 폼 데이터를 받아 처리
 */
function doPost(e) {
  try {
    let formData;
    if (e && e.postData && e.postData.contents) {
      formData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      formData = e.parameter;
    } else {
      throw new Error("No data received");
    }

    const result = submitReport(formData);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 스프레드시트 셀 값을 날짜 문자열(yyyy-MM-dd)로 변환
 * Sheets가 날짜를 Date 객체로 저장한 경우에도 일관된 형식 반환
 */
function cellToDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd");
  }
  const s = String(val);
  // "yyyy-MM-dd HH:mm:ss" 또는 "yyyy-MM-dd" 형식에서 날짜 부분만 추출
  const m = s.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

/**
 * 보고서 목록을 JSON으로 반환 (admin.html 연동용)
 */
function getReportsJson() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let allReports = [];

    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return; // 헤더만 있는 경우 패스

      // 헤더: [0:제출일시, 1:모임날짜, 2:쉴물가, 3:장소, 4:참석인원, 5:불참인원 및 사유, 6:모임내용, 7:오늘 모임은 어땠나요, 8:기도제목]
      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        const attendeesStr = row[4] ? String(row[4]) : '';
        const count = attendeesStr ? attendeesStr.split(',').length : 0;
        const sheetChoJang = sheet.getName();

        allReports.push({
          submitDate: row[0] ? String(row[0]) : '',
          choJang: sheetChoJang,
          shilMulGa: row[2] ? String(row[2]) : '',
          meetingDate: cellToDateString(row[1]),
          meetingPlace: row[3] ? String(row[3]) : '',
          attendeeCount: count,
          attendees: attendeesStr,
          absentees: row[5] ? String(row[5]) : '',
          content: row[6] ? String(row[6]) : '',
          evaluation: row[7] ? String(row[7]) : '',
          prayerRequest: row[8] ? String(row[8]) : ''
        });
      }
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true, reports: allReports }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 보고서 제출 처리 (클라이언트에서 호출)
 */
function submitReport(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = formData.choJang; // 예: "1초장"

    // 해당 초장 시트 가져오기 (없으면 생성)
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(HEADERS);
      formatHeaderRow(sheet);
    }

    // 헤더가 없는 경우 (빈 시트) 추가
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      formatHeaderRow(sheet);
    }

    // 데이터 행 작성
    const now = new Date();
    const row = [
      Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
      formData.meetingDate,
      formData.shilMulGa,
      formData.meetingPlace,
      formData.attendees,
      formData.absentees,
      formData.content,
      formData.evaluation,
      formData.prayerRequest,
    ];
    sheet.appendRow(row);

    // 마지막 행 스타일 적용
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setVerticalAlignment("middle").setWrap(true);

    // 이메일 알림 발송
    sendEmailNotification(formData, now);

    return { success: true, message: "보고서가 성공적으로 제출되었습니다." };
  } catch (e) {
    console.error("submitReport 오류:", e);
    return { success: false, message: "오류가 발생했습니다: " + e.message };
  }
}

/**
 * 담당 목사에게 이메일 알림 발송
 */
function sendEmailNotification(formData, submitTime) {
  const toEmail = CONFIG.PASTOR_EMAILS[formData.choJang];
  if (!toEmail) {
    console.warn("이메일 주소가 설정되지 않은 초장:", formData.choJang);
    return;
  }

  const dateStr = Utilities.formatDate(submitTime, "Asia/Seoul", "yyyy년 MM월 dd일 HH:mm");
  const meetingDateStr = formData.meetingDate;

  const subject = `[목자보고서] ${formData.choJang} ${formData.shilMulGa} - ${meetingDateStr}`;

  const htmlBody = `
  <div style="font-family: 'Malgun Gothic', sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
    <div style="background: #2E7D32; color: white; padding: 20px 24px;">
      <h2 style="margin: 0; font-size: 18px;">목자 모임 보고서</h2>
      <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.85;">${CONFIG.CHURCH_NAME}</p>
    </div>
    <div style="padding: 24px; background: #f9f9f9;">
      <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <tr style="background: #E8F5E9;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; width: 35%; border-bottom: 1px solid #e0e0e0;">초장</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.choJang}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">쉴물가(또는 목자이름)</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.shilMulGa}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">모임 날짜</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${meetingDateStr}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">모임 장소</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.meetingPlace || "-"}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">참석자</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.attendees || "-"} (${formData.attendees ? formData.attendees.split(',').length : 0}명)</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">불참자 및 사유</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.absentees || "없음"}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">모임 내용</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.content || "-"}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">오늘 모임은 어땠나요?</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.evaluation || "없음"}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #C62828; background: #FFEBEE; border-bottom: 1px solid #e0e0e0;">기도제목</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap; color: ${formData.prayerRequest ? '#C62828' : '#666'}; font-weight: ${formData.prayerRequest ? 'bold' : 'normal'};">${formData.prayerRequest || "없음"}</td>
        </tr>
      </table>
      <p style="margin: 16px 0 0; font-size: 12px; color: #888; text-align: right;">제출 시간: ${dateStr}</p>
    </div>
    <div style="padding: 12px 24px; background: #f0f0f0; font-size: 11px; color: #999; text-align: center;">
      이 메일은 목자 보고서 시스템에서 자동으로 발송되었습니다.
    </div>
  </div>
  `;

  const options = {
    htmlBody: htmlBody,
    name: `${CONFIG.CHURCH_NAME} 목자보고서 시스템`,
  };

  if (CONFIG.CC_EMAILS && CONFIG.CC_EMAILS.length > 0) {
    options.cc = CONFIG.CC_EMAILS.join(",");
  }

  GmailApp.sendEmail(toEmail, subject, "", options);
}

/**
 * 헤더 행 스타일 적용
 */
function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setBackground("#2E7D32")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  // 열 너비 설정
  const colWidths = [150, 100, 150, 150, 200, 200, 200, 250, 250];
  colWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

/**
 * 스프레드시트 최초 초기화 (처음 한 번만 실행)
 * 스크립트 에디터에서 직접 실행하세요.
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheetNames = ["1초장", "2/3초장", "4초장", "5초장", "6초장"];

  sheetNames.forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    sheet.clearContents();
    sheet.appendRow(HEADERS);
    formatHeaderRow(sheet);
  });

  // 기본 시트 삭제 (있을 경우)
  const defaultSheet = ss.getSheetByName("시트1") || ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log("초기화 완료! 1~6초장 탭이 생성되었습니다.");
}
