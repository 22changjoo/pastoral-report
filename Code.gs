// ============================================================
// 목자 보고서 시스템 - Google Apps Script 백엔드 v2
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: "1_kfHyRKf094BJkIl-OTL98AV5vhTAYhkrQk0Fs89Teg",

  PASTOR_EMAILS: {
    "1초장": "21cdavidmy@igodswill.org",
    "2/3초장": "hwjcho@igodswill.org",
    "4초장": "ahavtika@igodswill.org",
    "5초장": "gwcj@igodswill.org",
    "6초장": "prophetcho@igodswill.org",
  },

  CC_EMAILS: [],
  CHURCH_NAME: "높은뜻푸른교회",
};

const HEADERS = [
  "제출일시",
  "모임날짜",
  "초장",
  "쉴물가(또는 목자이름)",
  "모임 장소",
  "참석인원",
  "불참자 및 사유",
  "모임내용",
  "오늘 모임은 어땠나요?",
  "기도제목",
];

// ── CORS 헤더 공통 응답 함수 ──
function makeJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET 요청: 보고서 조회 ──
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getReports') {
    return handleGetReports(e);
  }

  // 기본: 폼 HTML 반환 (기존 Apps Script URL용 — 이제 GitHub Pages로 이전)
  return HtmlService.createHtmlOutput('<p>GitHub Pages 버전을 사용해 주세요.</p>');
}

function handleGetReports(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetNames = ["1초장", "2/3초장", "4초장", "5초장", "6초장"];
    const reports = [];

    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
      data.forEach(row => {
        if (!row[0]) return; // 빈 행 스킵
        reports.push({
          submitDate: row[0] ? String(row[0]).substring(0, 19) : '',
          meetingDate: row[1] || '',
          choJang: row[2] || '',
          shilMulGa: row[3] || '',
          meetingPlace: row[4] || '',
          attendees: row[5] || '',
          absentees: row[6] || '',
          content: row[7] || '',
          evaluation: row[8] || '',
          prayerRequest: row[9] || '',
        });
      });
    });

    return makeJsonResponse({ success: true, reports });
  } catch (e) {
    return makeJsonResponse({ success: false, message: e.message });
  }
}

// ── POST 요청: 보고서 제출 ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = submitReport(data);
    return makeJsonResponse(result);
  } catch (err) {
    return makeJsonResponse({ success: false, message: err.message });
  }
}

// ── 보고서 저장 ──
function submitReport(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = formData.choJang;

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(HEADERS);
      formatHeaderRow(sheet);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      formatHeaderRow(sheet);
    }

    const now = new Date();
    const row = [
      Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
      formData.meetingDate,
      formData.choJang,
      formData.shilMulGa,
      formData.meetingPlace,
      formData.attendees,
      formData.absentees,
      formData.content,
      formData.evaluation,
      formData.prayerRequest,
    ];
    sheet.appendRow(row);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setVerticalAlignment("middle").setWrap(true);

    sendEmailNotification(formData, now);

    return { success: true, message: "보고서가 성공적으로 제출되었습니다." };
  } catch (e) {
    console.error("submitReport 오류:", e);
    return { success: false, message: "오류가 발생했습니다: " + e.message };
  }
}

// ── 이메일 알림 ──
function sendEmailNotification(formData, submitTime) {
  const toEmail = CONFIG.PASTOR_EMAILS[formData.choJang];
  if (!toEmail) return;

  const dateStr = Utilities.formatDate(submitTime, "Asia/Seoul", "yyyy년 MM월 dd일 HH:mm");
  const subject = `[목자보고서] ${formData.choJang} ${formData.shilMulGa} - ${formData.meetingDate}`;

  const htmlBody = `
  <!DOCTYPE html>
  <html>
  <head><meta charset="UTF-8"></head>
  <body style="margin:0;padding:0;background:#f9f9f9;">
    <div style="font-family:'Malgun Gothic',sans-serif;max-width:600px;margin:20px auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;background:white;">
      <div style="background:#2E7D32;color:white;padding:20px 24px;">
        <h2 style="margin:0;font-size:18px;">📋 목자 모임 보고서</h2>
        <p style="margin:4px 0 0;font-size:13px;opacity:.85;">${CONFIG.CHURCH_NAME}</p>
      </div>
      <div style="padding:24px;background:#f9f9f9;">
        <table style="width:100%;border-collapse:collapse;background:white;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.1);">
          <tr style="background:#E8F5E9;">
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;width:35%;border-bottom:1px solid #e0e0e0;">모임 날짜</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;">${formData.meetingDate}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">초장 / 쉴물가</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;">${formData.choJang} / ${formData.shilMulGa}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">모임 장소</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;">${formData.meetingPlace || '-'}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">참석자</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;">${formData.attendees || '-'}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">불참자 및 사유</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;white-space:pre-wrap;">${formData.absentees || '없음'}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">모임 내용</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;">${formData.content || '-'}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#1B5E20;background:#E8F5E9;border-bottom:1px solid #e0e0e0;">오늘 모임은?</td>
            <td style="padding:12px 16px;border-bottom:1px solid #e0e0e0;white-space:pre-wrap;">${formData.evaluation || '없음'}</td>
          </tr>
          <tr>
            <td style="padding:12px 16px;font-weight:bold;color:#C62828;background:#FFEBEE;border-bottom:1px solid #e0e0e0;">🙏 기도제목</td>
            <td style="padding:12px 16px;white-space:pre-wrap;color:${formData.prayerRequest ? '#C62828' : '#666'};font-weight:${formData.prayerRequest ? 'bold' : 'normal'};">${formData.prayerRequest || '없음'}</td>
          </tr>
        </table>
        <p style="margin:16px 0 0;font-size:12px;color:#888;text-align:right;">제출 시간: ${dateStr}</p>
      </div>
      <div style="padding:12px 24px;background:#f0f0f0;font-size:11px;color:#999;text-align:center;">
        이 메일은 목자 보고서 시스템에서 자동으로 발송되었습니다.
      </div>
    </div>
  </body>
  </html>`;

  const options = { htmlBody, name: `${CONFIG.CHURCH_NAME} 목자보고서 시스템` };
  if (CONFIG.CC_EMAILS && CONFIG.CC_EMAILS.length > 0) {
    options.cc = CONFIG.CC_EMAILS.join(',');
  }

  GmailApp.sendEmail(toEmail, subject, '', options);
}

// ── 헤더 스타일 ──
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
  const colWidths = [150, 100, 70, 100, 120, 70, 150, 200, 300, 200, 250];
  colWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

// ── 초기화 (최초 1회) ──
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheetNames = ["1초장", "2/3초장", "4초장", "5초장", "6초장"];
  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clearContents();
    sheet.appendRow(HEADERS);
    formatHeaderRow(sheet);
  });
  const defaultSheet = ss.getSheetByName("시트1") || ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  Logger.log("초기화 완료!");
}
