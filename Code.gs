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
    "2초장": "hwjcho@igodswill.org",
    "3초장": "hwjcho@igodswill.org",
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
  "초장",
  "쉴물가",
  "장소",
  "참석인원수",
  "참석인원",
  "불참인원 및 사유",
  "모임내용",
  "모임평가 및 건의사항",
  "기도제목",
];

/**
 * 웹앱 GET 요청 - 폼 HTML을 반환
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("쉴물가 모임 보고서")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 웹앱 POST 요청 - 외부 fetch()로 제출 처리
 */
function doPost(e) {
  try {
    const formData = JSON.parse(e.postData.contents);
    const result = submitReport(formData);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 보고서 제출 처리 (클라이언트에서 호출)
 */
function submitReport(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    // 시트명에 "/" 허용 안 되므로 "2/3초장" → "2-3초장"으로 변환
    const sheetName = formData.choJang.replace(/\//g, '-');

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
      formData.choJang,
      formData.shilMulGa,
      formData.meetingPlace,
      formData.attendeeCount,
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

    // 5초장: AI 분석 + Notion 심방 기록 저장
    if (formData.choJang === "5초장") {
      processChoJang5Report(formData);
    }

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
      <h2 style="margin: 0; font-size: 18px;">📋 목자 모임 보고서</h2>
      <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.85;">${CONFIG.CHURCH_NAME}</p>
    </div>
    <div style="padding: 24px; background: #f9f9f9;">
      <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <tr style="background: #E8F5E9;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; width: 35%; border-bottom: 1px solid #e0e0e0;">모임 날짜</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${meetingDateStr}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">초장 / 쉴물가</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.choJang} / ${formData.shilMulGa}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">장소</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.meetingPlace || "-"}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">참석인원</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0;">${formData.attendeeCount}명 | ${formData.attendees || "-"}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">불참인원 및 사유</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.absentees || "없음"}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">모임 내용</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.content || "-"}</td>
        </tr>
        <tr style="background: #fafafa;">
          <td style="padding: 12px 16px; font-weight: bold; color: #1B5E20; background: #E8F5E9; border-bottom: 1px solid #e0e0e0;">모임 평가 및 건의사항</td>
          <td style="padding: 12px 16px; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">${formData.evaluation || "없음"}</td>
        </tr>
        <tr>
          <td style="padding: 12px 16px; font-weight: bold; color: #C62828; background: #FFEBEE; border-bottom: 1px solid #e0e0e0;">🙏 기도제목</td>
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
  const colWidths = [150, 100, 70, 100, 120, 70, 150, 200, 300, 200, 250];
  colWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

/**
 * 스프레드시트 최초 초기화 (처음 한 번만 실행)
 * 스크립트 에디터에서 직접 실행하세요.
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheetNames = ["1초장", "2초장", "3초장", "4초장", "5초장", "6초장"];

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

// ============================================================
// 5초장 자동화: Claude AI 분석 + Notion 심방 기록
// Script Properties에 CLAUDE_API_KEY, NOTION_API_KEY 설정 필요
// ============================================================

const NOTION_DB_ID = "378b5584-3f1e-4c61-8333-db18ee1f1776";

/**
 * 5초장 보고서 제출 시 AI 분석 후 Notion에 저장
 */
function processChoJang5Report(formData) {
  try {
    const analysis = analyzeWithClaude(
      formData.content || "",
      formData.prayerRequest || "",
      formData.shilMulGa || "",
      formData.attendees || ""
    );
    const notionPageId = saveToNotion({
      shilMulGa: formData.shilMulGa || "",
      meetingDate: formData.meetingDate || "",
      meetingPlace: formData.meetingPlace || "",
      summary: analysis.summary,
      urgentPrayer: analysis.urgentPrayer,
      followUp: analysis.followUp,
    });
    console.log("Notion 심방 기록 저장 완료 - 페이지 ID:", notionPageId);
  } catch (err) {
    // 자동화 실패가 보고서 제출 자체를 막지 않도록 에러만 로깅
    console.error("5초장 자동화 오류:", err.message);
  }
}

/**
 * Claude API로 모임 내용 분석
 */
function analyzeWithClaude(content, prayerRequest, shilMulGa, attendees) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("CLAUDE_API_KEY가 Script Properties에 설정되지 않았습니다.");

  const prompt = `당신은 교회 목자 모임 보고서를 분석하는 전문가입니다.

쉴물가명: ${shilMulGa}
참석인원: ${attendees}
모임 내용: ${content}
기도제목: ${prayerRequest}

다음 세 가지를 JSON으로 답해주세요:
1. summary: 모임 내용 3~4문장 요약 (핵심 나눔, 분위기, 결정사항 포함)
2. urgentPrayer: 기도제목에서 질병·경제 위기·가정 위기 등 긴급한 상황만 추출. 없으면 빈 문자열.
3. followUp: 목사의 돌봄이 필요한 성도나 상황에 대한 구체적 후속조치 제안. 없으면 빈 문자열.

반드시 아래 JSON 형식으로만 응답하세요 (설명 없이):
{"summary":"...","urgentPrayer":"...","followUp":"..."}`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 1024,
      messages: [{ role: "user", content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error("Claude API 오류: " + result.error.message);

  return JSON.parse(result.content[0].text);
}

/**
 * Notion DB에 심방 기록 페이지 생성
 */
function saveToNotion(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("NOTION_API_KEY");
  if (!apiKey) throw new Error("NOTION_API_KEY가 Script Properties에 설정되지 않았습니다.");

  const payload = {
    parent: { database_id: NOTION_DB_ID },
    properties: {
      "성도명": { title: [{ text: { content: data.shilMulGa } }] },
      "심방일자": { date: { start: data.meetingDate } },
      "심방방법": { select: { name: "방문심방" } },
      "장소": { rich_text: [{ text: { content: data.meetingPlace } }] },
      "내용": { rich_text: [{ text: { content: data.summary } }] },
      "긴급기도제목": { rich_text: [{ text: { content: data.urgentPrayer } }] },
      "후속조치": { rich_text: [{ text: { content: data.followUp } }] },
    },
  };

  const response = UrlFetchApp.fetch("https://api.notion.com/v1/pages", {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Notion-Version": "2022-06-28",
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (result.object === "error") throw new Error("Notion API 오류: " + result.message);

  return result.id;
}

/**
 * Notion DB 속성 타입 확인용 (처음 한 번만 실행)
 * Apps Script 에디터에서 직접 실행하여 컬럼 타입을 확인하세요.
 */
function checkNotionDbSchema() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("NOTION_API_KEY");
  const response = UrlFetchApp.fetch("https://api.notion.com/v1/databases/" + NOTION_DB_ID, {
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Notion-Version": "2022-06-28",
    },
    muteHttpExceptions: true,
  });
  const result = JSON.parse(response.getContentText());
  const props = result.properties;
  Object.keys(props).forEach(key => {
    Logger.log(key + " → " + props[key].type);
  });
}
