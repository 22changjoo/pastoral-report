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
  "쉴물가",
  "장소",
  "참석인원",
  "불참인원 및 사유",
  "모임내용",
  "오늘 모임은 어땠나요?",
  "기도제목",
];

/**
 * 웹앱 GET 요청 - 폼 HTML 반환 또는 보고서 데이터 JSON 반환
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getReports') {
    return getReportsJson();
  }
  if (e && e.parameter && e.parameter.action === 'getShilMulGaMap') {
    return getShilMulGaMapJson();
  }
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("쉴물가 모임 보고서")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 스프레드시트 셀 값을 날짜 문자열(yyyy-MM-dd)로 변환
 */
function cellToDateString(val) {
  if (!val) return '';
  try {
    return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd");
  } catch(e) {}
  const s = String(val);
  const m = s.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

/**
 * 보고서 목록을 JSON으로 반환 (admin.html 대시보드용)
 * 컬럼: [0:제출일시, 1:모임날짜, 2:쉴물가, 3:장소, 4:참석인원,
 *        5:불참인원 및 사유, 6:모임내용, 7:오늘 모임은 어땠나요, 8:기도제목]
 */
function getReportsJson() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let allReports = [];

    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const attendeesStr = row[4] ? String(row[4]) : '';
        const count = attendeesStr ? attendeesStr.split(',').length : 0;

        allReports.push({
          submitDate: row[0] ? String(row[0]) : '',
          choJang: sheet.getName(),
          shilMulGa: row[2] ? String(row[2]) : '',
          meetingDate: cellToDateString(row[1]),
          meetingPlace: row[3] ? String(row[3]) : '',
          attendeeCount: count,
          attendees: attendeesStr,
          absentees: row[5] ? String(row[5]) : '',
          content: row[6] ? String(row[6]) : '',
          evaluation: row[7] ? String(row[7]) : '',
          prayerRequest: row[8] ? String(row[8]) : '',
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
 * '쉴물가이름' 탭에서 초장별 쉴물가 목록을 읽어 JSON으로 반환
 */
function getShilMulGaMapJson() {
  try {
    const mapData = getShilMulGaMap();
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: mapData }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 구글 스프레드시트의 '쉴물가이름' 탭에서 데이터를 추출하여 초장별 쉴물가 목록 맵 반환 (모든 양식 자동 파싱)
 */
function getShilMulGaMap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1. '쉴물가' 키워드가 들어간 탭 스마트 탐색
  let sheet = ss.getSheetByName("쉴물가이름");
  if (!sheet) sheet = ss.getSheetByName("쉴물가 이름");
  if (!sheet) sheet = ss.getSheetByName("쉴물가");
  if (!sheet) {
    const sheets = ss.getSheets();
    sheet = sheets.find(s => s.getName().includes("쉴물가") || s.getName().includes("이름"));
  }
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) return {};

  const result = {
    "1초장": [], "2초장": [], "3초장": [], "2/3초장": [], "4초장": [], "5초장": [], "6초장": []
  };

  // 초장 키 정규화 함수
  function normalizeChoJang(str) {
    const clean = String(str).replace(/\s+/g, '');
    const m = clean.match(/([1-6](\/[2-3])?초장)/);
    return m ? m[1] : null;
  }

  // A. 열(Column) 기반 가로형 탐색 (상단 1~3행 내에 초장명이 있는 경우)
  const colChoJangMap = {};
  for (let r = 0; r < Math.min(values.length, 3); r++) {
    for (let c = 0; c < values[r].length; c++) {
      const choKey = normalizeChoJang(values[r][c]);
      if (choKey) {
        colChoJangMap[c] = choKey;
      }
    }
  }

  const hasColMap = Object.keys(colChoJangMap).length > 0;
  if (hasColMap) {
    for (let c in colChoJangMap) {
      const choKey = colChoJangMap[c];
      for (let r = 0; r < values.length; r++) {
        const val = String(values[r][c]).trim();
        const isChoHeader = normalizeChoJang(val);
        // 초장 헤더 제목이 아닌 실제 쉴물가 이름만 추가
        if (val && !isChoHeader && val.length > 0) {
          if (!result[choKey]) result[choKey] = [];
          if (!result[choKey].includes(val)) {
            result[choKey].push(val);
          }
        }
      }
    }
  }

  // B. 행(Row) 기반 세로형 또는 자유 스캔 (A열에 초장, B열에 쉴물가 이름 등)
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cellVal = String(values[r][c]).trim();
      if (!cellVal) continue;

      // 셀 자체가 '5초장' 등인 경우 다음 옆 셀을 쉴물가 이름으로 간주
      const choKey = normalizeChoJang(cellVal);
      if (choKey && c + 1 < values[r].length) {
        const nextVal = String(values[r][c + 1]).trim();
        if (nextVal && !normalizeChoJang(nextVal)) {
          if (!result[choKey]) result[choKey] = [];
          if (!result[choKey].includes(nextVal)) {
            result[choKey].push(nextVal);
          }
        }
      }
      // "5초장 관악1쉴물가" 복합 형태인 경우
      if (choKey && cellVal.length > choKey.length) {
        const namePart = cellVal.replace(/^[1-6](\/[2-3])?초장\s*/, '').trim();
        if (namePart) {
          if (!result[choKey]) result[choKey] = [];
          if (!result[choKey].includes(namePart)) {
            result[choKey].push(namePart);
          }
        }
      }
    }
  }

  // 빈 항목 정리
  Object.keys(result).forEach(k => {
    if (result[k].length === 0) delete result[k];
  });

  return result;
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

    // 5초장: AI 분석 + Notion 심방 기록 저장
    if (formData.choJang === "5초장") {
      processChoJang5Report(formData);
    }

    return { success: true, message: "보고서가 성공적으로 제출되었습니다." };
  } catch (e) {
    Logger.log("submitReport 오류:", e);
    return { success: false, message: "오류가 발생했습니다: " + e.message };
  }
}

/**
 * 담당 목사에게 이메일 알림 발송
 */
function sendEmailNotification(formData, submitTime) {
  const toEmail = CONFIG.PASTOR_EMAILS[formData.choJang];
  if (!toEmail) {
    Logger.log("이메일 주소가 설정되지 않은 초장:", formData.choJang);
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
  const colWidths = [150, 100, 100, 120, 200, 200, 300, 200, 250];
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
// 5초장 자동화: Claude AI 특이사항 판단 + Notion 심방 기록
// Script Properties에 CLAUDE_API_KEY, NOTION_API_KEY 설정 필요
// ============================================================

const NOTION_DB_ID = "378b5584-3f1e-4c61-8333-db18ee1f1776";
const SLACK_CHANNEL_ID = "C0B6T5M3H0D";

/**
 * 5초장 보고서 제출 시 AI로 특이사항 판단 후 해당되면 Notion에 저장
 */
function processChoJang5Report(formData) {
  try {
    Logger.log("[1단계] Claude 분석 시작");
    const analysis = analyzeWithClaude(formData);
    Logger.log("[2단계] Claude 분석 완료 - hasSpecialCase: " + analysis.hasSpecialCase);
    Logger.log("[2단계] greetingMessage: " + (analysis.greetingMessage || "(없음)"));

    if (!analysis.hasSpecialCase) {
      Logger.log("특이사항 없음 - Notion 저장 생략 (" + (formData.shilMulGa || "") + ")");
      return;
    }

    Logger.log("[3단계] Notion 저장 시작");
    const notionPageId = saveToNotion({
      shilMulGa: formData.shilMulGa || "",
      targetName: analysis.targetName || formData.shilMulGa || "",
      meetingDate: formData.meetingDate || "",
      meetingPlace: formData.meetingPlace || "",
      summary: analysis.summary,
      urgentPrayer: analysis.urgentPrayer,
      followUp: analysis.followUp,
      greetingMessage: analysis.greetingMessage || "",
    });
    Logger.log("특이사항 감지 - Notion 저장 완료: " + notionPageId);

    Logger.log("[4단계] Slack 발송 시작");
    sendSlackGreeting({
      shilMulGa: formData.shilMulGa || "",
      meetingDate: formData.meetingDate || "",
      greetingMessage: analysis.greetingMessage || "",
    });
    Logger.log("[4단계] Slack 발송 완료");
  } catch (err) {
    Logger.log("5초장 자동화 오류: " + (err.message || String(err)));
  }
}

/**
 * Claude Sonnet으로 보고서 분석 및 특이사항 판단
 * @returns {{ hasSpecialCase: boolean, summary: string, urgentPrayer: string, followUp: string }}
 */
function analyzeWithClaude(formData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("CLAUDE_API_KEY가 Script Properties에 설정되지 않았습니다.");

  const prompt = `당신은 교회 목자 모임 보고서를 분석하는 전문가입니다. 담당 목사가 직접 돌봄이 필요한 특이사항이 있는지 판단해주세요.

[보고서]
쉴물가명: ${formData.shilMulGa || ""}
모임날짜: ${formData.meetingDate || ""}
참석인원: ${formData.attendees || ""}
불참인원 및 사유: ${formData.absentees || ""}
모임 내용: ${formData.content || ""}
오늘 모임 평가: ${formData.evaluation || ""}
기도제목: ${formData.prayerRequest || ""}

[특이사항 해당 기준 - 하나라도 해당되면 hasSpecialCase: true]
- 질병, 수술, 입원 언급
- 가족 사망, 이별, 심각한 가족 갈등
- 정신건강 문제 (공황장애, 우울증 등)
- 경제적 어려움 (실직, 파산, 채무 등)
- 불참이 반복되거나 불참 사유가 심각한 경우
- 기도제목에 긴급하거나 위기적인 내용

[일반 보고서 기준 - 아래만 해당되면 hasSpecialCase: false]
- 평범한 말씀 나눔, 교제, 식사
- "은혜로웠다", "감사하다" 수준의 내용
- 자녀 시험, 취업 등 일반적인 기도제목

반드시 아래 JSON 형식으로만 응답하세요 (설명 없이):
{"hasSpecialCase":true,"targetName":"...","summary":"...","urgentPrayer":"...","followUp":"...","greetingMessage":"..."}

hasSpecialCase가 false인 경우에도 나머지 필드는 빈 문자열로 채워 동일한 형식을 유지하세요.
targetName은 심방 및 돌봄이 필요한 구체적인 대상자(성도명)의 이름입니다. 만약 구체적인 대상자를 특정할 수 없다면 빈 문자열로 남겨두세요.
summary는 3~4문장, urgentPrayer는 감지된 위기 내용, followUp은 담당목사가 취해야 할 구체적 돌봄 방향(성경구절·안부메시지 제외).

greetingMessage 작성 규칙:
- 형식: 카카오톡/문자로 바로 보낼 수 있는 메시지 본문만 (📖💬 등 라벨 없이 순수 텍스트)
- 상황별 성경구절 선정: 질병/입원→사 41:10·시23편·롬8:28, 재정어려움→빌4:19·마6:31-33, 사별/상실→시34:18·사43:2, 정신건강→마11:28·사40:31, 가족갈등→골3:13-14·엡4:32, 진로불확실→잠3:5-6·렘29:11
- 목사가 직접 쓴 1인칭 구어체, 3-5줄, 성도명으로 시작, 구체적 문제 직접 언급 금지, 성경구절 자연스럽게 녹이기, 마지막은 기도하고 있다는 내용, 이모지 1-2개 절제
- 좋은 예: "집사님, 오늘 뵙고 많이 생각이 났어요.\n새로운 곳에서 적응하는 게 생각보다 쉽지 않으시죠.\n\"내가 너와 함께 하리라\" 하신 말씀 붙들고 하루하루 살아가시길 바랍니다.\n늘 기도하고 있어요 🙏"`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    payload: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1024,
      messages: [{ role: "user", content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error("Claude API 오류: " + result.error.message);

  const raw = result.content[0].text.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
  return JSON.parse(raw);
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
      "성도명": { title: [{ text: { content: data.targetName } }] },
      "심방일자": { date: { start: data.meetingDate } },
      "심방방법": { select: { name: "방문심방" } },
      "장소": { rich_text: [{ text: { content: data.meetingPlace } }] },
      "내용": { rich_text: [{ text: { content: data.summary } }] },
      "긴급기도제목": { rich_text: [{ text: { content: data.urgentPrayer } }] },
      "후속조치": { rich_text: [{ text: { content: data.followUp } }] },
      "안부메시지": { rich_text: [{ text: { content: data.greetingMessage } }] },
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

function testNotion5() {
  const testData = {
    choJang: "5초장",
    shilMulGa: "테스트 쉴물가",
    meetingDate: "2026-06-20",
    meetingPlace: "테스트 장소",
    attendees: "홍길동, 김철수",
    absentees: "",
    content: "말씀 나눔을 했습니다",
    evaluation: "은혜로웠습니다",
    prayerRequest: "홍길동이 많이 아파요. 수술도 해야해요.",
  };
  processChoJang5Report(testData);
}

/**
 * Notion DB 속성 타입 확인용 (처음 한 번만 실행)
 * Apps Script 에디터에서 직접 실행하여 컬럼 타입을 확인하세요.
 */
/**
 * Slack #심방안부메시지 채널에 안부메시지 발송
 */
function sendSlackGreeting(data) {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  if (!token) throw new Error("SLACK_BOT_TOKEN이 Script Properties에 설정되지 않았습니다.");

  const text = `📬 ${data.shilMulGa} | ${data.meetingDate}\n${data.greetingMessage}\n발송 후 ✅ 이모지 반응 달아주세요`;

  const response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: JSON.stringify({
      channel: SLACK_CHANNEL_ID,
      text: text,
    }),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error("Slack API 오류: " + result.error);
}

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
