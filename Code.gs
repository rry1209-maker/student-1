/**
 * ============================================================
 * 학생 유형별 데이터 분석 시뮬레이터 - 백엔드 (Google Apps Script)
 * 구글 시트에 교사의 진단/처방 실습 결과를 저장합니다.
 * ============================================================
 *
 * [배포 방법]
 * 1. https://sheets.google.com 에서 새 구글 시트를 만듭니다.
 * 2. 상단 메뉴 [확장 프로그램] > [Apps Script] 를 클릭합니다.
 * 3. 기본 Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣습니다.
 * 4. 상단 [배포] > [새 배포] > 유형 선택(톱니바퀴) > [웹 앱] 을 선택합니다.
 *    - 설명: 아무거나
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한이 있는 사용자: "모든 사용자"
 * 5. [배포]를 누르고 권한을 승인하면 "웹 앱 URL"이 나옵니다.
 * 6. 그 URL을 프론트엔드(index.html)의 GAS_URL 변수에 붙여넣으면 끝!
 * ============================================================
 */

// 저장할 시트 이름
const SHEET_NAME = "실습기록";

// 시트 헤더(열 제목)
const HEADERS = [
  "기록시각",
  "교사명",
  "학생ID",
  "학생명",
  "실제유형",
  "교사진단",
  "진단정답여부",
  "교사처방",
  "처방적절성",
  "소요시간(초)",
  "비고",
];

/**
 * 웹앱 GET 요청 처리 - 동작 확인 및 통계 조회용
 */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;

  if (action === "stats") {
    return jsonResponse(getStats());
  }

  return jsonResponse({
    status: "ok",
    message: "학생 유형별 데이터 분석 시뮬레이터 백엔드가 정상 동작 중입니다.",
    time: new Date().toISOString(),
  });
}

/**
 * 웹앱 POST 요청 처리 - 실습 기록 저장
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "전송된 데이터가 없습니다." });
    }
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    const row = [
      new Date(),
      data.teacherName || "익명",
      data.studentId || "",
      data.studentName || "",
      data.actualType || "",
      data.teacherDiagnosis || "",
      data.diagnosisCorrect ? "정답" : "오답",
      data.teacherPrescription || "",
      data.prescriptionQuality || "",
      data.elapsedSeconds || "",
      data.note || "",
    ];

    sheet.appendRow(row);

    return jsonResponse({
      status: "success",
      message: "실습 기록이 저장되었습니다.",
      savedRow: row,
    });
  } catch (err) {
    return jsonResponse({
      status: "error",
      message: "저장 중 오류: " + err.message,
    });
  }
}

/**
 * 시트를 가져오거나, 없으면 생성하고 헤더를 추가
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // 헤더 스타일링
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4F46E5");
    headerRange.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 간단한 통계 반환 (총 실습 수, 진단 정답률 등)
 */
function getStats() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { totalRecords: 0, diagnosisAccuracy: 0 };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const total = data.length;
  let correct = 0;
  data.forEach(function (r) {
    if (r[6] === "정답") correct++;
  });

  return {
    totalRecords: total,
    diagnosisAccuracy: Math.round((correct / total) * 100),
  };
}

/**
 * JSON 응답 헬퍼
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
