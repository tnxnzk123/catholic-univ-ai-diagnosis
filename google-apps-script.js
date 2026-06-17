/**
 * ============================================================
 * 가톨릭대학교 AI 역량교육 사전 수준 진단 — Google Apps Script
 * ============================================================
 * 
 * 이 코드를 구글 시트의 Apps Script 편집기에 붙여넣으세요.
 * 자세한 설정 방법은 '연동_설정_가이드.md'를 참고하세요.
 * 
 * 기능:
 * - 웹앱에서 POST로 전송된 진단 응답 데이터를 구글 시트에 한 행으로 저장
 * - 컬럼 순서: 제출시각, 이름, 소속부서, Q1~Q6, 권장트랙, 산출사유
 */


/**
 * doPost — 웹앱으로 배포 시 POST 요청을 처리하는 함수
 * 프론트엔드에서 fetch POST로 JSON 데이터를 보내면 이 함수가 실행됩니다.
 */
function doPost(e) {
  try {
    // 요청 본문에서 JSON 파싱
    const data = JSON.parse(e.postData.contents);

    // 현재 스프레드시트의 첫 번째 시트 가져오기
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 시트에 한 행 추가 (컬럼 순서: 개발문서 3.3절 스키마)
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('ko-KR'),  // 제출시각
      data.name      || '',                                   // 이름
      data.dept      || '',                                   // 소속부서
      data.q1        || '',                                   // Q1 선택 번호
      data.q2        || '',                                   // Q2 선택 번호
      data.q3        || '',                                   // Q3 선택 번호
      data.q4        || '',                                   // Q4 선택 번호
      data.q5        || '',                                   // Q5 선택 번호
      data.q6        || '',                                   // Q6 선택 번호 (초급 미해당 시 공란)
      data.track     || '',                                   // 권장트랙
      data.reason    || ''                                    // 산출사유
    ]);

    // 성공 응답
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: '저장 완료' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // 에러 응답
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * doGet — GET 요청 처리 (연동 테스트용)
 * 브라우저에서 웹앱 URL을 직접 열면 이 메시지가 표시됩니다.
 */
function doGet(e) {
  return ContentService
    .createTextOutput('가톨릭대학교 AI 역량교육 사전 수준 진단 — 데이터 수집 서버가 정상 작동 중입니다.')
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * setupHeaders — 시트 헤더 자동 설정 (최초 1회 실행)
 * Apps Script 편집기에서 이 함수를 선택하고 ▶ 실행 버튼을 누르면,
 * 시트 첫 번째 행에 컬럼 헤더가 자동으로 입력됩니다.
 */
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 헤더 데이터 (개발문서 3.3절 스키마 순서)
  const headers = [
    '제출시각', '이름', '소속부서',
    'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6',
    '권장트랙', '산출사유'
  ];

  // 첫 번째 행에 헤더 입력
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // 헤더 스타일 설정 (굵게, 배경색)
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a2744');   // 네이비 배경
  headerRange.setFontColor('#ffffff');    // 흰색 글씨
  headerRange.setHorizontalAlignment('center');

  // 열 너비 자동 조정
  sheet.autoResizeColumns(1, headers.length);

  // 첫 번째 행 고정 (스크롤해도 헤더 보임)
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ 헤더 설정이 완료되었습니다!');
}
