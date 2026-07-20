/* =========================================================================
   Config.gs
   글로벌 설정 및 데이터베이스(스프레드시트) 초기화/관리
   ========================================================================= */

// 한국 표준시(KST) 타임존 설정
var TIME_ZONE = "GMT+9";

/**
 * 활성화된 스프레드시트 객체를 가져옵니다.
 * 지정된 SPREADSHEET_ID가 없거나 유효하지 않으면 새 스프레드시트를 생성하여 속성에 저장합니다.
 */
function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      console.warn("기존 스프레드시트를 열 수 없어 재설정합니다: " + e.toString());
      ssId = null;
    }
  }

  if (!ssId || !ss) {
    // 새 스프레드시트 생성
    ss = SpreadsheetApp.create("스케쥴관리_리마인더_DB");
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    console.log("새 스프레드시트 생성 완료. ID: " + ssId);
    
    // 기본 시트(시트1) 삭제 혹은 이름 변경하여 초기화 진행
    initSpreadsheet_(ss);
  }

  return ss;
}

/**
 * 스프레드시트 초기화 및 시트 구조 생성
 */
function initSpreadsheet_(ss) {
  // 1. 리마인더 시트
  var reminderSheet = ss.getSheetByName('리마인더');
  if (!reminderSheet) {
    reminderSheet = ss.insertSheet('리마인더');
    reminderSheet.appendRow([
      'ID', '제목', '내용', '발송대상', '발송주기', 
      '커스텀날짜', '요일선택', '발송시간', '루프여부', 
      '상태', '최근발송일시', '생성일시'
    ]);
    reminderSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#f3f3f3');
  }

  // 2. 수신자 시트
  var recipientSheet = ss.getSheetByName('수신자');
  if (!recipientSheet) {
    recipientSheet = ss.insertSheet('수신자');
    recipientSheet.appendRow(['이름', '이메일', '비고', '생성일시']);
    recipientSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
  }

  // 3. 발송로그 시트
  var logSheet = ss.getSheetByName('발송로그');
  if (!logSheet) {
    logSheet = ss.insertSheet('발송로그');
    logSheet.appendRow(['로그ID', '리마인더ID', '리마인더제목', '수신자이메일', '발송일시', '상태']);
    logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f3f3');
  }

  // 기본 생성되는 '시트1'이 남아있으면 제거
  var defaultSheet = ss.getSheetByName('시트1');
  if (defaultSheet) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch(e) {}
  }
}
