/* =========================================================================
   RecipientManager.gs
   수신자(이메일 주소록) 조회, 대량 추가, 선택 대량 삭제
   ========================================================================= */

/**
 * 수신자 목록을 조회합니다.
 */
function getRecipients_(ss) {
  var sheet = ss.getSheetByName('수신자');
  if (!sheet) return [];
  
  var values = sheet.getDataRange().getValues();
  var recipients = [];
  
  // 첫 번째 행은 헤더
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue; // 이메일 주소가 없으면 스킵
    
    recipients.push({
      name: String(row[0] || ''),
      email: String(row[1]),
      memo: String(row[2] || ''),
      createdAt: row[3] ? Utilities.formatDate(new Date(row[3]), TIME_ZONE, "yyyy-MM-dd HH:mm:ss") : ''
    });
  }
  return recipients;
}

/**
 * 수신자를 대량 등록합니다. (단일 등록 시에도 배열에 담아서 처리 가능)
 * 이메일을 기준으로 중복 검사를 수행하며, 이미 존재하는 이메일일 경우 이름과 비고를 업데이트합니다.
 */
function addRecipientsBulk_(ss, data) {
  var list = data.recipients; // [{name, email, memo}, ...]
  if (!list || !Array.isArray(list)) {
    throw new Error('등록할 수신자 목록이 올바르지 않습니다.');
  }

  var sheet = ss.getSheetByName('수신자');
  if (!sheet) {
    getSpreadsheet();
    sheet = ss.getSheetByName('수신자');
  }

  var values = sheet.getDataRange().getValues();
  var emailRowMap = {}; // 이메일별 기존 행 인덱스 저장 (1-indexed)
  
  for (var i = 1; i < values.length; i++) {
    var email = String(values[i][1]).toLowerCase().trim();
    if (email) {
      emailRowMap[email] = i + 1;
    }
  }

  var nowStr = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");

  // 스프레드시트 갱신 효율을 위해 하나씩 setValue를 하기보다는 필요한 경우 추가하고 덮어씁니다.
  list.forEach(function(item) {
    var name = String(item.name || '').trim();
    var email = String(item.email || '').toLowerCase().trim();
    var memo = String(item.memo || '').trim();

    if (!email) return; // 이메일이 없으면 패스
    
    // 이메일 유효성 체크 정규식
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      console.warn("유효하지 않은 이메일 건너뜀: " + email);
      return;
    }

    if (emailRowMap[email]) {
      // 기존에 존재하면 업데이트
      var rNum = emailRowMap[email];
      if (name) sheet.getRange(rNum, 1).setValue(name);
      sheet.getRange(rNum, 3).setValue(memo);
    } else {
      // 존재하지 않으면 새 행 추가
      sheet.appendRow([name, email, memo, nowStr]);
      // 방금 추가된 행을 맵에 업데이트 (동일 배치 내 중복 방지)
      emailRowMap[email] = sheet.getLastRow();
    }
  });

  return true;
}

/**
 * 선택된 수신자들을 대량 삭제합니다.
 * data.emails: [email1, email2, ...]
 */
function deleteRecipientsBulk_(ss, data) {
  var emailsToDelete = data.emails;
  if (!emailsToDelete || !Array.isArray(emailsToDelete)) {
    throw new Error('삭제할 이메일 목록이 올바르지 않습니다.');
  }

  var sheet = ss.getSheetByName('수신자');
  if (!sheet) return false;

  var values = sheet.getDataRange().getValues();
  
  // 소문자 셋으로 변환하여 대소문자 구분 없이 매칭
  var emailSet = {};
  emailsToDelete.forEach(function(email) {
    emailSet[String(email).toLowerCase().trim()] = true;
  });

  // 행을 삭제할 때는 아래에서부터 지워 올라가야 인덱스가 흐트러지지 않습니다.
  for (var i = values.length - 1; i >= 1; i--) {
    var email = String(values[i][1]).toLowerCase().trim();
    if (emailSet[email]) {
      sheet.deleteRow(i + 1);
    }
  }

  return true;
}
