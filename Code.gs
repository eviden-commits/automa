/* =========================================================================
   Code.gs
   진입점(doGet/doPost) 라우팅
   모든 민감한 비즈니스 로직과 데이터 조회/변경은 POST 요청을 통해서만 처리하여
   비밀번호가 URL에 노출되는 것을 완전히 차단합니다.
   ========================================================================= */

/**
 * GET 요청 처리 (웹앱 접속 테스트 및 비밀번호 설정 여부 체크만 지원)
 */
function doGet(e) {
  try {
    // 최초 설치 상태(비밀번호 미설정) 여부만 안전하게 체크하여 반환
    var firstSetup = isFirstSetup_();
    return outputJson_({ 
      status: "ok", 
      isFirstSetup: firstSetup,
      message: "Automail GAS 백엔드가 정상 작동 중입니다. 모든 데이터 통신은 POST를 사용하십시오."
    });
  } catch (error) {
    return outputJson_({ status: "error", message: error.toString() });
  }
}

/**
 * POST 요청 처리 (인증 확인, 데이터 로드, 데이터 추가/수정/삭제 등 모든 비즈니스 로직)
 * 비밀번호를 Request Body에 포함시켜 안전하게 검증합니다.
 */
function doPost(e) {
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    
    // 1. 최초 설치 비밀번호 등록
    if (data.action === 'setAdminPassword') {
      if (!isFirstSetup_()) {
        return outputJson_({ status: "error", message: "이미 비밀번호가 설정되어 있습니다." });
      }
      setAdminPassword_(data.newPassword);
      return outputJson_({ status: "ok", message: "관리자 비밀번호가 성공적으로 설정되었습니다." });
    }

    // 2. 관리자 인증 검증 (모든 요청에서 비밀번호가 필수로 동반됨)
    if (!checkPassword_(data.password)) {
      return outputJson_({ status: "auth_error", message: "비밀번호가 올바르지 않습니다." });
    }

    var ss = getSpreadsheet();
    var result = { status: "ok" };

    // 3. 액션 라우팅
    switch (data.action) {
      // 3-1. 데이터 전체 로드 (Read)
      case 'loadData':
        var reminders = getReminders_(ss);
        var recipients = getRecipients_(ss);
        
        var activeReminders = [];
        var finishedReminders = [];
        
        reminders.forEach(function(r) {
          if (r.status === 'active') {
            activeReminders.push(r);
          } else {
            finishedReminders.push(r);
          }
        });
        
        result.spreadsheetUrl = ss.getUrl();
        result.activeReminders = activeReminders;
        result.finishedReminders = finishedReminders;
        result.recipients = recipients;
        result.lastSync = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
        break;

      // 3-2. 리마인더 관리 (CUD)
      case 'addReminder':
        var newId = addReminder_(ss, data);
        result.id = newId;
        break;
        
      case 'updateReminder':
        updateReminder_(ss, data);
        break;
        
      case 'deleteReminder':
        deleteReminder_(ss, data);
        break;
        
      case 'toggleSuspendReminder':
        var newStatus = toggleSuspendReminder_(ss, data);
        result.newStatus = newStatus;
        break;

      // 3-3. 수신자 관리
      case 'addRecipientsBulk':
        addRecipientsBulk_(ss, data);
        break;
        
      case 'deleteRecipientsBulk':
        deleteRecipientsBulk_(ss, data);
        break;

      // 3-4. 시스템 제어
      case 'installTrigger':
        installHourlyTrigger();
        result.message = "배치 트리거가 정상적으로 설치되었습니다.";
        break;
        
      case 'runTriggerManual':
        checkAndSendReminders_();
        result.message = "배치가 수동으로 실행되었습니다. 발송로그를 확인하세요.";
        break;
        
      default:
        throw new Error('지원하지 않는 액션입니다: ' + data.action);
    }

    return outputJson_(result);
  } catch (error) {
    return outputJson_({ status: "error", message: error.toString() });
  }
}

/**
 * JSON 규격에 맞춰 ContentService를 생성하여 반환합니다.
 */
function outputJson_(obj) {
  var outputText = JSON.stringify(obj);
  return ContentService.createTextOutput(outputText).setMimeType(ContentService.MimeType.JSON);
}
