/* =========================================================================
   TriggerHandler.gs
   매시간 구글 서버 타임(KST) 기준으로 스케쥴 검사 및 리마인드 메일 발송
   ========================================================================= */

/**
 * 매시간 실행되어 조건에 맞는 리마인더를 발송하는 메인 배치 함수입니다.
 */
function checkAndSendReminders_() {
  var ss = getSpreadsheet();
  var reminders = getReminders_(ss);
  
  // 1. KST 기준 현재 시간 정보 구하기
  var kstStr = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
  var parts = kstStr.split(" ");
  var dateParts = parts[0].split("-");
  var timeParts = parts[1].split(":");
  
  var year = Number(dateParts[0]);
  var month = Number(dateParts[1]);
  var day = Number(dateParts[2]);
  var hour = Number(timeParts[0]); // 0 ~ 23
  
  var todayStr = parts[0]; // "YYYY-MM-DD"
  var todayDate = new Date(year, month - 1, day);
  
  var dayOfWeek = todayDate.getDay(); // 0(일) ~ 6(토)
  var isoDay = dayOfWeek === 0 ? 7 : dayOfWeek; // 1(월) ~ 7(일)

  // 2. 이번 달의 마지막 날인지 체크
  var tomorrow = new Date(year, month - 1, day + 1);
  var isLastDayOfMonth = (tomorrow.getDate() === 1);

  // 3. 리마인더 시트의 각 행을 검사하며 발송 대상 판별
  var reminderSheet = ss.getSheetByName('리마인더');
  var rawValues = reminderSheet.getDataRange().getValues();
  
  for (var i = 0; i < reminders.length; i++) {
    var reminder = reminders[i];
    
    // 활성 상태가 아니면 건너뜀
    if (reminder.status !== 'active') continue;
    
    // 발송 시간(시)이 일치하지 않으면 건너뜀
    if (reminder.hour !== hour) continue;
    
    // 주기 조건 만족 여부 확인
    var shouldSend = false;
    
    switch (reminder.interval) {
      case 'daily':
        shouldSend = true;
        break;
        
      case 'weekly':
        if (reminder.weeklyDay === isoDay) {
          shouldSend = true;
        }
        break;
        
      case 'biweekly':
        // 요일이 맞아야 함
        if (reminder.weeklyDay === isoDay) {
          if (!reminder.lastSentAt) {
            // 발송 이력이 없으면 즉시 발송 가능
            shouldSend = true;
          } else {
            // 마지막 발송일로부터 최소 13일이 경과했는지 체크 (2주 주기)
            var lastSentDate = new Date(reminder.lastSentAt.split(" ")[0]);
            var diffDays = Math.floor((todayDate - lastSentDate) / (1000 * 60 * 60 * 24));
            if (diffDays >= 13) {
              shouldSend = true;
            }
          }
        }
        break;
        
      case 'month_start':
        if (day === 1) {
          shouldSend = true;
        }
        break;
        
      case 'month_end':
        if (isLastDayOfMonth) {
          shouldSend = true;
        }
        break;
        
      case 'custom':
        if (reminder.customDate === todayStr) {
          shouldSend = true;
        }
        break;
    }
    
    if (shouldSend) {
      // 4. 메일 발송 수행
      var sendStatus = sendReminderMailDirect_(reminder);
      
      // 5. 시트 데이터 업데이트 (스프레드시트는 2행부터 시작하므로 인덱스는 i + 2)
      var rowNum = i + 2;
      
      // 최근 발송일시 기록
      reminderSheet.getRange(rowNum, 11).setValue(kstStr);
      
      // 루프가 없는 단발성 리마인더의 경우 '완료(done)' 처리
      if (!reminder.isLoop) {
        reminderSheet.getRange(rowNum, 10).setValue('done');
      }
      
      // 6. 로그 기록
      logSentHistory_(ss, reminder, sendStatus, kstStr);
    }
  }
}

/**
 * 리마인더 메일을 실제로 발송합니다.
 */
function sendReminderMailDirect_(reminder) {
  if (!reminder.recipients || reminder.recipients.length === 0) {
    console.warn("발송 대상 이메일이 없어 메일을 발송하지 못했습니다. ID: " + reminder.id);
    return 'FAIL (No Recipients)';
  }
  
  // HTML 메일 템플릿 생성 (Apple / McKinsey 스타일의 고급스러운 디자인)
  var contentHtml = reminder.content.replace(/\n/g, '<br>');
  
  var subject = '[리마인더] ' + reminder.title;
  var htmlBody = 
    '<div style="background-color:#F5F5F7; padding:40px 20px; font-family:-apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; color:#1D1D1F;">' +
      '<div style="max-width:600px; margin:0 auto; background-color:#FFFFFF; border-radius:18px; box-shadow:0 4px 20px rgba(0,0,0,0.05); overflow:hidden; border:1px solid rgba(60,60,67,0.06);">' +
        // 헤더 영역
        '<div style="background-color:#1D1D1F; padding:24px 32px; color:#FFFFFF;">' +
          '<div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.12em; color:#AEAEB2; margin-bottom:4px;">Automated Schedule Reminder</div>' +
          '<div style="font-size:22px; font-weight:700; letter-spacing:-0.02em;">' + escapeHtml_(reminder.title) + '</div>' +
        '</div>' +
        // 본문 영역
        '<div style="padding:32px;">' +
          '<div style="font-size:15px; line-height:1.65; color:#3C3C43; margin-bottom:28px;">' +
            contentHtml +
          '</div>' +
          
          // 메타 정보 테이블
          '<div style="background-color:#F5F5F7; border-radius:12px; padding:18px 24px; font-size:13px; border:0.5px solid rgba(60,60,67,0.12);">' +
            '<table style="width:100%; border-collapse:collapse;">' +
              '<tr>' +
                '<td style="color:#8E8E93; padding:6px 0; width:90px; font-weight:600;">발송 주기</td>' +
                '<td style="color:#1D1D1F; padding:6px 0; font-weight:600;">' + getIntervalLabel_(reminder) + '</td>' +
              '</tr>' +
              '<tr>' +
                '<td style="color:#8E8E93; padding:6px 0; font-weight:600;">발송 시각</td>' +
                '<td style="color:#1D1D1F; padding:6px 0; font-weight:500;">' + reminder.hour + '시 (KST)</td>' +
              '</tr>' +
              '<tr>' +
                '<td style="color:#8E8E93; padding:6px 0; font-weight:600;">반복 여부</td>' +
                '<td style="color:#1D1D1F; padding:6px 0; font-weight:500;">' + (reminder.isLoop ? '반복 설정 (Loop)' : '1회성 발송 (Single)') + '</td>' +
              '</tr>' +
            '</table>' +
          '</div>' +
        '</div>' +
        
        // 푸터 영역
        '<div style="background-color:#F5F5F7; padding:20px 32px; border-top:0.5px solid rgba(60,60,67,0.08); text-align:center; font-size:11px; color:#8E8E93;">' +
          '본 메일은 설정된 스케쥴에 따라 구글 서버 타임 기준 자동 발송되었습니다.<br>' +
          'Copyright &copy; SeBang Tech Quality Safety Health Dept. All rights reserved.' +
        '</div>' +
      '</div>' +
    '</div>';
    
  try {
    MailApp.sendEmail({
      to: reminder.recipients.join(','),
      subject: subject,
      htmlBody: htmlBody
    });
    return 'SUCCESS';
  } catch (e) {
    console.error("이메일 발송 에러: " + e.toString());
    return 'FAIL (' + e.toString() + ')';
  }
}

/**
 * 발송 결과 로그를 시트에 기록합니다.
 */
function logSentHistory_(ss, reminder, status, kstStr) {
  var logSheet = ss.getSheetByName('발송로그');
  if (!logSheet) return;
  
  var logId = Utilities.getUuid();
  var reminderId = reminder.id;
  var title = reminder.title;
  var recipients = reminder.recipients.join(', ');
  
  logSheet.appendRow([logId, reminderId, title, recipients, kstStr, status]);
}

/**
 * 발송 주기를 한글 레이블로 포맷팅합니다.
 */
function getIntervalLabel_(reminder) {
  var map = {
    'daily': '매일',
    'weekly': '매주 요일',
    'biweekly': '격주 요일',
    'month_start': '월초 (1일)',
    'month_end': '월말 (말일)',
    'custom': '지정일 (커스텀)'
  };
  
  var label = map[reminder.interval] || reminder.interval;
  
  if (reminder.interval === 'weekly' || reminder.interval === 'biweekly') {
    var days = ['', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일'];
    label += ' (' + (days[reminder.weeklyDay] || '') + ')';
  } else if (reminder.interval === 'custom') {
    label += ' (' + reminder.customDate + ')';
  }
  
  return label;
}

/**
 * HTML 특수문자를 이스케이프합니다.
 */
function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * 매시간 도는 Google Apps Script 시간 트리거를 설치합니다.
 * 기존에 설정된 checkAndSendReminders_ 트리거가 있으면 삭제 후 재설치합니다.
 */
function installHourlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'checkAndSendReminders_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger('checkAndSendReminders_')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log("매시간 배치 트리거 설치 완료.");
}
