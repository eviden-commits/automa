/* =========================================================================
   ReminderManager.gs
   리마인더 데이터 조회, 추가, 수정, 삭제, 강제종료 상태 전환
   ========================================================================= */

/**
 * 리마인더 목록을 가져옵니다.
 */
function getReminders_(ss) {
  var sheet = ss.getSheetByName('리마인더');
  if (!sheet) return [];
  
  var values = sheet.getDataRange().getValues();
  var reminders = [];
  
  // 첫 번째 행은 헤더
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue; // ID가 없는 빈 행 스킵
    
    reminders.push({
      id: String(row[0]),
      title: String(row[1] || ''),
      content: String(row[2] || ''),
      recipients: String(row[3] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
      interval: String(row[4] || 'daily'),
      customDate: row[5] ? Utilities.formatDate(new Date(row[5]), TIME_ZONE, "yyyy-MM-dd") : '',
      weeklyDay: row[6] !== '' ? Number(row[6]) : '', // 1(월)~7(일)
      hour: row[7] !== '' ? Number(row[7]) : 9, // 발송 시간(시, 0~23)
      isLoop: row[8] === true || String(row[8]).toUpperCase() === 'TRUE',
      status: String(row[9] || 'active'),
      lastSentAt: row[10] ? Utilities.formatDate(new Date(row[10]), TIME_ZONE, "yyyy-MM-dd HH:mm:ss") : '',
      createdAt: row[11] ? Utilities.formatDate(new Date(row[11]), TIME_ZONE, "yyyy-MM-dd HH:mm:ss") : ''
    });
  }
  return reminders;
}

/**
 * 새 리마인더를 추가합니다.
 */
function addReminder_(ss, data) {
  var sheet = ss.getSheetByName('리마인더');
  if (!sheet) {
    getSpreadsheet(); // 시트 재생성 시도
    sheet = ss.getSheetByName('리마인더');
  }
  
  var id = Utilities.getUuid();
  var title = String(data.title || '').trim();
  var content = String(data.content || '').trim();
  var recipients = Array.isArray(data.recipients) ? data.recipients.join(', ') : String(data.recipients || '');
  var interval = String(data.interval || 'daily');
  var customDate = data.customDate || '';
  var weeklyDay = data.weeklyDay !== undefined && data.weeklyDay !== null ? Number(data.weeklyDay) : '';
  var hour = data.hour !== undefined && data.hour !== null ? Number(data.hour) : 9;
  var isLoop = data.isLoop === true || String(data.isLoop).toUpperCase() === 'TRUE';
  var status = 'active'; // 기본 상태는 활성
  var lastSentAt = '';
  var createdAt = Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
  
  if (!title) throw new Error('제목을 입력해주세요.');
  
  sheet.appendRow([
    id, title, content, recipients, interval, 
    customDate, weeklyDay, hour, isLoop, 
    status, lastSentAt, createdAt
  ]);
  
  return id;
}

/**
 * 기존 리마인더 정보를 수정합니다.
 */
function updateReminder_(ss, data) {
  var sheet = ss.getSheetByName('리마인더');
  if (!sheet) return false;
  
  var id = String(data.id || '');
  if (!id) throw new Error('수정할 리마인더 ID가 필요합니다.');
  
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      var rowNum = i + 1;
      
      var title = String(data.title || '').trim();
      var content = String(data.content || '').trim();
      var recipients = Array.isArray(data.recipients) ? data.recipients.join(', ') : String(data.recipients || '');
      var interval = String(data.interval || 'daily');
      var customDate = data.customDate || '';
      var weeklyDay = data.weeklyDay !== undefined && data.weeklyDay !== null ? Number(data.weeklyDay) : '';
      var hour = data.hour !== undefined && data.hour !== null ? Number(data.hour) : 9;
      var isLoop = data.isLoop === true || String(data.isLoop).toUpperCase() === 'TRUE';
      var status = String(data.status || values[i][9]); // 상태가 넘어오면 변경, 없으면 유지
      
      if (!title) throw new Error('제목을 입력해주세요.');
      
      sheet.getRange(rowNum, 2).setValue(title);
      sheet.getRange(rowNum, 3).setValue(content);
      sheet.getRange(rowNum, 4).setValue(recipients);
      sheet.getRange(rowNum, 5).setValue(interval);
      sheet.getRange(rowNum, 6).setValue(customDate);
      sheet.getRange(rowNum, 7).setValue(weeklyDay);
      sheet.getRange(rowNum, 8).setValue(hour);
      sheet.getRange(rowNum, 9).setValue(isLoop);
      sheet.getRange(rowNum, 10).setValue(status);
      
      return true;
    }
  }
  throw new Error('리마인더를 찾을 수 없습니다.');
}

/**
 * 리마인더를 삭제합니다.
 */
function deleteReminder_(ss, data) {
  var sheet = ss.getSheetByName('리마인더');
  if (!sheet) return false;
  
  var id = String(data.id || '');
  if (!id) throw new Error('삭제할 리마인더 ID가 필요합니다.');
  
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('리마인더를 찾을 수 없습니다.');
}

/**
 * 리마인더 상태를 강제종료(suspended) 혹은 활성(active) 상태로 토글 전환합니다.
 * 루프가 없는 단발성 리마인더가 완료(done) 상태인 경우에는 강제종료 처리가 아닌 다시 활성화할 수 있도록 변경합니다.
 */
function toggleSuspendReminder_(ss, data) {
  var sheet = ss.getSheetByName('리마인더');
  if (!sheet) return false;
  
  var id = String(data.id || '');
  if (!id) throw new Error('상태를 전환할 리마인더 ID가 필요합니다.');
  
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      var rowNum = i + 1;
      var currentStatus = String(values[i][9] || 'active');
      var newStatus = 'active';
      
      if (currentStatus === 'active') {
        newStatus = 'suspended'; // 강제종료
      } else {
        newStatus = 'active'; // 재활성화
      }
      
      sheet.getRange(rowNum, 10).setValue(newStatus);
      return newStatus;
    }
  }
  throw new Error('리마인더를 찾을 수 없습니다.');
}
