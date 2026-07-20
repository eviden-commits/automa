/* =========================================================================
   Auth.gs
   관리자 비밀번호 설정 및 검증
   ========================================================================= */

/**
 * 비밀번호가 스크립트 속성에 올바르게 설정되어 있는지, 그리고 입력된 비밀번호와 일치하는지 확인합니다.
 * 만약 비밀번호가 설정되어 있지 않은 최초 실행 상태라면, false를 반환하고
 * 프론트엔드에서 초기 비밀번호를 설정하도록 유도합니다.
 */
function checkPassword_(password) {
  var props = PropertiesService.getScriptProperties();
  var current = props.getProperty('ADMIN_PASSWORD');
  
  // 비밀번호가 아예 설정되지 않은 최초 상태
  if (!current) {
    return false; 
  }
  
  return !!password && String(password) === current;
}

/**
 * 최초 실행 여부(비밀번호 설정 여부)를 확인합니다.
 */
function isFirstSetup_() {
  var current = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !current;
}

/**
 * 새로운 관리자 비밀번호를 설정합니다.
 */
function setAdminPassword_(newPassword) {
  var newPwd = String(newPassword || '').trim();
  if (!newPwd) {
    throw new Error('비밀번호는 빈 값일 수 없습니다.');
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPwd);
  return true;
}
