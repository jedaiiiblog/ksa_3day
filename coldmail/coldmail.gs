/**
 * 콜드메일 자동 발송 도구
 * 시트에 입력한 명단으로 콜드메일을 보내고, 답장이 없으면 3일 뒤 후속 메일을 한 번 더 보낸다.
 */

// 시트 헤더 목록 (열 순서)
var HEADERS = ['보내는 사람', '이름', '회사', '직책', '이메일', '메모(나만 봄)', '상태'];

// 상태 값에 쓰이는 문구들
var STATUS_SENT_PREFIX = '발송완료';
var STATUS_REPLIED = '답장옴';
var STATUS_FOLLOWUP_DONE = '후속완료';

// 후속 메일을 보내기까지 기다리는 일수
var FOLLOW_UP_WAIT_DAYS = 3;

// 자동 후속 메일 트리거에 사용할 함수 이름
var FOLLOW_UP_FUNCTION_NAME = 'sendFollowUps';


/**
 * 1. 시트를 열었을 때 상단에 '콜드메일' 메뉴를 만든다.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('콜드메일')
    .addItem('처음 설정하기 (헤더 만들기)', 'setupColdMailSheet')
    .addItem('지금 발송하기', 'sendColdMailsNow')
    .addItem('자동 후속 메일 켜기', 'enableFollowUpTrigger')
    .addItem('자동 후속 메일 끄기', 'disableFollowUpTrigger')
    .addToUi();
}


/**
 * 2. 처음 설정하기: 1행에 헤더를 만들고 서식을 적용한다.
 */
function setupColdMailSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);

  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d9d9');
  sheet.setFrozenRows(1);

  SpreadsheetApp.getActiveSpreadsheet().toast('헤더 설정이 완료되었습니다.', '콜드메일', 5);
}


/**
 * 3. 메일 제목과 본문을 만드는 함수.
 * placeholders 객체의 값으로 {{sender}}, {{name}}, {{company}}, {{title}}을 치환한다.
 * @param {{sender:string, name:string, company:string, title:string}} placeholders
 * @return {{subject:string, body:string}}
 */
function buildMail(placeholders) {
  // 수강생 수정: 메일 제목을 원하는 문구로 바꿔서 사용하세요.
  var subjectTemplate = '{{company}} {{name}}{{title}}님께 - 업무 자동화 제안드립니다';

  // 수강생 수정: 메일 본문을 원하는 문구로 바꿔서 사용하세요.
  var bodyTemplate =
    '{{name}} {{title}}님, 안녕하세요.\n' +
    '{{sender}}입니다.\n\n' +
    '{{company}}에서 반복적으로 처리하시는 업무 중 자동화로 시간을 절약할 수 있는 부분이 있을지 궁금해 연락드립니다.\n' +
    '데이터 정리, 보고서 작성, 이메일 발송처럼 손이 많이 가는 작업을 자동화해 드린 경험이 있습니다.\n\n' +
    '지금 보고 계신 이 메일도 구글 시트와 연동해 자동으로 발송된 것입니다.\n\n' +
    '혹시 관심이 있으시다면 편하신 시간에 짧게 통화나 미팅으로 소개해 드리고 싶습니다.\n' +
    '감사합니다.\n\n' +
    '{{sender}} 드림';

  var subject = applyPlaceholders(subjectTemplate, placeholders);
  var body = applyPlaceholders(bodyTemplate, placeholders);

  return { subject: subject, body: body };
}


/**
 * buildMail에서 사용하는 자리표시자 치환 헬퍼 함수.
 * @param {string} template
 * @param {{sender:string, name:string, company:string, title:string}} placeholders
 * @return {string}
 */
function applyPlaceholders(template, placeholders) {
  return template
    .replace(/{{sender}}/g, placeholders.sender)
    .replace(/{{name}}/g, placeholders.name)
    .replace(/{{company}}/g, placeholders.company)
    .replace(/{{title}}/g, placeholders.title);
}


/**
 * 4. 지금 발송하기: 확인 팝업을 먼저 보여주고, 예를 눌렀을 때만 메일을 보낸다.
 */
function sendColdMailsNow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var colIndex = getColumnIndexMap(data[0]);

  var remainingQuota = MailApp.getRemainingDailyQuota();

  // 발송 대상(상태가 비어 있거나 발송완료가 아닌 행)과 제외 대상(이미 발송완료)을 분류
  var targetRows = [];
  var alreadySentCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = String(row[colIndex['상태']] || '');

    if (status.indexOf(STATUS_SENT_PREFIX) === 0) {
      alreadySentCount++;
      continue;
    }

    var email = String(row[colIndex['이메일']] || '').trim();
    if (email === '') {
      continue;
    }

    targetRows.push(i);
  }

  var sendableCount = Math.min(targetRows.length, remainingQuota);

  // 팝업에 보여줄 안내 메시지 구성
  var message =
    '이번에 보낼 수 있는 메일 건수: ' + targetRows.length + '건\n' +
    '오늘 남은 발송 한도: ' + remainingQuota + '건\n' +
    '이미 발송완료라서 제외되는 건수: ' + alreadySentCount + '건';

  if (targetRows.length > remainingQuota) {
    message += '\n\n남은 한도보다 보낼 메일이 많아 ' + sendableCount + '건만 발송됩니다.';
  }

  message += '\n\n메일을 발송하시겠습니까?';

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('메일 발송 확인', message, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  var sentCount = 0;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (var j = 0; j < targetRows.length; j++) {
    if (sentCount >= remainingQuota) {
      break;
    }

    var rowIndex = targetRows[j];
    var rowData = data[rowIndex];

    var placeholders = {
      sender: String(rowData[colIndex['보내는 사람']] || ''),
      name: String(rowData[colIndex['이름']] || ''),
      company: String(rowData[colIndex['회사']] || ''),
      title: String(rowData[colIndex['직책']] || '')
    };
    var email = String(rowData[colIndex['이메일']] || '').trim();

    var mail = buildMail(placeholders);
    MailApp.sendEmail(email, mail.subject, mail.body);

    var sheetRowNumber = rowIndex + 1; // 시트의 실제 행 번호 (헤더 포함)
    sheet.getRange(sheetRowNumber, colIndex['상태'] + 1).setValue(STATUS_SENT_PREFIX + ' ' + today);

    sentCount++;
    Utilities.sleep(2000);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(sentCount + '건 발송 완료되었습니다.', '콜드메일', 5);
}


/**
 * 시트 헤더 이름을 열 인덱스(0부터 시작)로 매핑한다.
 * @param {Array<string>} headerRow
 * @return {Object<string, number>}
 */
function getColumnIndexMap(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    map[headerRow[i]] = i;
  }
  return map;
}


/**
 * 5. 자동 후속 메일 보내기.
 * 처음 메일을 보낸 지 3일 이상 지났고 아직 후속 처리를 하지 않은 행을 확인해
 * 답장이 왔으면 상태를 '답장옴'으로, 없으면 후속 메일을 보내고 '후속완료'로 바꾼다.
 */
function sendFollowUps() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var colIndex = getColumnIndexMap(data[0]);

  var timeZone = Session.getScriptTimeZone();
  var todayDate = new Date();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = String(row[colIndex['상태']] || '');

    // 이미 답장옴 또는 후속완료 상태면 건너뜀
    if (status.indexOf(STATUS_REPLIED) === 0 || status.indexOf(STATUS_FOLLOWUP_DONE) === 0) {
      continue;
    }

    // 발송완료 상태가 아니면(아직 첫 메일을 보내지 않았으면) 건너뜀
    if (status.indexOf(STATUS_SENT_PREFIX) !== 0) {
      continue;
    }

    var sentDateText = status.substring(STATUS_SENT_PREFIX.length).trim();
    var sentDate = new Date(sentDateText);
    if (isNaN(sentDate.getTime())) {
      continue;
    }

    var daysPassed = Math.floor((todayDate.getTime() - sentDate.getTime()) / (1000 * 60 * 60 * 24));
    if (daysPassed < FOLLOW_UP_WAIT_DAYS) {
      continue;
    }

    var email = String(row[colIndex['이메일']] || '').trim();
    if (email === '') {
      continue;
    }

    var sheetRowNumber = i + 1;
    var hasReply = checkHasReplyFrom(email);

    if (hasReply) {
      sheet.getRange(sheetRowNumber, colIndex['상태'] + 1).setValue(STATUS_REPLIED);
      continue;
    }

    var placeholders = {
      sender: String(row[colIndex['보내는 사람']] || ''),
      name: String(row[colIndex['이름']] || ''),
      company: String(row[colIndex['회사']] || ''),
      title: String(row[colIndex['직책']] || '')
    };

    var followUpMail = buildFollowUpMail(placeholders);
    MailApp.sendEmail(email, followUpMail.subject, followUpMail.body);

    sheet.getRange(sheetRowNumber, colIndex['상태'] + 1).setValue(STATUS_FOLLOWUP_DONE);

    Utilities.sleep(2000);
  }
}


/**
 * 해당 이메일 주소에서 나에게 보낸 메일이 있는지 Gmail에서 검색한다.
 * @param {string} email
 * @return {boolean} 답장이 있으면 true
 */
function checkHasReplyFrom(email) {
  var searchQuery = 'from:' + email;
  var threads = GmailApp.search(searchQuery, 0, 1);
  return threads.length > 0;
}


/**
 * 후속 메일 제목과 본문을 만드는 함수.
 * @param {{sender:string, name:string, company:string, title:string}} placeholders
 * @return {{subject:string, body:string}}
 */
function buildFollowUpMail(placeholders) {
  // 수강생 수정: 후속 메일 제목을 원하는 문구로 바꿔서 사용하세요.
  var subjectTemplate = '[다시 연락드립니다] {{company}} {{name}}{{title}}님께 업무 자동화 제안';

  // 수강생 수정: 후속 메일 본문을 원하는 문구로 바꿔서 사용하세요.
  var bodyTemplate =
    '{{name}} {{title}}님, 안녕하세요.\n' +
    '{{sender}}입니다.\n\n' +
    '지난번에 업무 자동화 관련하여 메일을 드렸는데, 혹시 확인해 보셨는지 궁금해 다시 연락드립니다.\n' +
    '필요하신 부분이 있으시면 편하신 때에 말씀해 주세요.\n\n' +
    '감사합니다.\n\n' +
    '{{sender}} 드림';

  var subject = applyPlaceholders(subjectTemplate, placeholders);
  var body = applyPlaceholders(bodyTemplate, placeholders);

  return { subject: subject, body: body };
}


/**
 * 6. 자동 후속 메일 켜기: 기존 트리거를 지우고 매일 오전 9시에 실행되는 트리거를 새로 만든다.
 */
function enableFollowUpTrigger() {
  removeFollowUpTriggers();

  ScriptApp.newTrigger(FOLLOW_UP_FUNCTION_NAME)
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('자동 후속 메일이 켜졌습니다. (매일 오전 9시 실행)', '콜드메일', 5);
}


/**
 * 7. 자동 후속 메일 끄기: sendFollowUps를 실행하는 트리거를 모두 삭제한다.
 */
function disableFollowUpTrigger() {
  removeFollowUpTriggers();

  SpreadsheetApp.getActiveSpreadsheet().toast('자동 후속 메일이 꺼졌습니다.', '콜드메일', 5);
}


/**
 * sendFollowUps 함수를 실행하는 트리거를 모두 찾아 삭제하는 헬퍼 함수.
 */
function removeFollowUpTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === FOLLOW_UP_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
