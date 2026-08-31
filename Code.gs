// ─────────────────────────────────────────────────────────────────────────────
//  비즈 아카데미 예약 시스템  |  Google Apps Script  |  Code.gs
// ─────────────────────────────────────────────────────────────────────────────

// 사전 프로필 파일을 모을 구글 드라이브 폴더 ID
var UPLOAD_FOLDER_ID = '1ftuxvLNvJOegM-wFSizLJtAShxSb04eu';

// 스프레드시트 열 번호 (1-based)
var COL = {
  TIMESTAMP:  1,
  NAME:       2,
  STUDENT_ID: 3,
  CONTACT:    4,
  EMAIL:      5,
  DAY:        6,
  DATE:       7,
  TIME:       8,
  PROFESSOR:  9,
  LOCATION:   10,
  STATUS:     11,
  EVENT_ID:   12,
  FILE_ID:    13,
  TOTAL:      13
};

// ─────────────────────────────────────────────────────────────────────────────
//  doGet — 슬롯 조회 / 예약 조회 / 예약 신청 / 취소 / 변경 (GET, CORS 허용)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  var p      = e.parameter || {};
  var action = p.action    || '';
  var result;

  try {
    if      (action === 'getSlots') result = getAvailableSlots(p.date, p.day || '');
    else if (action === 'lookup')   result = lookupReservation(p.studentId, p.email);
    else if (action === 'book')     result = createReservation(p);
    else if (action === 'cancel')   result = cancelReservation(p.studentId, p.email);
    else if (action === 'modify')   result = modifyReservation(p);
    else                            result = { status: 'error', message: '알 수 없는 액션: ' + action };
  } catch (err) {
    Logger.log('doGet 오류 [' + action + ']: ' + err.message);
    result = { status: 'error', message: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────────────
//  doPost — 파일 업로드 전용 (base64 → Google Drive)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action || '';
    var result;

    if (action === 'uploadFile') {
      result = uploadProfileFile(data);
    } else {
      result = { status: 'error', message: '알 수 없는 액션: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost 오류: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  파일 업로드 → Google Drive 폴더 저장
// ─────────────────────────────────────────────────────────────────────────────
function uploadProfileFile(data) {
  var studentId = String(data.studentId || '').trim();
  var name      = String(data.name      || '').trim();
  var date      = String(data.date      || '').trim();
  var origName  = String(data.fileName  || 'profile');
  var fileData  = String(data.fileData  || '');

  if (!fileData) return { status: 'error', message: '파일 데이터가 없습니다.' };

  try {
    var commaIdx = fileData.indexOf(',');
    var meta     = fileData.substring(0, commaIdx);
    var base64   = fileData.substring(commaIdx + 1);
    var mimeType = meta.split(';')[0].split(':')[1] || 'application/octet-stream';

    // 저장 파일명: 멘토링날짜_이름_학번_원본파일명
    var saveAs  = date + '_' + name + '_' + studentId + '_' + origName;
    var bytes   = Utilities.base64Decode(base64);
    var blob    = Utilities.newBlob(bytes, mimeType, saveAs);
    var folder  = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    var file    = folder.createFile(blob);

    Logger.log('파일 업로드 완료: ' + file.getName() + ' (' + file.getId() + ')');
    return { status: 'success', fileId: file.getId(), fileName: saveAs };
  } catch (err) {
    Logger.log('파일 업로드 실패: ' + err.message);
    return { status: 'error', message: '파일 업로드 실패: ' + err.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  신규 예약
// ─────────────────────────────────────────────────────────────────────────────
function createReservation(data) {
  var name      = String(data.name      || '').trim();
  var studentId = String(data.studentId || '').trim();
  var contact   = String(data.contact   || '').trim();
  var email     = String(data.email     || '').trim();
  var day       = String(data.day       || '').trim();
  var date      = String(data.date      || '').trim();
  var time      = String(data.time      || '').trim();
  var fileId    = String(data.fileId    || '').trim();
  var fileName  = String(data.fileName  || '').trim();

  if (!name || !studentId || !email || !date || !time) {
    return { status: 'error', message: '필수 항목이 누락되었습니다.' };
  }

  // 슬롯 중복 체크
  var avail = getAvailableSlots(date);
  Logger.log('슬롯 체크: date=' + date + ' time=' + time + ' taken=' + JSON.stringify(avail.takenSlots));
  if (avail.takenSlots.indexOf(time) !== -1) {
    return { status: 'slotTaken', message: '선택하신 시간대는 이미 예약되어 있습니다.' };
  }

  var routing = getRouting(day);

  // 시트 저장
  saveToSheet(name, studentId, contact, email, day, date, time,
              routing.professorName, routing.location, new Date(), 'active', '', fileId);
  var savedRow = getSheet().getLastRow();
  Logger.log('시트 저장 완료 (행 ' + savedRow + ')');

  var eventDate = buildDateTime(date, time);

  // 확인 이메일 (학생)
  try {
    sendConfirmationEmail(name, email, day, date, time,
                          routing.professorName, routing.location, eventDate,
                          fileId, fileName);
  } catch (err) {
    Logger.log('이메일 발송 실패: ' + err.message);
  }

  // 관리자 알림 이메일
  try {
    sendAdminNotification(name, studentId, contact, email, day, date, time, routing.professorName);
  } catch (err) {
    Logger.log('관리자 알림 실패: ' + err.message);
  }

  return { status: 'success', message: '예약이 완료되었습니다.' };
}

// ─────────────────────────────────────────────────────────────────────────────
//  예약 취소
// ─────────────────────────────────────────────────────────────────────────────
function cancelReservation(studentId, email) {
  var lookup = lookupReservation(studentId, email);
  if (lookup.status === 'notFound') {
    return { status: 'error', message: '예약을 찾을 수 없습니다.' };
  }
  getSheet().getRange(lookup.rowIndex, COL.STATUS).setValue('cancelled');
  deleteFiles(lookup.fileId);
  try {
    sendCancellationEmail(lookup.name, lookup.email, lookup.day, lookup.date,
                          lookup.time, lookup.professor, lookup.location);
  } catch (err) {
    Logger.log('취소 이메일 실패: ' + err.message);
  }
  try {
    sendAdminCancelNotification(lookup.name, lookup.studentId, lookup.contact,
                                lookup.day, lookup.date, lookup.time, lookup.professor);
  } catch (err) {
    Logger.log('관리자 취소 알림 실패: ' + err.message);
  }
  return { status: 'success', message: '예약이 취소되었습니다.' };
}

// ─────────────────────────────────────────────────────────────────────────────
//  예약 변경
// ─────────────────────────────────────────────────────────────────────────────
function modifyReservation(data) {
  var studentId = String(data.studentId || '').trim();
  var email     = String(data.email     || '').trim();
  var newDay    = String(data.day       || '').trim();
  var newDate   = String(data.date      || '').trim();
  var newTime   = String(data.time      || '').trim();

  var lookup = lookupReservation(studentId, email);
  if (lookup.status === 'notFound') {
    return { status: 'error', message: '예약을 찾을 수 없습니다.' };
  }

  var avail = getAvailableSlots(newDate);
  if (avail.takenSlots.indexOf(newTime) !== -1) {
    return { status: 'slotTaken', message: '선택하신 시간대는 이미 예약되어 있습니다.' };
  }

  var routing = getRouting(newDay);
  var sheet   = getSheet();

  // 날짜가 바뀐 경우 Drive 파일명도 변경
  if (lookup.date !== newDate && lookup.fileId) {
    renameFiles(lookup.fileId, lookup.date, newDate, lookup.name, lookup.studentId);
  }

  sheet.getRange(lookup.rowIndex, COL.TIMESTAMP).setValue(
    Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss') + ' (수정됨)');
  sheet.getRange(lookup.rowIndex, COL.DAY).setValue(newDay);
  sheet.getRange(lookup.rowIndex, COL.DATE).setValue(newDate);
  sheet.getRange(lookup.rowIndex, COL.TIME).setValue(newTime);
  sheet.getRange(lookup.rowIndex, COL.PROFESSOR).setValue(routing.professorName);
  sheet.getRange(lookup.rowIndex, COL.LOCATION).setValue(routing.location);

  try {
    sendModificationEmail(lookup.name, email, newDay, newDate, newTime,
                          routing.professorName, routing.location, buildDateTime(newDate, newTime));
  } catch (err) {
    Logger.log('변경 이메일 실패: ' + err.message);
  }
  try {
    sendAdminModifyNotification(lookup.name, lookup.studentId, lookup.contact,
                                lookup.day, lookup.date, lookup.time, lookup.professor,
                                newDay, newDate, newTime, routing.professorName);
  } catch (err) {
    Logger.log('관리자 변경 알림 실패: ' + err.message);
  }
  return { status: 'success', message: '예약이 변경되었습니다.' };
}

// ─────────────────────────────────────────────────────────────────────────────
//  가용 슬롯 조회
// ─────────────────────────────────────────────────────────────────────────────
function getAvailableSlots(date, day) {
  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  var taken   = [];

  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues();
    rows.forEach(function(row) {
      var rowDate   = normalizeDate(row[COL.DATE   - 1]);
      var rowTime   = normalizeTime(row[COL.TIME   - 1]);
      var rowStatus = String(row[COL.STATUS - 1]).trim();
      var rowDay    = String(row[COL.DAY    - 1]).trim();
      // day가 전달된 경우: 같은 요일 예약만 충돌로 처리
      // → 같은 날짜라도 다른 요일(교수)의 예약은 간섭하지 않음
      if (rowDate === date && rowStatus !== 'cancelled' &&
          (!day || rowDay === day)) {
        taken.push(rowTime);
      }
    });
  }

  Logger.log('getAvailableSlots(' + date + ', ' + (day||'*') + ') → ' + JSON.stringify(taken));
  return { status: 'success', takenSlots: taken };
}

// ─────────────────────────────────────────────────────────────────────────────
//  예약 조회 (학번 + 이메일)
// ─────────────────────────────────────────────────────────────────────────────
function lookupReservation(studentId, email) {
  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: 'notFound' };

  var rows = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    var row      = rows[i];
    var rowSid   = String(row[COL.STUDENT_ID - 1]).trim();
    var rowEmail = String(row[COL.EMAIL      - 1]).trim().toLowerCase();
    var rowStat  = String(row[COL.STATUS     - 1]).trim();

    if (rowSid   === String(studentId).trim() &&
        rowEmail === String(email).trim().toLowerCase() &&
        rowStat  !== 'cancelled') {
      return {
        status:    'found',
        rowIndex:  i + 2,
        name:      row[COL.NAME      - 1],
        studentId: rowSid,
        contact:   row[COL.CONTACT   - 1],
        email:     row[COL.EMAIL     - 1],
        day:       row[COL.DAY       - 1],
        date:      normalizeDate(row[COL.DATE - 1]),
        time:      normalizeTime(row[COL.TIME - 1]),
        professor: row[COL.PROFESSOR - 1],
        location:  row[COL.LOCATION  - 1],
        fileId:    String(row[COL.FILE_ID   - 1]).trim()
      };
    }
  }
  return { status: 'notFound' };
}

// ─────────────────────────────────────────────────────────────────────────────
//  Drive 파일 삭제 (예약 취소 시)
//  fileIds: 쉼표로 구분된 Drive 파일 ID 문자열
// ─────────────────────────────────────────────────────────────────────────────
function deleteFiles(fileIds) {
  if (!fileIds) return;
  var ids = String(fileIds).split(',');
  ids.forEach(function(id) {
    id = id.trim();
    if (!id) return;
    try {
      var file = DriveApp.getFileById(id);
      file.setTrashed(true);
      Logger.log('파일 삭제(휴지통): ' + id);
    } catch (err) {
      Logger.log('파일 삭제 실패 (' + id + '): ' + err.message);
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────────
//  Drive 파일명 변경 (예약 날짜 변경 시)
//  fileIds: 쉼표로 구분된 Drive 파일 ID 문자열
//  oldDate / newDate: 'yyyy-MM-dd' 형식
// ─────────────────────────────────────────────────────────────────────────────
function renameFiles(fileIds, oldDate, newDate, name, studentId) {
  if (!fileIds || !newDate) return;
  var ids = String(fileIds).split(',');
  ids.forEach(function(id) {
    id = id.trim();
    if (!id) return;
    try {
      var file    = DriveApp.getFileById(id);
      var oldName = file.getName();
      // 기존 파일명: oldDate_name_studentId_원본파일명
      // oldDate 앞부분만 newDate로 교체
      var prefix  = oldDate + '_' + name + '_' + studentId + '_';
      var newPrefix = newDate + '_' + name + '_' + studentId + '_';
      var newName;
      if (oldName.indexOf(prefix) === 0) {
        newName = newPrefix + oldName.substring(prefix.length);
      } else {
        // 예상 패턴이 아니면 앞에 새 날짜 붙이기
        newName = newDate + '_' + oldName;
      }
      file.setName(newName);
      Logger.log('파일명 변경: ' + oldName + ' → ' + newName);
    } catch (err) {
      Logger.log('파일명 변경 실패 (' + id + '): ' + err.message);
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────────
//  공통 헬퍼
// ─────────────────────────────────────────────────────────────────────────────
function getRouting(day) {
  if (day === '화요일') return { professorName: '허재영 교수님', location: '광개토관 312호' };
  return { professorName: '김용관 대우교수님', location: '광개토관 3층 CDC실 팀플룸' };
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function normalizeDate(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'GMT+9', 'yyyy-MM-dd');
  return String(val).trim().slice(0, 10);
}

function normalizeTime(val) {
  if (val instanceof Date) {
    // 1899-12-30 기준 Date 객체: 역사적 LMT 오프셋 문제 방지를 위해 스프레드시트 타임존 사용
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  if (typeof val === 'number') {
    var total = Math.round(val * 24 * 60);
    return String(Math.floor(total / 60)).padStart(2, '0') + ':' + String(total % 60).padStart(2, '0');
  }
  return String(val).trim().slice(0, 5);
}

// ─────────────────────────────────────────────────────────────────────────────
//  스프레드시트 저장
// ─────────────────────────────────────────────────────────────────────────────
function saveToSheet(name, studentId, contact, email, day, date, time,
                     professorName, location, timestamp, status, eventId, fileId) {
  var sheet = getSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시간','이름','학번','연락처','이메일',
                     '희망요일','희망날짜','희망시간','담당교수','장소','상태','캘린더ID','파일ID']);
    sheet.getRange(1, 1, 1, COL.TOTAL)
         .setFontWeight('bold').setBackground('#28C158').setFontColor('#ffffff');
  }

  sheet.appendRow([
    Utilities.formatDate(timestamp, 'GMT+9', 'yyyy-MM-dd HH:mm:ss'),
    name, studentId, contact, email,
    day, date, time, professorName, location,
    status || 'active', eventId || '', fileId || ''
  ]);

  // 시간 셀을 텍스트로 고정 (Sheets 자동 변환 방지)
  var newRow = sheet.getLastRow();
  sheet.getRange(newRow, COL.TIME).setNumberFormat('@').setValue(time);
}

// ─────────────────────────────────────────────────────────────────────────────
//  date('yyyy-MM-dd') + time('HH:mm') 문자열 → Date 객체
//  (이메일 본문 날짜 표기용. 캘린더 연동 없이 시트에 있는 값만으로 생성)
// ─────────────────────────────────────────────────────────────────────────────
function buildDateTime(date, time) {
  var parts     = String(date).split('-');
  var timeParts = String(time).split(':');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]),
                  parseInt(timeParts[0] || 0), parseInt(timeParts[1] || 0), 0);
}

// ─────────────────────────────────────────────────────────────────────────────
//  이메일 템플릿
// ─────────────────────────────────────────────────────────────────────────────
function emailWrap(body) {
  return '<div style="font-family:\'Apple SD Gothic Neo\',sans-serif;max-width:600px;margin:0 auto;background:#fff;">' +
    '<div style="background:linear-gradient(135deg,#28C158,#34D068);padding:32px 36px;border-radius:16px 16px 0 0;">' +
      '<p style="color:rgba(255,255,255,.8);font-size:12px;letter-spacing:2px;margin:0 0 4px;">SEJONG UNIV. SEED</p>' +
      '<h1 style="color:#fff;font-size:22px;margin:0;font-weight:800;">비즈 아카데미</h1>' +
      '<p style="color:rgba(255,255,255,.85);font-size:13px;margin:6px 0 0;">1:1 교수 상담</p>' +
    '</div>' +
    '<div style="padding:32px 36px;">' + body + '</div>' +
    '<div style="background:#F9FAFB;padding:20px 36px;border-radius:0 0 16px 16px;text-align:center;' +
         'font-size:12px;color:#9CA3AF;border-top:1px solid #F3F4F6;">' +
      '세종대학교 경영경제대학 학생회 <b style="color:#28C158;">Seed</b><br>본 메일은 발신 전용입니다.' +
    '</div></div>';
}

function infoCard(professorName, formattedDate, day, time, location) {
  return '<div style="background:#F4FBF7;border:1.5px solid #28C158;border-radius:14px;padding:22px 24px;margin-bottom:24px;">' +
    '<p style="font-size:12px;font-weight:700;color:#28C158;letter-spacing:1.5px;margin:0 0 14px;">📋 상담 정보</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
      '<tr><td style="padding:5px 0;color:#777;width:80px;">담당교수</td>' +
          '<td style="padding:5px 0;font-weight:600;">' + professorName + '</td></tr>' +
      '<tr><td style="padding:5px 0;color:#777;">상담일시</td>' +
          '<td style="padding:5px 0;font-weight:600;">' + formattedDate + ' (' + day + ') ' + time + '</td></tr>' +
      '<tr><td style="padding:5px 0;color:#777;">상담장소</td>' +
          '<td style="padding:5px 0;font-weight:600;">' + location + '</td></tr>' +
    '</table></div>';
}

function surveyButton() {
  return '<div style="text-align:center;margin-bottom:28px;">' +
    '<p style="font-size:13px;color:#555;margin:0 0 12px;">상담 후 아래 링크를 통해 만족도 조사에 참여해 주세요.</p>' +
    '<a href="https://docs.google.com/forms/d/e/1FAIpQLSe_bVOkqjOvl7OWdGHg0ZsjFQe5eIPlI0A_RvyvWiRnopqz8Q/viewform?usp=dialog" ' +
       'style="display:inline-block;background:#28C158;color:#fff;font-weight:700;font-size:13px;' +
              'text-decoration:none;padding:12px 28px;border-radius:50px;">' +
      '👉 비즈 아카데미 만족도 조사 참여하기</a></div>';
}

// ─────────────────────────────────────────────────────────────────────────────
//  관리자 신규 예약 알림
// ─────────────────────────────────────────────────────────────────────────────
var ADMIN_EMAIL = 'sejong.bne.35th@gmail.com';

function sendAdminNotification(name, studentId, contact, email, day, date, time, professorName) {
  var parts         = String(date).split('-');
  var formattedDate = parts[0] + '년 ' + parseInt(parts[1]) + '월 ' + parseInt(parts[2]) + '일';
  var now           = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss');

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 6px;font-weight:800;">📬 신규 예약이 접수되었습니다</p>' +
    '<p style="font-size:13px;color:#6B7280;margin:0 0 24px;">접수 시각: ' + now + '</p>' +

    '<div style="background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:14px;' +
         'padding:22px 24px;margin-bottom:24px;">' +
      '<p style="font-size:12px;font-weight:700;color:#28C158;letter-spacing:1.5px;margin:0 0 14px;">👤 예약자 정보</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
        '<tr><td style="padding:6px 0;color:#777;width:90px;">성명</td>' +
            '<td style="padding:6px 0;font-weight:700;">' + name + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">학번</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + studentId + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">연락처</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + contact + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">이메일</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + email + '</td></tr>' +
      '</table>' +
    '</div>' +

    '<div style="background:#F4FBF7;border:1.5px solid #28C158;border-radius:14px;' +
         'padding:22px 24px;">' +
      '<p style="font-size:12px;font-weight:700;color:#28C158;letter-spacing:1.5px;margin:0 0 14px;">📋 상담 정보</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
        '<tr><td style="padding:6px 0;color:#777;width:90px;">담당교수</td>' +
            '<td style="padding:6px 0;font-weight:700;">' + professorName + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">날짜</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + formattedDate + ' (' + day + ')</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">시간</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + time + '</td></tr>' +
      '</table>' +
    '</div>'
  );

  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  '[비즈아카데미] 신규 예약 — ' + name + ' (' + formattedDate + ' ' + time + ')',
    htmlBody: body
  });
  Logger.log('관리자 알림 발송 → ' + ADMIN_EMAIL);
}

// ─────────────────────────────────────────────────────────────────────────────
//  관리자 예약 취소 알림
// ─────────────────────────────────────────────────────────────────────────────
function sendAdminCancelNotification(name, studentId, contact, day, date, time, professorName) {
  var parts         = String(date).split('-');
  var formattedDate = parts[0] + '년 ' + parseInt(parts[1]) + '월 ' + parseInt(parts[2]) + '일';
  var now           = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss');

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 6px;font-weight:800;">🚫 예약이 취소되었습니다</p>' +
    '<p style="font-size:13px;color:#6B7280;margin:0 0 24px;">취소 시각: ' + now + '</p>' +

    '<div style="background:#FEF2F2;border:1.5px solid #FCA5A5;border-radius:14px;' +
         'padding:22px 24px;margin-bottom:24px;">' +
      '<p style="font-size:12px;font-weight:700;color:#EF4444;letter-spacing:1.5px;margin:0 0 14px;">👤 취소 예약자 정보</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
        '<tr><td style="padding:6px 0;color:#777;width:90px;">성명</td>' +
            '<td style="padding:6px 0;font-weight:700;">' + name + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">학번</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + studentId + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">연락처</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + contact + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">담당교수</td>' +
            '<td style="padding:6px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + professorName + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">날짜</td>' +
            '<td style="padding:6px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + formattedDate + ' (' + day + ')</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">시간</td>' +
            '<td style="padding:6px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + time + '</td></tr>' +
      '</table>' +
    '</div>'
  );

  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  '[비즈아카데미] 예약 취소 — ' + name + ' (' + formattedDate + ' ' + time + ')',
    htmlBody: body
  });
  Logger.log('관리자 취소 알림 발송 → ' + ADMIN_EMAIL);
}

// ─────────────────────────────────────────────────────────────────────────────
//  관리자 예약 변경 알림
// ─────────────────────────────────────────────────────────────────────────────
function sendAdminModifyNotification(name, studentId, contact,
                                     oldDay, oldDate, oldTime, oldProfessor,
                                     newDay, newDate, newTime, newProfessor) {
  var oldParts = String(oldDate).split('-');
  var newParts = String(newDate).split('-');
  var oldFormatted = oldParts[0] + '년 ' + parseInt(oldParts[1]) + '월 ' + parseInt(oldParts[2]) + '일';
  var newFormatted = newParts[0] + '년 ' + parseInt(newParts[1]) + '월 ' + parseInt(newParts[2]) + '일';
  var now = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss');

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 6px;font-weight:800;">✏️ 예약이 변경되었습니다</p>' +
    '<p style="font-size:13px;color:#6B7280;margin:0 0 24px;">변경 시각: ' + now + '</p>' +

    '<div style="background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:14px;' +
         'padding:22px 24px;margin-bottom:16px;">' +
      '<p style="font-size:12px;font-weight:700;color:#28C158;letter-spacing:1.5px;margin:0 0 14px;">👤 예약자 정보</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
        '<tr><td style="padding:6px 0;color:#777;width:90px;">성명</td>' +
            '<td style="padding:6px 0;font-weight:700;">' + name + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">학번</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + studentId + '</td></tr>' +
        '<tr><td style="padding:6px 0;color:#777;">연락처</td>' +
            '<td style="padding:6px 0;font-weight:600;">' + contact + '</td></tr>' +
      '</table>' +
    '</div>' +

    '<div style="display:flex;gap:12px;margin-bottom:4px;">' +
      // 변경 전
      '<div style="flex:1;background:#FEF2F2;border:1.5px solid #FCA5A5;border-radius:14px;padding:18px 20px;">' +
        '<p style="font-size:11px;font-weight:700;color:#EF4444;letter-spacing:1px;margin:0 0 10px;">변경 전</p>' +
        '<table style="font-size:13px;color:#333;border-collapse:collapse;">' +
          '<tr><td style="padding:4px 0;color:#777;padding-right:10px;">교수</td>' +
              '<td style="padding:4px 0;text-decoration:line-through;color:#9CA3AF;">' + oldProfessor + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#777;">날짜</td>' +
              '<td style="padding:4px 0;text-decoration:line-through;color:#9CA3AF;">' + oldFormatted + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#777;">시간</td>' +
              '<td style="padding:4px 0;text-decoration:line-through;color:#9CA3AF;">' + oldTime + '</td></tr>' +
        '</table>' +
      '</div>' +
      // 변경 후
      '<div style="flex:1;background:#F4FBF7;border:1.5px solid #28C158;border-radius:14px;padding:18px 20px;">' +
        '<p style="font-size:11px;font-weight:700;color:#28C158;letter-spacing:1px;margin:0 0 10px;">변경 후</p>' +
        '<table style="font-size:13px;color:#333;border-collapse:collapse;">' +
          '<tr><td style="padding:4px 0;color:#777;padding-right:10px;">교수</td>' +
              '<td style="padding:4px 0;font-weight:700;">' + newProfessor + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#777;">날짜</td>' +
              '<td style="padding:4px 0;font-weight:600;">' + newFormatted + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#777;">시간</td>' +
              '<td style="padding:4px 0;font-weight:600;">' + newTime + '</td></tr>' +
        '</table>' +
      '</div>' +
    '</div>'
  );

  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  '[비즈아카데미] 예약 변경 — ' + name + ' (' + newFormatted + ' ' + newTime + ')',
    htmlBody: body
  });
  Logger.log('관리자 변경 알림 발송 → ' + ADMIN_EMAIL);
}

// ─────────────────────────────────────────────────────────────────────────────
//  확인 이메일 (사전 프로필 제출 결과 포함)
// ─────────────────────────────────────────────────────────────────────────────
function sendConfirmationEmail(name, email, day, date, time,
                               professorName, location, eventDate,
                               fileId, fileName) {
  var formattedDate = Utilities.formatDate(eventDate, 'GMT+9', 'yyyy년 MM월 dd일');

  // 파일 제출 상태 섹션
  var fileSection;
  if (fileId) {
    var fileLink = 'https://drive.google.com/file/d/' + fileId + '/view';
    fileSection =
      '<div style="background:#F0FDF4;border:1.5px solid #86EFAC;border-radius:14px;padding:18px 22px;margin-bottom:24px;">' +
        '<p style="font-size:12px;font-weight:700;color:#16A34A;letter-spacing:1.2px;margin:0 0 10px;">✅ 사전 프로필 제출 완료</p>' +
        '<p style="font-size:13px;color:#374151;margin:0 0 12px;">제출하신 파일이 정상적으로 접수되었습니다.</p>' +
        '<table style="width:100%;font-size:13px;color:#374151;">' +
          '<tr><td style="color:#777;width:60px;padding:3px 0;">파일명</td>' +
              '<td style="padding:3px 0;font-weight:600;">' + (fileName || '파일') + '</td></tr>' +
        '</table>' +
        '<p style="margin:12px 0 0;font-size:12px;color:#6B7280;">제출된 파일은 담당 교수님께 전달됩니다.</p>' +
      '</div>';
  } else {
    fileSection =
      '<div style="background:#FFFBEB;border:1.5px solid #FCD34D;border-radius:14px;padding:18px 22px;margin-bottom:24px;">' +
        '<p style="font-size:12px;font-weight:700;color:#92400E;letter-spacing:1.2px;margin:0 0 10px;">⚠️ 사전 프로필 미제출</p>' +
        '<p style="font-size:13px;color:#374151;margin:0;">파일 업로드에 문제가 발생했습니다.<br>' +
        '이 이메일에 사전 프로필 파일을 첨부하여 회신해 주세요.</p>' +
      '</div>';
  }

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + name + '</b>님, 안녕하세요!</p>' +
    '<p style="font-size:14px;color:#555;margin:0 0 28px;line-height:1.7;">' +
      '비즈 아카데미 교수 상담 예약이 완료되었습니다.<br>아래 상담 정보를 확인해 주세요.</p>' +
    infoCard(professorName, formattedDate, day, time, location) +
    fileSection +
    '<div style="background:#FFFBEB;border:1.5px solid #FCD34D;border-radius:14px;' +
         'padding:18px 22px;margin-bottom:24px;font-size:13px;color:#92400E;line-height:1.8;">' +
      '<b>⚠️ 유의사항</b><br>' +
      '• 추가 작성 자료가 있는 경우 이메일로 함께 제출해 주시기 바랍니다.<br>' +
      '• <b>상담 불참 시 추후 프로그램 참여에 제한이 있을 수 있습니다.</b>' +
    '</div>' +
    surveyButton()
  );

  MailApp.sendEmail({ to: email, subject: '[비즈아카데미] 상담 예약이 완료되었습니다', htmlBody: body });
  Logger.log('확인 이메일 발송 → ' + email);
}

// ─────────────────────────────────────────────────────────────────────────────
//  취소 이메일
// ─────────────────────────────────────────────────────────────────────────────
function sendCancellationEmail(name, email, day, date, time, professorName, location) {
  var parts = String(date).split('-');
  var formattedDate = parts[0] + '년 ' + parseInt(parts[1]) + '월 ' + parseInt(parts[2]) + '일';

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + name + '</b>님, 안녕하세요!</p>' +
    '<p style="font-size:14px;color:#555;margin:0 0 28px;line-height:1.7;">' +
      '아래 상담 예약이 <b style="color:#EF4444;">취소</b>되었습니다.</p>' +
    '<div style="background:#FEF2F2;border:1.5px solid #FCA5A5;border-radius:14px;padding:22px 24px;margin-bottom:24px;">' +
      '<p style="font-size:12px;font-weight:700;color:#EF4444;letter-spacing:1.5px;margin:0 0 14px;">🚫 취소된 예약</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;color:#333;">' +
        '<tr><td style="padding:5px 0;color:#777;width:80px;">담당교수</td>' +
            '<td style="padding:5px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + professorName + '</td></tr>' +
        '<tr><td style="padding:5px 0;color:#777;">상담일시</td>' +
            '<td style="padding:5px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + formattedDate + ' (' + day + ') ' + time + '</td></tr>' +
        '<tr><td style="padding:5px 0;color:#777;">상담장소</td>' +
            '<td style="padding:5px 0;font-weight:600;text-decoration:line-through;color:#9CA3AF;">' + location + '</td></tr>' +
      '</table>' +
    '</div>'
  );

  MailApp.sendEmail({ to: email, subject: '[비즈아카데미] 상담 예약이 취소되었습니다', htmlBody: body });
  Logger.log('취소 이메일 발송 → ' + email);
}

// ─────────────────────────────────────────────────────────────────────────────
//  변경 이메일
// ─────────────────────────────────────────────────────────────────────────────
function sendModificationEmail(name, email, day, date, time, professorName, location, eventDate) {
  var formattedDate = Utilities.formatDate(eventDate, 'GMT+9', 'yyyy년 MM월 dd일');

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + name + '</b>님, 안녕하세요!</p>' +
    '<p style="font-size:14px;color:#555;margin:0 0 28px;line-height:1.7;">' +
      '상담 예약이 아래와 같이 <b style="color:#28C158;">변경</b>되었습니다.</p>' +
    infoCard(professorName, formattedDate, day, time, location) +
    '<div style="background:#FFFBEB;border:1.5px solid #FCD34D;border-radius:14px;' +
         'padding:18px 22px;margin-bottom:24px;font-size:13px;color:#92400E;line-height:1.8;">' +
      '<b>⚠️ 유의사항</b><br>' +
      '• 사전 프로필을 아직 제출하지 않으셨다면 빠른 시일 내에 제출해 주세요.<br>' +
      '• <b>상담 불참 시 추후 프로그램 참여에 제한이 있을 수 있습니다.</b>' +
    '</div>' +
    surveyButton()
  );

  MailApp.sendEmail({ to: email, subject: '[비즈아카데미] 상담 예약이 변경되었습니다', htmlBody: body });
  Logger.log('변경 이메일 발송 → ' + email);
}

// ─────────────────────────────────────────────────────────────────────────────
//  만료 파일 자동 정리 — 트리거로 매일 실행
//  조건: 예약 상태가 active이고 멘토링 날짜로부터 7일이 경과한 경우
// ─────────────────────────────────────────────────────────────────────────────
function cleanupExpiredFiles() {
  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('cleanupExpiredFiles: 데이터 없음'); return; }

  // 오늘 기준 7일 전 날짜 (KST)
  var now    = new Date();
  var cutoff = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  Logger.log('cleanupExpiredFiles 실행 | 기준일: ' +
    Utilities.formatDate(cutoff, 'GMT+9', 'yyyy-MM-dd'));

  var rows    = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues();
  var cleaned = 0;

  rows.forEach(function(row, i) {
    var fileId = String(row[COL.FILE_ID - 1]).trim();
    if (!fileId) return; // 파일 없음 or 이미 정리됨

    var status   = String(row[COL.STATUS - 1]).trim();
    if (status === 'cancelled') return; // 취소 건은 cancelReservation에서 이미 삭제

    var rowDate  = normalizeDate(row[COL.DATE - 1]);
    var parts    = rowDate.split('-');
    if (parts.length !== 3) return;

    var mentoringDate = new Date(
      parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]),
      23, 59, 59 // 당일 끝까지 보호
    );

    if (mentoringDate < cutoff) {
      deleteFiles(fileId);
      // FILE_ID 셀을 비워 중복 삭제 방지
      sheet.getRange(i + 2, COL.FILE_ID).setValue('');
      cleaned++;
      Logger.log('  ▶ 만료 삭제 | 행 ' + (i + 2) + ' | 날짜: ' + rowDate +
                 ' | 이름: ' + row[COL.NAME - 1]);
    }
  });

  Logger.log('cleanupExpiredFiles 완료: ' + cleaned + '건 삭제');
}

// ─────────────────────────────────────────────────────────────────────────────
//  자동 트리거 등록 (최초 1회만 Apps Script 에디터에서 수동 실행)
//  이미 트리거가 있으면 중복 생성하지 않음
// ─────────────────────────────────────────────────────────────────────────────
function createCleanupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupExpiredFiles') {
      Logger.log('트리거 이미 존재 — 중복 생성 건너뜀');
      return;
    }
  }
  ScriptApp.newTrigger('cleanupExpiredFiles')
    .timeBased()
    .everyDays(1)
    .atHour(3)   // 매일 오전 3시 실행
    .create();
  Logger.log('cleanupExpiredFiles 트리거 생성 완료 (매일 오전 3시)');
}

// ─────────────────────────────────────────────────────────────────────────────
//  멘토링 전날 리마인더 발송 — 트리거로 매일 오전 9시 실행
//  시트 데이터를 직접 읽어 내일 예약된 활성 건을 대상으로 이메일 발송
// ─────────────────────────────────────────────────────────────────────────────
function sendDayBeforeReminders() {
  var tz      = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var now     = new Date();
  var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  var tomorrowStr = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');

  Logger.log('=== sendDayBeforeReminders 시작 | 내일: ' + tomorrowStr + ' ===');

  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('데이터 없음'); return; }

  var rows = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues();
  var sent = 0;

  rows.forEach(function(row, i) {
    var status  = String(row[COL.STATUS - 1]).trim();
    if (status === 'cancelled') return;

    var rowDate = normalizeDate(row[COL.DATE - 1]);
    if (rowDate !== tomorrowStr) return;

    var name      = String(row[COL.NAME      - 1]).trim();
    var email     = String(row[COL.EMAIL     - 1]).trim();
    var day       = String(row[COL.DAY       - 1]).trim();
    var time      = normalizeTime(row[COL.TIME - 1]);
    var professor = String(row[COL.PROFESSOR - 1]).trim();
    var location  = String(row[COL.LOCATION  - 1]).trim();

    if (!email) { Logger.log('이메일 없음 — 행 ' + (i + 2)); return; }

    try {
      sendReminderEmail(name, email, day, rowDate, time, professor, location);
      sent++;
      Logger.log('  ▶ 발송 성공: ' + name + ' <' + email + '> | ' + rowDate + ' ' + time);
    } catch (err) {
      Logger.log('  X 발송 실패: ' + name + ' | ' + err.message);
    }
  });

  Logger.log('=== sendDayBeforeReminders 완료: ' + sent + '건 발송 ===');
}

// ─────────────────────────────────────────────────────────────────────────────
//  리마인더 이메일 본문
// ─────────────────────────────────────────────────────────────────────────────
function sendReminderEmail(name, email, day, date, time, professorName, location) {
  var parts         = String(date).split('-');
  var formattedDate = parts[0] + '년 ' + parseInt(parts[1]) + '월 ' + parseInt(parts[2]) + '일';

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + name + '</b>님, 안녕하세요!</p>' +
    '<p style="font-size:14px;color:#555;margin:0 0 6px;line-height:1.7;">' +
      '내일 비즈 아카데미 교수 상담이 예정되어 있습니다.<br>' +
      '아래 일정을 다시 한 번 확인해 주세요 😊</p>' +

    // 강조 배너
    '<div style="background:linear-gradient(135deg,#28C158,#34D068);border-radius:14px;' +
         'padding:16px 22px;margin-bottom:24px;display:flex;align-items:center;gap:14px;">' +
      '<div style="font-size:28px;line-height:1;">📅</div>' +
      '<div>' +
        '<p style="color:rgba(255,255,255,.8);font-size:11px;margin:0 0 3px;letter-spacing:1px;">D-1 · 내일 일정</p>' +
        '<p style="color:#fff;font-size:15px;font-weight:800;margin:0;">' +
          formattedDate + ' (' + day.replace('요일','') + ') ' + time +
        '</p>' +
      '</div>' +
    '</div>' +

    infoCard(professorName, formattedDate, day, time, location) +

    // 체크리스트
    '<div style="background:#F0FDF4;border:1.5px solid #86EFAC;border-radius:14px;' +
         'padding:18px 22px;margin-bottom:24px;">' +
      '<p style="font-size:12px;font-weight:700;color:#16A34A;letter-spacing:1.2px;margin:0 0 12px;">✅ 상담 전 체크리스트</p>' +
      '<div style="font-size:13px;color:#374151;line-height:2;">' +
        '<p style="margin:0;">☐ &nbsp;사전 프로필 파일 제출 여부 확인</p>' +
        '<p style="margin:0;">☐ &nbsp;상담 장소 사전 확인 (' + location + ')</p>' +
        '<p style="margin:0;">☐ &nbsp;질문 사항 미리 정리하기</p>' +
      '</div>' +
    '</div>' +

    // 주의사항
    '<div style="background:#FEF2F2;border:1.5px solid #FCA5A5;border-radius:14px;' +
         'padding:16px 22px;margin-bottom:24px;font-size:13px;color:#991B1B;line-height:1.8;">' +
      '<b>⚠️ 불참 안내</b><br>' +
      '상담에 불참하실 경우 추후 프로그램 참여에 제한이 있을 수 있습니다.<br>' +
      '부득이한 사정이 있을 경우 <b>반드시 사전에 취소</b>해 주시기 바랍니다.' +
    '</div>' +

    surveyButton()
  );

  MailApp.sendEmail({
    to:       email,
    subject:  '[비즈아카데미] 내일 교수 상담이 예정되어 있습니다 — ' + formattedDate,
    htmlBody: body
  });
}

// ─────────────────────────────────────────────────────────────────────────────
//  리마인더 트리거 등록 (최초 1회만 Apps Script 에디터에서 수동 실행)
// ─────────────────────────────────────────────────────────────────────────────
function createReminderTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDayBeforeReminders') {
      Logger.log('리마인더 트리거 이미 존재 — 중복 생성 건너뜀');
      return;
    }
  }
  ScriptApp.newTrigger('sendDayBeforeReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)   // 매일 오전 9시 실행
    .create();
  Logger.log('sendDayBeforeReminders 트리거 생성 완료 (매일 오전 9시)');
}

// ─────────────────────────────────────────────────────────────────────────────
//  주간 멘토링 일정 요약 — 매주 월요일 오전 9시 자동 발송
// ─────────────────────────────────────────────────────────────────────────────
function sendWeeklySummary() {
  var tz    = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var today = new Date();
  var dow   = today.getDay(); // 0=일, 1=월 ...

  // 이번 주 월요일 기준 날짜 계산
  var monday = new Date(today);
  monday.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1));
  monday.setHours(0, 0, 0, 0);

  // 이번 주 화요일(+1), 수요일(+2)
  var tuesday   = new Date(monday); tuesday.setDate(monday.getDate() + 1);
  var wednesday = new Date(monday); wednesday.setDate(monday.getDate() + 2);
  var tuesdayStr   = Utilities.formatDate(tuesday,   tz, 'yyyy-MM-dd');
  var wednesdayStr = Utilities.formatDate(wednesday, tz, 'yyyy-MM-dd');

  Logger.log('=== sendWeeklySummary | 화: ' + tuesdayStr + ' | 수: ' + wednesdayStr + ' ===');

  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  var tueSessions = [], wedSessions = [];

  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues();
    rows.forEach(function(row) {
      var status    = String(row[COL.STATUS     - 1]).trim();
      if (status === 'cancelled') return;

      var rowDate   = normalizeDate(row[COL.DATE      - 1]);
      var rowTime   = normalizeTime(row[COL.TIME      - 1]);
      var name      = String(row[COL.NAME       - 1]).trim();
      var studentId = String(row[COL.STUDENT_ID - 1]).trim();
      var professor = String(row[COL.PROFESSOR  - 1]).trim();
      var contact   = String(row[COL.CONTACT    - 1]).trim();

      var entry = { name: name, studentId: studentId, contact: contact,
                    time: rowTime, professor: professor };

      if      (rowDate === tuesdayStr)   tueSessions.push(entry);
      else if (rowDate === wednesdayStr) wedSessions.push(entry);
    });
  }

  // 시간 오름차순 정렬
  tueSessions.sort(function(a, b) { return a.time.localeCompare(b.time); });
  wedSessions.sort(function(a, b) { return a.time.localeCompare(b.time); });

  var total = tueSessions.length + wedSessions.length;
  Logger.log('이번 주 예약: 화요일 ' + tueSessions.length + '건 / 수요일 ' + wedSessions.length + '건');

  sendWeeklySummaryEmail(tuesdayStr, wednesdayStr, tueSessions, wedSessions, total);
}

function sendWeeklySummaryEmail(tuesdayStr, wednesdayStr, tueSessions, wedSessions, total) {
  var tParts = tuesdayStr.split('-');
  var wParts = wednesdayStr.split('-');
  var tLabel = parseInt(tParts[1]) + '월 ' + parseInt(tParts[2]) + '일 (화)';
  var wLabel = parseInt(wParts[1]) + '월 ' + parseInt(wParts[2]) + '일 (수)';
  var now    = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss');

  function sessionTable(sessions) {
    if (sessions.length === 0) {
      return '<p style="font-size:13px;color:#9CA3AF;margin:8px 0 0;">이번 주 예약 없음</p>';
    }
    var rows = sessions.map(function(s) {
      return '<tr>' +
        '<td style="padding:8px 10px;font-weight:700;font-size:14px;">' + s.time + '</td>' +
        '<td style="padding:8px 10px;font-size:14px;">' + s.name + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;color:#6B7280;">' + s.studentId + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;color:#6B7280;">' + s.contact + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;">' + s.professor + '</td>' +
      '</tr>';
    }).join('');
    return '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
      '<thead><tr style="background:#F3F4F6;">' +
        '<th style="padding:8px 10px;text-align:left;font-size:11px;color:#6B7280;font-weight:600;">시간</th>' +
        '<th style="padding:8px 10px;text-align:left;font-size:11px;color:#6B7280;font-weight:600;">성명</th>' +
        '<th style="padding:8px 10px;text-align:left;font-size:11px;color:#6B7280;font-weight:600;">학번</th>' +
        '<th style="padding:8px 10px;text-align:left;font-size:11px;color:#6B7280;font-weight:600;">연락처</th>' +
        '<th style="padding:8px 10px;text-align:left;font-size:11px;color:#6B7280;font-weight:600;">담당교수</th>' +
      '</tr></thead>' +
      '<tbody>' + rows + '</tbody>' +
    '</table>';
  }

  var body = emailWrap(
    '<p style="font-size:15px;color:#111;margin:0 0 4px;font-weight:800;">📅 이번 주 멘토링 일정 요약</p>' +
    '<p style="font-size:13px;color:#6B7280;margin:0 0 24px;">발송 시각: ' + now +
      ' &nbsp;·&nbsp; 총 <b style="color:#28C158;">' + total + '명</b></p>' +

    // 화요일 섹션
    '<div style="margin-bottom:20px;">' +
      '<div style="background:linear-gradient(135deg,#28C158,#34D068);border-radius:10px 10px 0 0;' +
           'padding:12px 16px;display:flex;align-items:center;justify-content:space-between;">' +
        '<span style="color:#fff;font-weight:700;font-size:14px;">🗓 ' + tLabel + '</span>' +
        '<span style="background:rgba(255,255,255,.25);color:#fff;font-size:12px;font-weight:700;' +
              'padding:2px 10px;border-radius:20px;">' + tueSessions.length + '명</span>' +
      '</div>' +
      '<div style="border:1.5px solid #28C158;border-top:none;border-radius:0 0 10px 10px;' +
           'padding:14px 16px;background:#fff;">' +
        sessionTable(tueSessions) +
      '</div>' +
    '</div>' +

    // 수요일 섹션
    '<div style="margin-bottom:4px;">' +
      '<div style="background:linear-gradient(135deg,#3B82F6,#60A5FA);border-radius:10px 10px 0 0;' +
           'padding:12px 16px;display:flex;align-items:center;justify-content:space-between;">' +
        '<span style="color:#fff;font-weight:700;font-size:14px;">🗓 ' + wLabel + '</span>' +
        '<span style="background:rgba(255,255,255,.25);color:#fff;font-size:12px;font-weight:700;' +
              'padding:2px 10px;border-radius:20px;">' + wedSessions.length + '명</span>' +
      '</div>' +
      '<div style="border:1.5px solid #3B82F6;border-top:none;border-radius:0 0 10px 10px;' +
           'padding:14px 16px;background:#fff;">' +
        sessionTable(wedSessions) +
      '</div>' +
    '</div>'
  );

  var weekLabel = tParts[1] + '월 ' + tParts[2] + '일~' + wParts[2] + '일';
  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  '[비즈아카데미] 이번 주 멘토링 일정 (' + weekLabel + ') — 총 ' + total + '명',
    htmlBody: body
  });
  Logger.log('주간 요약 발송 → ' + ADMIN_EMAIL);
}

// ─────────────────────────────────────────────────────────────────────────────
//  주간 요약 트리거 등록 (최초 1회만 Apps Script 에디터에서 수동 실행)
// ─────────────────────────────────────────────────────────────────────────────
function createWeeklySummaryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendWeeklySummary') {
      Logger.log('주간 요약 트리거 이미 존재 — 중복 생성 건너뜀');
      return;
    }
  }
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)   // 매주 월요일 오전 9시
    .create();
  Logger.log('sendWeeklySummary 트리거 생성 완료 (매주 월요일 오전 9시)');
}

// ─────────────────────────────────────────────────────────────────────────────
//  진단용 (Apps Script 에디터에서 직접 실행)
// ─────────────────────────────────────────────────────────────────────────────
function debugSheet() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  Logger.log('총 데이터 행: ' + (lastRow - 1));
  if (lastRow <= 1) { Logger.log('데이터 없음'); return; }
  sheet.getRange(2, 1, lastRow - 1, COL.TOTAL).getValues().forEach(function(row, i) {
    Logger.log('행' + (i+2) + ' | 날짜=' + normalizeDate(row[COL.DATE-1]) +
               ' | 시간=' + normalizeTime(row[COL.TIME-1]) +
               ' | 상태=' + row[COL.STATUS-1] +
               ' | 파일=' + row[COL.FILE_ID-1]);
  });
}

function testGetSlots() {
  var testDate = '2026-05-20';
  Logger.log('testGetSlots(' + testDate + ') → ' + JSON.stringify(getAvailableSlots(testDate)));
}
