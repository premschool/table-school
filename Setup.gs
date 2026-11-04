var SETUP_CONFIG = {
  SCHOOL_NAME: 'โรงเรียนเปรมติณสูลานนท์',
  ACADEMIC_YEAR: '2569',
  SEMESTER: '2',
  ADMIN_PASSWORD: '40140', // รหัสเริ่มต้น admin; เปลี่ยนทันทีหลัง Login ครั้งแรก
  TELEGRAM_BOT_TOKEN: '',    // ← ใส่ Bot Token (ถ้ามี)
  ADMIN_TELEGRAM_ID: ''      // ← ใส่ Chat ID (ถ้ามี)
};

/**
 * ★ MAIN SETUP — รันครั้งแรก ★
 */
function setupSystem() {
  var spreadsheetId = '1lXCgkCg1XAZAdgGntnlcRZdfKVb1LLFrOQdwalCyNoA';
  var ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (err) {
    throw new Error('เปิด Google Sheet ไม่ได้ กรุณาตรวจสอบ ID และสิทธิ์เข้าถึงของบัญชี Apps Script: ' + err.message);
  }
  if (!ss) throw new Error('ไม่พบ Spreadsheet ID: ' + spreadsheetId);

  var ssId = ss.getId();
  Logger.log('🔧 เริ่มตั้งค่าระบบ ตารางเรียน & จัดสอนแทน...');
  Logger.log('📊 Spreadsheet ID: ' + ssId);

  // 1. สร้าง Sheets ทั้ง 8 ชีต
  _createSheets(ss);
  Logger.log('✅ สร้าง Sheets เรียบร้อย');

  // 2. ใส่ Default Settings
  _insertDefaultSettings(ss);
  Logger.log('✅ ใส่ค่า Settings เริ่มต้นเรียบร้อย');

  // 3. ใส่ Default Periods (คาบเรียน)
  _insertDefaultPeriods(ss);
  Logger.log('✅ ใส่คาบเรียนเริ่มต้นเรียบร้อย');

  // 4. ใส่ Default Admin
  _insertDefaultAdmin(ss);
  Logger.log('✅ สร้าง Admin User เรียบร้อย');

  // 5. สร้าง Drive Folder สำหรับไฟล์แนบ
  var folderId = _createDriveFolder();
  Logger.log('✅ สร้าง Drive Folder: ' + folderId);

  // 6. บันทึก Spreadsheet ID + Folder ID ลง Script Properties
  PropertiesService.getScriptProperties().setProperties({
    'SPREADSHEET_ID': ssId,
    'DRIVE_FOLDER_ID': folderId
  });
  Logger.log('✅ บันทึก Script Properties เรียบร้อย');

  // 7. สร้าง Triggers
  _setupTriggers();
  Logger.log('✅ สร้าง Triggers เรียบร้อย');

  // 8. ลบ Sheet1 เริ่มต้น (ถ้ามี)
  _removeDefaultSheet(ss);

  Logger.log('');
  Logger.log('🎉 ===== ตั้งค่าระบบ PTNTIME-SCHOOL สำเร็จ! =====');
  Logger.log('📌 Spreadsheet: ' + ss.getUrl());
  Logger.log('📌 Web App URL: ' + ScriptApp.getService().getUrl());
  return '✅ ตั้งค่าระบบสำเร็จ! เข้าสู่ระบบด้วย admin / 40140 แล้วเปลี่ยนรหัสผ่านทันที';
}

// ── สร้าง Sheets ──
function _createSheets(ss) {
  var sheets = [
    {
      name: 'Teachers',
      headers: ['teacher_id','user_id','password','prefix','first_name','last_name','department',
                'teaching_room','subjects','grade_levels','phone','email','telegram_chat_id',
                'max_substitute_periods','status','role']
    },
    {
      name: 'Subjects',
      headers: ['subject_id','subject_code','subject_name','type','grade_levels',
                'class_names','total_periods','room']
    },
    {
      name: 'Schedules',
      headers: ['schedule_id','day_of_week','period','subject_id','class_name',
                'teacher_id','room']
    },
    {
      name: 'LeaveRequests',
      headers: ['leave_id','teacher_id','leave_date','leave_type','periods','reason',
                'attachment_file_id','attachment_url','attachment_name',
                'status','approved_by','approved_at','created_at']
    },
    {
      name: 'SubstituteAssignments',
      headers: ['assignment_id','leave_id','original_teacher_id','substitute_teacher_id',
                'schedule_date','period','class_name','subject','status',
                'notification_sent','assigned_by','assigned_at']
    },
    {
      name: 'Periods',
      headers: ['periodNumber','periodName','periodTime','periodType','duration']
    },
    {
      name: 'Settings',
      headers: ['key','value']
    },
    {
      name: 'Logs',
      headers: ['timestamp','action','user','details']
    },
    {
      name: 'Classrooms',
      headers: ['classroom_id','grade_level','room_number','class_name','homeroom_teacher_id','student_count']
    }
  ];

  for (var i = 0; i < sheets.length; i++) {
    var cfg = sheets[i];
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }
    // ใส่ Headers (ถ้าแถวแรกว่าง)
    var firstRow = sheet.getRange(1, 1, 1, cfg.headers.length).getValues()[0];
    var isEmpty = firstRow.every(function(v) { return !v; });
    if (isEmpty) {
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      // จัดรูปแบบ Header
      sheet.getRange(1, 1, 1, cfg.headers.length)
        .setBackground('#0a1628')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(10);
      sheet.setFrozenRows(1);
    }
  }
}

// ── Default Settings ──
function _insertDefaultSettings(ss) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var existingKeys = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) existingKeys[data[i][0]] = true;
  }

  var defaults = [
    ['school_name', SETUP_CONFIG.SCHOOL_NAME],
    ['school_logo_url', ''],
    ['academic_year', SETUP_CONFIG.ACADEMIC_YEAR],
    ['semester', SETUP_CONFIG.SEMESTER],
    ['telegram_bot_token', SETUP_CONFIG.TELEGRAM_BOT_TOKEN],
    ['admin_telegram_id', SETUP_CONFIG.ADMIN_TELEGRAM_ID],
    ['drive_folder_id', ''],
    ['schedule_days', 'จันทร์,อังคาร,พุธ,พฤหัสบดี,ศุกร์'],
    ['director_name', ''],
    ['academic_head_name', ''],
    ['print_header_text', SETUP_CONFIG.SCHOOL_NAME]
  ];

  var toInsert = [];
  for (var j = 0; j < defaults.length; j++) {
    if (!existingKeys[defaults[j][0]]) {
      toInsert.push(defaults[j]);
    }
  }
  if (toInsert.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toInsert.length, 2).setValues(toInsert);
  }
}

// ── Default Periods (คาบเรียน) ──
function _insertDefaultPeriods(ss) {
  var sheet = ss.getSheetByName('Periods');
  if (!sheet || sheet.getLastRow() > 1) return; // มีข้อมูลแล้ว skip

  var periods = [
    [1, 'คาบที่ 1', '08:30-09:20', 'เรียน', 50],
    [2, 'คาบที่ 2', '09:20-10:10', 'เรียน', 50],
    [3, 'คาบที่ 3', '10:10-11:00', 'เรียน', 50],
    [4, 'คาบที่ 4', '11:00-11:50', 'เรียน', 50],
    [5, 'พักเที่ยง', '11:50-12:40', 'พักเที่ยง', 50],
    [6, 'คาบที่ 5', '12:40-13:30', 'เรียน', 50],
    [7, 'คาบที่ 6', '13:30-14:20', 'เรียน', 50],
    [8, 'คาบที่ 7', '14:20-15:10', 'เรียน', 50],
    [9, 'คาบที่ 8', '15:10-16:00', 'เรียน', 50],
    [10, 'ซ่อมเสริม/PLC', '16:00-16:50', 'ซ่อมเสริม', 50]
  ];
  sheet.getRange(2, 1, periods.length, 5).setValues(periods);
}

// ── Default Admin ──
function _insertDefaultAdmin(ss) {
  var sheet = ss.getSheetByName('Teachers');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  var initializedKey = 'ADMIN_DEFAULT_PASSWORD_SET';
  // ตั้งรหัสเริ่มต้นเพียงครั้งเดียว เพื่อไม่รีเซ็ตรหัสทุกครั้งที่รัน setupSystem()
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'admin') {
      if (props.getProperty(initializedKey) !== 'true') {
        var headerMap = _buildHeaderIndex(data[0]);
        if (headerMap.password === undefined) throw new Error('ไม่พบคอลัมน์ password ในชีต Teachers');
        sheet.getRange(i + 1, headerMap.password + 1).setValue(_hashPw(SETUP_CONFIG.ADMIN_PASSWORD));
        props.setProperty(initializedKey, 'true');
      }
      return;
    }
  }

  var hashedPw = _hashPw(SETUP_CONFIG.ADMIN_PASSWORD);

  sheet.appendRow([
    'admin',           // teacher_id
    'admin',           // user_id
    hashedPw,          // password hash only
    '',                // prefix
    'ผู้ดูแลระบบ',      // first_name
    '',                // last_name
    'ฝ่ายวิชาการ',     // department
    '',                // teaching_room
    '[]',              // subjects
    '[]',              // grade_levels
    '',                // phone
    '',                // email
    '',                // telegram_chat_id
    0,                 // max_substitute_periods
    'active',          // status
    'Admin'            // role
  ]);
  props.setProperty(initializedKey, 'true');
}

// ── Drive Folder ──
function _createDriveFolder() {
  var folderName = 'PTNTIME-SCHOOL_LeaveAttachments';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  }
  var folder = DriveApp.createFolder(folderName);
  return folder.getId();
}

// ── Triggers ──
function _setupTriggers() {
  // ลบ triggers เก่าที่เกี่ยวข้อง
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'onSheetEdit' || fn === 'cleanupExpiredSessions') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // สร้าง onEdit trigger
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  // สร้าง cleanup trigger (ทุก 6 ชม.)
  ScriptApp.newTrigger('cleanupExpiredSessions')
    .timeBased()
    .everyHours(6)
    .create();
}

// ── ลบ Sheet1 เริ่มต้น ──
function _removeDefaultSheet(ss) {
  try {
    var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
  } catch (e) {}
}

