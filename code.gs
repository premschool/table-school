function _triggerDriveScope() { DriveApp.getRootFolder(); }

function doGet(e) {
  var tpl = HtmlService.createTemplateFromFile('index');
  // Template Injection — ลด API calls ตอน page load
  var userId = '';
  var sessionToken = '';
  try {
    // ถ้ามี session ค้างอยู่ ให้ inject initialData
    tpl.initialData = JSON.stringify({ ready: true });
  } catch (err) {
    tpl.initialData = JSON.stringify({ ready: false });
  }
  return tpl.evaluate()
    .setTitle('ระบบตารางเรียน & จัดสอนแทน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(f) {
  return HtmlService.createHtmlOutputFromFile(f).getContent();
}

// ==================== AUTH ====================
function verifyUserLogin(userId, password) {
  if (!userId || !password) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
  userId = String(userId).trim();
  if (_loginIsThrottled_(userId)) return { success: false, message: 'ใส่ข้อมูลผิดครบ 5 ครั้งแล้ว กรุณารอ 15 นาทีแล้วลองใหม่' };
  var user = DataService.getTeacherCredential(userId);
  if (!user) {
    _recordLoginFailure_(userId);
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
  if (user.status !== 'active') return { success: false, message: 'บัญชีถูกระงับการใช้งาน' };
  // บัญชีที่ยังไม่ได้ตั้งรหัสผ่านจะเข้าไม่ได้จนกว่าแอดมินจะกำหนดรหัสให้
  if (!user.password || !_verifyPassword(password, user.password)) {
    _recordLoginFailure_(userId);
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
  _clearLoginFailures_(userId);
  // อัปเกรด credential แบบเก่าเป็น PBKDF2 หลัง Login สำเร็จ
  if (String(user.password).indexOf('pbkdf2-sha256$') !== 0) {
    DataService.upgradeLegacyCredential(user.teacher_id, _hashPw(password));
  }

  // สร้าง session
  var token = Utilities.getUuid();
  var sessionData = JSON.stringify({
    teacherId: user.teacher_id,
    userId: user.user_id,
    name: (user.prefix || '') + user.first_name + ' ' + user.last_name,
    role: user.role,
    expiry: new Date().getTime() + SESSION_TTL
  });
  PropertiesService.getScriptProperties().setProperty('sess_' + token, sessionData);
  CacheService.getScriptCache().put('sess_' + token, sessionData, Math.floor(SESSION_TTL / 1000));

  logAction('LOGIN', 'User logged in', user.teacher_id);

  return {
    success: true,
    token: token,
    teacherId: user.teacher_id,
    userId: user.user_id,
    name: (user.prefix || '') + user.first_name + ' ' + user.last_name,
    role: user.role,
    department: user.department
  };
}

function checkSession(token) {
  if (!token) return { valid: false };
  var cache = CacheService.getScriptCache();
  var raw = cache.get('sess_' + token);
  if (!raw) {
    raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
    if (raw) cache.put('sess_' + token, raw, 3600);
  }
  if (!raw) return { valid: false };
  var s = JSON.parse(raw);
  if (new Date().getTime() > s.expiry) {
    cache.remove('sess_' + token);
    PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
    return { valid: false };
  }
  return { valid: true, teacherId: s.teacherId, name: s.name, role: s.role };
}

function logout(token) {
  if (token) {
    CacheService.getScriptCache().remove('sess_' + token);
    PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
  }
  return { success: true };
}

function _requireAuth(token) {
  var s = checkSession(token);
  if (!s.valid) throw new Error('กรุณาเข้าสู่ระบบใหม่');
  return s;
}

function _requireRole(token, roles) {
  var s = _requireAuth(token);
  var userRoles = s.role.split(',').map(function(r) { return r.trim(); });
  for (var i = 0; i < userRoles.length; i++) {
    if (roles.indexOf(userRoles[i]) >= 0) return s;
  }
  throw new Error('คุณไม่มีสิทธิ์ดำเนินการนี้');
}

// ==================== INITIAL DATA (Batch) ====================
function getInitialData(token) {
  var sess = _requireAuth(token);
  var settings = DataService.getSettings();
  var periods = DataService.getPeriodSettings();
  var role = sess.role;
  var isAdmin = role.indexOf('Admin') >= 0 || role.indexOf('Director') >= 0;
  var isScheduler = role.indexOf('Scheduler') >= 0 || isAdmin;

  var result = {
    user: { teacherId: sess.teacherId, name: sess.name, role: role },
    settings: settings.data || {},
    periods: (periods.data || []),
  };

  // Role-specific data
  if (isAdmin || isScheduler) {
    var teachers = DataService.getTeachers();
    var subjects = DataService.getSubjects();
    var classrooms = DataService.getClassrooms();
    result.teachers = teachers.data || [];
    result.subjects = subjects.data || [];
    result.classrooms = classrooms.data || [];
  }

  return result;
}

// ==================== WRAPPER FUNCTIONS ====================
// ── Teachers ──
function getTeachersData(token) { _requireAuth(token); return DataService.getTeachers(); }
function addTeacher(data, token) { return DataService.addTeacher(data, _requireRole(token, ['Admin']).teacherId); }
function updateTeacher(data, token) { return DataService.updateTeacher(data, _requireRole(token, ['Admin']).teacherId); }
function deleteTeacher(id, token) { return DataService.deleteTeacher(id, _requireRole(token, ['Admin']).teacherId); }
function importTeachers(rows, token) { return DataService.importTeachers(rows, _requireRole(token, ['Admin']).teacherId); }

// ── Subjects ──
function getSubjectsData(token) { _requireAuth(token); return DataService.getSubjects(); }
function addSubject(data, token) { return DataService.addSubject(data, _requireRole(token, ['Admin','Scheduler']).teacherId); }
function updateSubject(data, token) { return DataService.updateSubject(data, _requireRole(token, ['Admin','Scheduler']).teacherId); }
function deleteSubject(id, token) { return DataService.deleteSubject(id, _requireRole(token, ['Admin','Scheduler']).teacherId); }
function importSubjects(rows, token) { return DataService.importSubjects(rows, _requireRole(token, ['Admin','Scheduler']).teacherId); }

// ── Schedules ──
function getSchedulesData(token) { _requireAuth(token); return DataService.getSchedules(); }
function getMySchedule(token) {
  var s = _requireAuth(token);
  return DataService.getSchedulesByTeacher(s.teacherId);
}

// ── Periods ──
function getPeriodsData(token) { _requireAuth(token); return DataService.getPeriodSettings(); }
function addPeriod(data, token) { return DataService.addPeriod(data, _requireRole(token, ['Admin']).teacherId); }
function updatePeriod(data, token) { return DataService.updatePeriod(data, _requireRole(token, ['Admin']).teacherId); }
function deletePeriod(num, token) { return DataService.deletePeriod(num, _requireRole(token, ['Admin']).teacherId); }

// ── Classrooms ──
function getClassroomsData(token) { _requireAuth(token); return DataService.getClassrooms(); }
function addClassroom(data, token) { return DataService.addClassroom(data, _requireRole(token, ['Admin']).teacherId); }
function updateClassroom(data, token) { return DataService.updateClassroom(data, _requireRole(token, ['Admin']).teacherId); }
function deleteClassroomItem(id, token) { return DataService.deleteClassroom(id, _requireRole(token, ['Admin']).teacherId); }
function importClassrooms(rows, token) { return DataService.importClassrooms(rows, _requireRole(token, ['Admin']).teacherId); }

// ── Settings ──
function getSettingsData(token) { _requireRole(token, ['Admin']); return DataService.getSettings(); }
function saveSettings(obj, token) { return DataService.saveSettings(obj, _requireRole(token, ['Admin']).teacherId); }

// ── Leave ──
function submitLeave(data, token) { var s = _requireAuth(token); data.teacher_id = s.teacherId; return LeaveService.submitLeave(data, s.teacherId); }
function getMyLeaves(token) { var s = _requireAuth(token); return LeaveService.getLeavesByTeacher(s.teacherId); }
function getAllLeaves(token) { _requireRole(token, ['Admin','Director']); return LeaveService.getAllLeaves(); }
function approveLeave(leaveId, token) { return LeaveService.approveLeave(leaveId, _requireRole(token, ['Admin','Director']).teacherId); }
function rejectLeave(leaveId, reason, token) { return LeaveService.rejectLeave(leaveId, reason, _requireRole(token, ['Admin','Director']).teacherId); }
function getPendingLeaves(token) { _requireRole(token, ['Admin','Director','Scheduler']); return LeaveService.getPendingLeaves(); }
function cancelLeave(leaveId, token) { var s = _requireAuth(token); return LeaveService.cancelLeave(leaveId, s.teacherId); }
function updateLeave(leaveId, data, token) { var s = _requireAuth(token); return LeaveService.updateLeave(leaveId, data, s.teacherId); }

// ── Assignments ──
function getUnassignedSlots(token) { _requireRole(token, ['Admin','Scheduler','Director']); return LeaveService.getApprovedAndUnassignedLeaves(); }
function findAvailableTeachers(date, period, token) { _requireRole(token, ['Admin','Scheduler','Director']); return AssignmentService.findAvailableTeachersForSlot(date, period); }
function assignSubstitute(data, token) { return AssignmentService.assignSubstitute(data, _requireRole(token, ['Admin','Scheduler','Director']).teacherId); }
function getMySubstituteAssignments(token) { var s = _requireAuth(token); return AssignmentService.getAssignmentsByTeacher(s.teacherId); }
function getAllAssignments(token) { _requireAuth(token); return AssignmentService.getAllAssignments(); }
function cancelAssignment(assignmentId, token) { return AssignmentService.cancelAssignment(assignmentId, _requireRole(token, ['Admin','Scheduler','Director']).teacherId); }
function updateAssignmentTeacher(assignmentId, newSubTeacherId, token) { return AssignmentService.updateAssignmentTeacher(assignmentId, newSubTeacherId, _requireRole(token, ['Admin','Scheduler','Director']).teacherId); }

// ── Timetable (Drag & Drop) ──
function saveTimetable(data, token) { return TimetableService.saveScheduleBatch(data, _requireRole(token, ['Admin','Scheduler']).teacherId); }
function getTimetableByClass(className, token) { _requireAuth(token); return TimetableService.getScheduleByClass(className); }
function getPotnClassTimetable(className, token) { _requireAuth(token); return TimetableService.getPotnScheduleByClass(className); }
function getTimetableByTeacher(teacherId, token) { _requireAuth(token); return TimetableService.getScheduleByTeacher(teacherId); }
function checkConflicts(data, token) { _requireAuth(token); return TimetableService.checkConflicts(data); }
function clearClassSchedule(className, token) { return TimetableService.clearClassSchedule(className, _requireRole(token, ['Admin','Scheduler']).teacherId); }

// ── File Upload ──
function uploadFileBase64(base64Data, fileName, mimeType, token) {
  _requireAuth(token);
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    success: true,
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    size: file.getSize()
  };
}

// ── Dashboard ──
function getDashboardData(token) {
  _requireRole(token, ['Admin','Director','Scheduler']);
  var leaves = LeaveService.getAllLeaves();
  var assignments = AssignmentService.getAllAssignments();
  var teachers = DataService.getTeachers();

  var pending = 0, approved = 0, rejected = 0, totalAssigned = 0;
  (leaves.data || []).forEach(function(l) {
    if (l.status === 'pending') pending++;
    else if (l.status === 'approved') approved++;
    else if (l.status === 'rejected') rejected++;
  });
  totalAssigned = (assignments.data || []).length;

  // สถิติรายเดือน
  var monthlyLeaves = {};
  (leaves.data || []).forEach(function(l) {
    if (l.leave_date) {
      var m = String(l.leave_date).substring(5, 7);
      monthlyLeaves[m] = (monthlyLeaves[m] || 0) + 1;
    }
  });

  // ภาระงานสอนแทนรายครู
  var subCountByTeacher = {};
  (assignments.data || []).forEach(function(a) {
    var tid = a.substitute_teacher_id;
    subCountByTeacher[tid] = (subCountByTeacher[tid] || 0) + 1;
  });

  return {
    totalTeachers: (teachers.data || []).length,
    pendingLeaves: pending,
    approvedLeaves: approved,
    rejectedLeaves: rejected,
    totalAssignments: totalAssigned,
    monthlyLeaves: monthlyLeaves,
    substituteWorkload: subCountByTeacher
  };
}

// ── Teacher Personal Dashboard ──
function getMyDashboard(token) {
  var s = _requireAuth(token);
  var tid = s.teacherId;

  // 1. ตารางสอนทั้งสัปดาห์
  var schedResult = TimetableService.getScheduleByTeacher(tid);
  var schedule = schedResult.data || [];

  // 2. สถิติการลาของตัวเอง (รายเดือน)
  var leavesResult = LeaveService.getAllLeaves();
  var myMonthly = {};
  var myLeaveTotal = 0;
  (leavesResult.data || []).forEach(function(l) {
    if (l.teacher_id === tid && l.leave_date) {
      var m = String(l.leave_date).substring(5, 7);
      myMonthly[m] = (myMonthly[m] || 0) + 1;
      myLeaveTotal++;
    }
  });

  // 3. ภาระงานสอนแทน (เฉพาะตัวเอง) + enrich ด้วย subject_code, room
  var assignResult = AssignmentService.getAllAssignments();
  var allSchedules = DataService.getSchedules();
  var schedLookup = {};
  (allSchedules.data || []).forEach(function(sc) {
    schedLookup[sc.teacher_id + '_' + sc.day_of_week + '_' + sc.period] = sc;
  });

  var mySubstitute = [];
  (assignResult.data || []).forEach(function(a) {
    if (a.substitute_teacher_id === tid && a.status !== 'cancelled') {
      // Enrich: หา subject_code + room จากตารางสอนของครูต้นทาง
      if (a.schedule_date && a.original_teacher_id) {
        var dow = getDayOfWeekFromDate(a.schedule_date);
        var origSched = schedLookup[a.original_teacher_id + '_' + dow + '_' + a.period];
        if (origSched) {
          a.subject_code = origSched.subject_code || '';
          a.room = origSched.room || '';
          if (!a.subject) a.subject = origSched.subject_name || '';
        }
      }
      mySubstitute.push(a);
    }
  });

  return {
    success: true,
    schedule: schedule,
    myMonthlyLeaves: myMonthly,
    myLeaveTotal: myLeaveTotal,
    mySubstitute: mySubstitute
  };
}

// ── Profile ──
function getMyProfile(token) {
  var s = _requireAuth(token);
  var result = DataService.getTeachers();
  var teachers = result.data || [];
  for (var i = 0; i < teachers.length; i++) {
    if (teachers[i].teacher_id === s.teacherId) {
      var t = teachers[i];
      delete t.password; // ไม่ส่ง password กลับ
      return { success: true, data: t };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function updateMyProfile(data, token) {
  var s = _requireAuth(token);
  data.teacher_id = s.teacherId;
  // ครูแก้ได้เฉพาะบางฟิลด์
  return DataService.updateTeacherProfile(data, s.teacherId);
}

function changeMyPassword(oldPw, newPw, token) {
  var s = _requireAuth(token);
  return DataService.changePassword(s.teacherId, oldPw, newPw);
}


// ==================== HTTP API FOR GITHUB PAGES ====================
// Front-end ส่งข้อมูลด้วย fetch() แบบ POST และ Content-Type: text/plain
// รูปแบบ JSON: { "action": "ชื่อฟังก์ชัน", "args": [อาร์กิวเมนต์...] }
//
// หมายเหตุ: อย่าสร้าง doPost ซ้ำ หากมี doPost อยู่ในไฟล์ .gs อื่นของโปรเจกต์
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _jsonApiResponse_({
        success: false,
        message: 'ไม่พบข้อมูลคำขอจากเว็บไซต์'
      });
    }

    var request;
    try {
      request = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return _jsonApiResponse_({
        success: false,
        message: 'รูปแบบข้อมูล JSON ไม่ถูกต้อง'
      });
    }

    var action = request && request.action;
    var args = request && Array.isArray(request.args) ? request.args : [];

    if (!action || typeof action !== 'string') {
      return _jsonApiResponse_({
        success: false,
        message: 'ไม่ระบุ action'
      });
    }

    // อนุญาตเฉพาะฟังก์ชัน API ที่กำหนดไว้เท่านั้น
    // ป้องกันไม่ให้ผู้เรียก fetch() สั่งเรียกฟังก์ชันภายในของโปรเจกต์ได้โดยตรง
    var allowedActions = {
      // Authentication / session
      verifyUserLogin: verifyUserLogin,
      checkSession: checkSession,
      logout: logout,
      getInitialData: getInitialData,

      // Teachers
      getTeachersData: getTeachersData,
      addTeacher: addTeacher,
      updateTeacher: updateTeacher,
      deleteTeacher: deleteTeacher,
      importTeachers: importTeachers,

      // Subjects
      getSubjectsData: getSubjectsData,
      addSubject: addSubject,
      updateSubject: updateSubject,
      deleteSubject: deleteSubject,
      importSubjects: importSubjects,

      // Schedules
      getSchedulesData: getSchedulesData,
      getMySchedule: getMySchedule,

      // Periods
      getPeriodsData: getPeriodsData,
      addPeriod: addPeriod,
      updatePeriod: updatePeriod,
      deletePeriod: deletePeriod,

      // Classrooms
      getClassroomsData: getClassroomsData,
      addClassroom: addClassroom,
      updateClassroom: updateClassroom,
      deleteClassroomItem: deleteClassroomItem,
      importClassrooms: importClassrooms,

      // Settings
      getSettingsData: getSettingsData,
      saveSettings: saveSettings,

      // Leave
      submitLeave: submitLeave,
      getMyLeaves: getMyLeaves,
      getAllLeaves: getAllLeaves,
      approveLeave: approveLeave,
      rejectLeave: rejectLeave,
      getPendingLeaves: getPendingLeaves,
      cancelLeave: cancelLeave,
      updateLeave: updateLeave,

      // Substitute assignments
      getUnassignedSlots: getUnassignedSlots,
      findAvailableTeachers: findAvailableTeachers,
      assignSubstitute: assignSubstitute,
      getMySubstituteAssignments: getMySubstituteAssignments,
      getAllAssignments: getAllAssignments,
      cancelAssignment: cancelAssignment,
      updateAssignmentTeacher: updateAssignmentTeacher,

      // Timetable
      saveTimetable: saveTimetable,
      getTimetableByClass: getTimetableByClass,
      getPotnClassTimetable: getPotnClassTimetable,
      getTimetableByTeacher: getTimetableByTeacher,
      checkConflicts: checkConflicts,
      clearClassSchedule: clearClassSchedule,

      // Upload
      uploadFileBase64: uploadFileBase64,

      // Dashboard
      getDashboardData: getDashboardData,
      getMyDashboard: getMyDashboard,

      // Profile
      getMyProfile: getMyProfile,
      updateMyProfile: updateMyProfile,
      changeMyPassword: changeMyPassword
    };

    if (!allowedActions[action]) {
      return _jsonApiResponse_({
        success: false,
        message: 'ไม่อนุญาตให้เรียก API: ' + action
      });
    }

    var handler = allowedActions[action];
    if (typeof handler !== 'function') {
      return _jsonApiResponse_({
        success: false,
        message: 'ไม่พบฟังก์ชัน API: ' + action
      });
    }

    var result = handler.apply(null, args);
    return _jsonApiResponse_(result);

  } catch (error) {
    return _jsonApiResponse_({
      success: false,
      message: error && error.message ? error.message : String(error)
    });
  }
}

// ส่งผลลัพธ์กลับเป็น JSON
function _jsonApiResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data === undefined ? null : data))
    .setMimeType(ContentService.MimeType.JSON);
}
