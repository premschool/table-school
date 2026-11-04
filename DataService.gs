var DataService = (function () {
  return {

    // ==================== SETTINGS ====================
    getSettings: function () {
      var cacheKey = 'settings_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(SETTINGS_SHEET);
        var map = getColumnMap(SETTINGS_SHEET);
        var settings = {};
        for (var i = 1; i < data.length; i++) {
          var key = data[i][map.key];
          var val = data[i][map.value];
          if (key) settings[key] = (val === null || val === undefined || val === '') ? '' : String(val);
        }
        var response = { success: true, data: settings };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('settings', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: {} };
      }
    },

    saveSettings: function (obj, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        var sheet = sheets.settings;
        var data = sheet.getDataRange().getValues();
        var map = getColumnMap(SETTINGS_SHEET);
        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
          var found = false;
          for (var i = 1; i < data.length; i++) {
            if (data[i][map.key] === keys[k]) {
              sheet.getRange(i + 1, map.value + 1).setValue(obj[keys[k]]);
              found = true;
              break;
            }
          }
          if (!found) sheet.appendRow([keys[k], obj[keys[k]]]);
        }
        _invalidateCache('settings');
        CACHE.remove('enable_logs');
        logAction('SAVE_SETTINGS', JSON.stringify(keys), requestorId);
        return { success: true, message: 'บันทึกตั้งค่าสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    // ==================== TEACHERS ====================
    getTeachers: function () {
      var cacheKey = 'teachers_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) {
        // Protect against cache entries created by the older version that included password hashes.
        (cached.data || []).forEach(function(t) { delete t.password; });
        return cached;
      }
      try {
        var data = getSheetData(TEACHERS_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        var headers = data[0];
        var teachers = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) {
            var teacher = _rowToObj(data[i], headers);
            delete teacher.password; // ห้ามส่ง hash หรือข้อมูล credential ไปยังหน้าเว็บ
            teachers.push(teacher);
          }
        }
        var response = { success: true, data: teachers };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('teachers', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    // เฉพาะฝั่ง server สำหรับตรวจสอบ Login; ห้ามส่งผลลัพธ์นี้กลับ client
    getTeacherCredential: function (userId) {
      var data = getSheetData(TEACHERS_SHEET);
      if (data.length < 2) return null;
      var headers = data[0], map = _buildHeaderIndex(headers);
      if (map.user_id === undefined || map.password === undefined) return null;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][map.user_id] || '') === String(userId)) return _rowToObj(data[i], headers);
      }
      return null;
    },

    upgradeLegacyCredential: function (teacherId, newHash) {
      var row = findRowIndexById(TEACHERS_SHEET, teacherId, 'teacher_id');
      var map = getColumnMap(TEACHERS_SHEET);
      if (row > 0 && map.password !== undefined) {
        getSheets().teachers.getRange(row, map.password + 1).setValue(newHash);
        _invalidateCache('teachers');
      }
    },

    addTeacher: function (teacherData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var passwordError = _passwordValidationMessage_(teacherData.password, teacherData.role || 'Teacher');
      if (passwordError) return { success: false, message: passwordError };
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        // Check duplicate user_id
        var existing = this.getTeachers().data || [];
        var uid = teacherData.user_id || teacherData.teacher_id || '';
        for (var i = 0; i < existing.length; i++) {
          if (existing[i].user_id === uid) return { success: false, message: 'ชื่อผู้ใช้ "' + uid + '" มีอยู่แล้วในระบบ' };
        }
        var hashedPw = _hashPw(teacherData.password);
        sheets.teachers.appendRow([
          teacherData.teacher_id || generateId('T'),
          teacherData.user_id || teacherData.teacher_id || '',
          hashedPw,
          teacherData.prefix || '',
          teacherData.first_name,
          teacherData.last_name,
          teacherData.department || '',
          teacherData.teaching_room || '',
          JSON.stringify(teacherData.subjects || []),
          JSON.stringify(teacherData.grade_levels || []),
          teacherData.phone || '',
          teacherData.email || '',
          teacherData.telegram_chat_id || '',
          teacherData.max_substitute_periods || 5,
          'active',
          teacherData.role || 'Teacher'
        ]);
        _invalidateCache('teachers');
        logAction('ADD_TEACHER', teacherData.teacher_id, requestorId);
        return { success: true, message: 'เพิ่มครูสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateTeacher: function (teacherData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        var row = findRowIndexById(TEACHERS_SHEET, teacherData.teacher_id, 'teacher_id');
        if (row === -1) return { success: false, message: 'ไม่พบข้อมูลครู' };
        var map = getColumnMap(TEACHERS_SHEET);
        var sheet = sheets.teachers;
        var currentRole = map.role === undefined ? 'Teacher' : sheet.getRange(row, map.role + 1).getValue();
        var targetRole = teacherData.role || currentRole || 'Teacher';
        if (teacherData.role && _roleHasTeacher_(currentRole) !== _roleHasTeacher_(targetRole) && !teacherData.new_password) {
          return { success: false, message: 'เมื่อเปลี่ยนสิทธิ์ระหว่างครูกับผู้ใช้ประเภทอื่น กรุณากำหนดรหัสผ่านใหม่ด้วย' };
        }
        if (teacherData.new_password) {
          var passwordError = _passwordValidationMessage_(teacherData.new_password, targetRole);
          if (passwordError) return { success: false, message: passwordError };
          if (map.password === undefined) return { success: false, message: 'ไม่พบคอลัมน์ password ในชีต Teachers' };
        }
        if (map.user_id !== undefined) sheet.getRange(row, map.user_id + 1).setValue(teacherData.user_id || '');
        if (map.prefix !== undefined) sheet.getRange(row, map.prefix + 1).setValue(teacherData.prefix || '');
        if (map.first_name !== undefined) sheet.getRange(row, map.first_name + 1).setValue(teacherData.first_name);
        if (map.last_name !== undefined) sheet.getRange(row, map.last_name + 1).setValue(teacherData.last_name);
        if (map.department !== undefined) sheet.getRange(row, map.department + 1).setValue(teacherData.department || '');
        if (map.teaching_room !== undefined) sheet.getRange(row, map.teaching_room + 1).setValue(teacherData.teaching_room || '');
        if (map.subjects !== undefined) sheet.getRange(row, map.subjects + 1).setValue(JSON.stringify(teacherData.subjects || []));
        if (map.grade_levels !== undefined) sheet.getRange(row, map.grade_levels + 1).setValue(JSON.stringify(teacherData.grade_levels || []));
        if (map.phone !== undefined) sheet.getRange(row, map.phone + 1).setValue(teacherData.phone || '');
        if (map.email !== undefined) sheet.getRange(row, map.email + 1).setValue(teacherData.email || '');
        if (map.telegram_chat_id !== undefined) sheet.getRange(row, map.telegram_chat_id + 1).setValue(teacherData.telegram_chat_id || '');
        if (map.max_substitute_periods !== undefined) sheet.getRange(row, map.max_substitute_periods + 1).setValue(teacherData.max_substitute_periods || 5);
        if (map.status !== undefined) sheet.getRange(row, map.status + 1).setValue(teacherData.status || 'active');
        if (map.role !== undefined) sheet.getRange(row, map.role + 1).setValue(teacherData.role || 'Teacher');
        // Reset password ถ้ามี
        if (teacherData.new_password) {
          var hash = _hashPw(teacherData.new_password);
          if (map.password === undefined) return { success: false, message: 'ไม่พบคอลัมน์ password ในชีต Teachers' };
          sheet.getRange(row, map.password + 1).setValue(hash);
        }
        _invalidateCache('teachers');
        logAction('UPDATE_TEACHER', teacherData.teacher_id, requestorId);
        return { success: true, message: 'แก้ไขข้อมูลครูสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateTeacherProfile: function (data, teacherId) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(TEACHERS_SHEET, teacherId, 'teacher_id');
        if (row === -1) return { success: false, message: 'ไม่พบข้อมูล' };
        var map = getColumnMap(TEACHERS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(TEACHERS_SHEET);
        if (data.phone !== undefined && map.phone !== undefined) sheet.getRange(row, map.phone + 1).setValue(data.phone);
        if (data.email !== undefined && map.email !== undefined) sheet.getRange(row, map.email + 1).setValue(data.email);
        if (data.telegram_chat_id !== undefined && map.telegram_chat_id !== undefined) sheet.getRange(row, map.telegram_chat_id + 1).setValue(data.telegram_chat_id);
        _invalidateCache('teachers');
        return { success: true, message: 'บันทึกข้อมูลส่วนตัวสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    changePassword: function (teacherId, oldPw, newPw) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(TEACHERS_SHEET, teacherId, 'teacher_id');
        if (row === -1) return { success: false, message: 'ไม่พบข้อมูล' };
        var map = getColumnMap(TEACHERS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(TEACHERS_SHEET);
        var role = map.role === undefined ? 'Teacher' : sheet.getRange(row, map.role + 1).getValue();
        var passwordError = _passwordValidationMessage_(newPw, role);
        if (passwordError) return { success: false, message: passwordError };
        if (map.password === undefined) return { success: false, message: 'ไม่พบคอลัมน์ password ในชีต Teachers' };
        var currentPw = sheet.getRange(row, map.password + 1).getValue();
        if (!_verifyPassword(oldPw, currentPw)) return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
        sheet.getRange(row, map.password + 1).setValue(_hashPw(newPw));
        _invalidateCache('teachers');
        logAction('CHANGE_PASSWORD', teacherId, teacherId);
        return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    deleteTeacher: function (teacherId, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(TEACHERS_SHEET, teacherId, 'teacher_id');
        if (row === -1) return { success: false, message: 'ไม่พบข้อมูลครู' };
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        ss.getSheetByName(TEACHERS_SHEET).deleteRow(row);
        _invalidateCache('teachers');
        logAction('DELETE_TEACHER', teacherId, requestorId);
        return { success: true, message: 'ลบข้อมูลครูสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    importTeachers: function (rows, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var sheets = getSheets();
        var sheet = sheets.teachers;
        var existing = this.getTeachers().data || [];
        var existingUserIds = {};
        existing.forEach(function(t) { existingUserIds[t.user_id] = true; });

        var added = 0, skipped = 0;
        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          var uid = r.user_id || r.teacher_id || '';
          if (!uid || !r.first_name || existingUserIds[uid]) { skipped++; continue; }
          // Import without an explicit password creates an account requiring admin password setup.
          var role = r.role || 'Teacher';
          var passwordError = r.password ? _passwordValidationMessage_(r.password, role) : '';
          if (passwordError) { skipped++; continue; }
          var pw = r.password ? _hashPw(r.password) : '';
          sheet.appendRow([
            r.teacher_id || generateId('T'),
            r.user_id || r.teacher_id || '',
            pw,
            r.prefix || '', r.first_name || '', r.last_name || '',
            r.department || '', r.teaching_room || '',
            JSON.stringify(r.subjects || []),
            JSON.stringify(r.grade_levels || []),
            r.phone || '', r.email || '', r.telegram_chat_id || '',
            r.max_substitute_periods || 5, 'active', role
          ]);
          added++;
        }
        _invalidateCache('teachers');
        logAction('IMPORT_TEACHERS', 'Added: ' + added + ', Skipped: ' + skipped, requestorId);
        return { success: true, message: 'นำเข้าสำเร็จ ' + added + ' รายการ (ข้าม ' + skipped + ' รายการ)' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    // ==================== SUBJECTS ====================
    getSubjects: function () {
      var cacheKey = 'subjects_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(SUBJECTS_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        var headers = data[0];
        var subjects = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) subjects.push(_rowToObj(data[i], headers));
        }
        var response = { success: true, data: subjects };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('subjects', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    addSubject: function (subData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        sheets.subjects.appendRow([
          subData.subject_id || generateId('SUB'),
          subData.subject_code || '',
          subData.subject_name,
          subData.type || 'พื้นฐาน',
          subData.grade_levels || '[]',
          subData.class_names || '[]',
          subData.total_periods || 1,
          subData.room || ''
        ]);
        _invalidateCache('subjects');
        logAction('ADD_SUBJECT', subData.subject_code, requestorId);
        return { success: true, message: 'เพิ่มรายวิชาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateSubject: function (subData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(SUBJECTS_SHEET, subData.subject_id, 'subject_id');
        if (row === -1) return { success: false, message: 'ไม่พบรายวิชา' };
        var map = getColumnMap(SUBJECTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SUBJECTS_SHEET);
        if (map.subject_code !== undefined) sheet.getRange(row, map.subject_code + 1).setValue(subData.subject_code || '');
        if (map.subject_name !== undefined) sheet.getRange(row, map.subject_name + 1).setValue(subData.subject_name);
        if (map.type !== undefined) sheet.getRange(row, map.type + 1).setValue(subData.type || 'พื้นฐาน');
        if (map.grade_levels !== undefined) sheet.getRange(row, map.grade_levels + 1).setValue(subData.grade_levels || '[]');
        if (map.class_names !== undefined) sheet.getRange(row, map.class_names + 1).setValue(subData.class_names || '[]');
        if (map.total_periods !== undefined) sheet.getRange(row, map.total_periods + 1).setValue(subData.total_periods || 1);
        if (map.room !== undefined) sheet.getRange(row, map.room + 1).setValue(subData.room || '');
        _invalidateCache('subjects');
        logAction('UPDATE_SUBJECT', subData.subject_id, requestorId);
        return { success: true, message: 'แก้ไขรายวิชาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    deleteSubject: function (subjectId, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(SUBJECTS_SHEET, subjectId, 'subject_id');
        if (row === -1) return { success: false, message: 'ไม่พบรายวิชา' };
        SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SUBJECTS_SHEET).deleteRow(row);
        _invalidateCache('subjects');
        logAction('DELETE_SUBJECT', subjectId, requestorId);
        return { success: true, message: 'ลบรายวิชาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    importSubjects: function (rows, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var sheets = getSheets();
        var sheet = sheets.subjects;
        var existing = this.getSubjects().data || [];
        var existingCodes = {};
        existing.forEach(function(s) { existingCodes[s.subject_code || s.subject_id] = true; });

        var added = 0, skipped = 0;
        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          if (!r.subject_name) { skipped++; continue; }
          var code = r.subject_code || '';
          if (code && existingCodes[code]) { skipped++; continue; }
          sheet.appendRow([
            generateId('SUB'),
            code,
            r.subject_name,
            r.type || 'พื้นฐาน',
            r.grade_levels || '[]',
            r.class_names || '[]',
            r.total_periods || 1,
            r.room || ''
          ]);
          added++;
        }
        _invalidateCache('subjects');
        logAction('IMPORT_SUBJECTS', 'Added: ' + added + ', Skipped: ' + skipped, requestorId);
        return { success: true, message: 'นำเข้าสำเร็จ ' + added + ' รายการ (ข้าม ' + skipped + ' รายการ)' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    // ==================== SCHEDULES ====================
    getSchedules: function () {
      var cacheKey = 'schedules_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(SCHEDULES_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        // Build lookup maps
        var subjectsResult = this.getSubjects();
        var teachersResult = this.getTeachers();
        var subjectsMap = new Map();
        (subjectsResult.data || []).forEach(function(s) { subjectsMap.set(s.subject_id, s); });
        var teachersMap = new Map();
        (teachersResult.data || []).forEach(function(t) {
          teachersMap.set(t.teacher_id, (t.prefix || '') + t.first_name + ' ' + t.last_name);
        });

        var headers = data[0];
        var map = _buildHeaderIndex(headers);
        var schedules = [];
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row[map.schedule_id]) {
            var subId = row[map.subject_id];
            var sub = subjectsMap.get(subId) || {};
            schedules.push({
              schedule_id: row[map.schedule_id],
              day_of_week: row[map.day_of_week],
              period: row[map.period],
              subject_id: subId,
              subject_code: sub.subject_code || '',
              subject_name: sub.subject_name || '',
              class_name: row[map.class_name],
              teacher_id: row[map.teacher_id],
              teacher_name: teachersMap.get(row[map.teacher_id]) || row[map.teacher_id],
              room: row[map.room] || sub.room || ''
            });
          }
        }
        var response = { success: true, data: schedules };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('schedules', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getSchedulesByTeacher: function (teacherId) {
      try {
        var result = this.getSchedules();
        if (!result.success) return result;
        var filtered = (result.data || []).filter(function(s) { return s.teacher_id === teacherId; });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getSchedulesByTeacherAndDay: function (teacherId, dayOfWeek) {
      try {
        var result = this.getSchedules();
        if (!result.success) return result;
        var filtered = (result.data || []).filter(function(s) {
          return s.teacher_id === teacherId && s.day_of_week === dayOfWeek;
        });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    // ==================== PERIODS ====================
    getPeriodSettings: function () {
      var cacheKey = 'periods_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(PERIODS_SHEET);
        var map = getColumnMap(PERIODS_SHEET);
        var periods = [];
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row[map.periodNumber] !== '' && row[map.periodNumber] !== null) {
            periods.push({
              periodNumber: row[map.periodNumber],
              periodName: row[map.periodName],
              periodTime: row[map.periodTime],
              periodType: map.periodType !== undefined ? row[map.periodType] : 'เรียน',
              duration: map.duration !== undefined ? row[map.duration] : 50
            });
          }
        }
        periods.sort(function(a, b) { return Number(a.periodNumber) - Number(b.periodNumber); });
        var response = { success: true, data: periods };
        _chunkedCachePut(cacheKey, response, 3600);
        _registerCacheKey('periods', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    addPeriod: function (periodData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        sheets.periods.appendRow([
          periodData.periodNumber, periodData.periodName,
          periodData.periodTime, periodData.periodType || 'เรียน',
          periodData.duration || 50
        ]);
        _invalidateCache('periods');
        logAction('ADD_PERIOD', 'Period ' + periodData.periodNumber, requestorId);
        return { success: true, message: 'เพิ่มคาบเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updatePeriod: function (periodData, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(PERIODS_SHEET, periodData.original_period_number, 'periodNumber');
        if (row === -1) return { success: false, message: 'ไม่พบคาบเรียน' };
        var map = getColumnMap(PERIODS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(PERIODS_SHEET);
        if (map.periodNumber !== undefined) sheet.getRange(row, map.periodNumber + 1).setValue(periodData.periodNumber);
        if (map.periodName !== undefined) sheet.getRange(row, map.periodName + 1).setValue(periodData.periodName);
        if (map.periodTime !== undefined) sheet.getRange(row, map.periodTime + 1).setValue(periodData.periodTime);
        if (map.periodType !== undefined) sheet.getRange(row, map.periodType + 1).setValue(periodData.periodType || 'เรียน');
        if (map.duration !== undefined) sheet.getRange(row, map.duration + 1).setValue(periodData.duration || 50);
        _invalidateCache('periods');
        logAction('UPDATE_PERIOD', 'Period ' + periodData.periodNumber, requestorId);
        return { success: true, message: 'แก้ไขคาบเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    deletePeriod: function (periodNumber, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(PERIODS_SHEET, periodNumber, 'periodNumber');
        if (row === -1) return { success: false, message: 'ไม่พบคาบเรียน' };
        SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PERIODS_SHEET).deleteRow(row);
        _invalidateCache('periods');
        logAction('DELETE_PERIOD', 'Period ' + periodNumber, requestorId);
        return { success: true, message: 'ลบคาบเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    // ==================== CLASSROOMS ====================
    getClassrooms: function () {
      var cacheKey = 'classrooms_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(CLASSROOMS_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        var headers = data[0];
        var classrooms = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) classrooms.push(_rowToObj(data[i], headers));
        }
        classrooms.sort(function(a, b) {
          if (a.grade_level === b.grade_level) return Number(a.room_number) - Number(b.room_number);
          return String(a.grade_level).localeCompare(String(b.grade_level));
        });
        var response = { success: true, data: classrooms };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('classrooms', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    addClassroom: function (data, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sheets = getSheets();
        var className = data.grade_level + '/' + data.room_number;
        sheets.classrooms.appendRow([
          generateId('CLS'),
          data.grade_level,
          data.room_number,
          data.class_name || className,
          data.homeroom_teacher_id || '',
          data.student_count || 0
        ]);
        _invalidateCache('classrooms');
        logAction('ADD_CLASSROOM', className, requestorId);
        return { success: true, message: 'เพิ่มชั้นเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateClassroom: function (data, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(CLASSROOMS_SHEET, data.classroom_id, 'classroom_id');
        if (row === -1) return { success: false, message: 'ไม่พบชั้นเรียน' };
        var map = getColumnMap(CLASSROOMS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(CLASSROOMS_SHEET);
        if (map.grade_level !== undefined) sheet.getRange(row, map.grade_level + 1).setValue(data.grade_level);
        if (map.room_number !== undefined) sheet.getRange(row, map.room_number + 1).setValue(data.room_number);
        var className = data.class_name || (data.grade_level + '/' + data.room_number);
        if (map.class_name !== undefined) sheet.getRange(row, map.class_name + 1).setValue(className);
        if (map.homeroom_teacher_id !== undefined) sheet.getRange(row, map.homeroom_teacher_id + 1).setValue(data.homeroom_teacher_id || '');
        if (map.student_count !== undefined) sheet.getRange(row, map.student_count + 1).setValue(data.student_count || 0);
        _invalidateCache('classrooms');
        logAction('UPDATE_CLASSROOM', className, requestorId);
        return { success: true, message: 'แก้ไขชั้นเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    deleteClassroom: function (classroomId, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(CLASSROOMS_SHEET, classroomId, 'classroom_id');
        if (row === -1) return { success: false, message: 'ไม่พบชั้นเรียน' };
        SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CLASSROOMS_SHEET).deleteRow(row);
        _invalidateCache('classrooms');
        logAction('DELETE_CLASSROOM', classroomId, requestorId);
        return { success: true, message: 'ลบชั้นเรียนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    importClassrooms: function (rows, requestorId) {
      validateRequestorRole_(requestorId, ['Admin']);
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var sheets = getSheets();
        var sheet = sheets.classrooms;
        var existing = this.getClassrooms().data || [];
        var existingNames = {};
        existing.forEach(function(c) { existingNames[c.class_name] = true; });

        var added = 0, skipped = 0;
        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          var className = r.class_name || (r.grade_level + '/' + r.room_number);
          if (!r.grade_level || existingNames[className]) { skipped++; continue; }
          sheet.appendRow([
            generateId('CLS'),
            r.grade_level,
            r.room_number || 1,
            className,
            r.homeroom_teacher_id || '',
            r.student_count || 0
          ]);
          added++;
        }
        _invalidateCache('classrooms');
        logAction('IMPORT_CLASSROOMS', 'Added: ' + added + ', Skipped: ' + skipped, requestorId);
        return { success: true, message: 'นำเข้าสำเร็จ ' + added + ' รายการ (ข้าม ' + skipped + ' รายการ)' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    }
  };
})();
