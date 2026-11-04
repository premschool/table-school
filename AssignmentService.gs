var AssignmentService = (function () {
  return {

    getAllAssignments: function () {
      var cacheKey = 'assignments_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(SUBSTITUTE_ASSIGNMENTS_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        var headers = data[0];
        var teachersResult = DataService.getTeachers();
        var tMap = new Map();
        (teachersResult.data || []).forEach(function(t) {
          tMap.set(t.teacher_id, (t.prefix || '') + t.first_name + ' ' + t.last_name);
        });
        var assignments = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) {
            var obj = _rowToObj(data[i], headers);
            obj.original_teacher_name = tMap.get(obj.original_teacher_id) || obj.original_teacher_id;
            obj.substitute_teacher_name = tMap.get(obj.substitute_teacher_id) || obj.substitute_teacher_id;
            assignments.push(obj);
          }
        }
        var response = { success: true, data: assignments };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('assignments', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getAssignmentsByTeacher: function (teacherId) {
      try {
        var result = this.getAllAssignments();
        if (!result.success) return result;
        var filtered = (result.data || []).filter(function(a) {
          return a.substitute_teacher_id === teacherId || a.original_teacher_id === teacherId;
        });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    /**
     * ⚡ O(T+S+A) Algorithm — ค้นหาครูที่ว่าง
     * 1. สร้าง Busy Set จากตารางสอนปกติ (วันนั้น + คาบนั้น)
     * 2. สร้าง Busy Set จากงานสอนแทนที่จัดแล้ว (วันนั้น + คาบนั้น)
     * 3. สร้าง Leave Set จากครูที่ลาในวันนั้น
     * 4. return ครูที่ไม่อยู่ใน Busy/Leave Set
     */
    findAvailableTeachersForSlot: function (dateStr, period) {
      try {
        var dayOfWeek = getDayOfWeekFromDate(dateStr);
        var periodStr = String(period);

        // 1. ดึงข้อมูลทั้งหมด (cached)
        var teachersResult = DataService.getTeachers();
        var schedulesResult = DataService.getSchedules();
        var assignmentsResult = this.getAllAssignments();
        var leavesResult = LeaveService.getAllLeaves();

        var teachers = teachersResult.data || [];
        var schedules = schedulesResult.data || [];
        var assignments = assignmentsResult.data || [];
        var leaves = leavesResult.data || [];

        // 2. Busy Set จากตารางสอนปกติ
        var busyFromSchedule = {};
        for (var s = 0; s < schedules.length; s++) {
          if (schedules[s].day_of_week === dayOfWeek && String(schedules[s].period) === periodStr) {
            busyFromSchedule[schedules[s].teacher_id] = true;
          }
        }

        // 3. Busy Set จากงานสอนแทนที่จัดแล้ว (วันเดียวกัน คาบเดียวกัน)
        var busyFromAssignment = {};
        for (var a = 0; a < assignments.length; a++) {
          if (assignments[a].schedule_date === dateStr &&
              String(assignments[a].period) === periodStr &&
              assignments[a].status !== 'cancelled') {
            busyFromAssignment[assignments[a].substitute_teacher_id] = true;
          }
        }

        // 4. Leave Set — ครูที่ลาในวันนั้น
        var onLeave = {};
        for (var l = 0; l < leaves.length; l++) {
          if (leaves[l].leave_date === dateStr && leaves[l].status === 'approved') {
            var periods = leaves[l].periods;
            if (typeof periods === 'string') {
              try { periods = JSON.parse(periods); } catch (e) { periods = [periods]; }
            }
            if (!Array.isArray(periods)) periods = [periods];
            for (var p = 0; p < periods.length; p++) {
              if (String(periods[p]) === periodStr) {
                onLeave[leaves[l].teacher_id] = true;
                break;
              }
            }
          }
        }

        // 5. Filter ครูที่ว่าง
        var available = [];
        for (var t = 0; t < teachers.length; t++) {
          var tid = teachers[t].teacher_id;
          if (teachers[t].status !== 'active') continue;
          if (busyFromSchedule[tid]) continue;
          if (busyFromAssignment[tid]) continue;
          if (onLeave[tid]) continue;
          available.push({
            teacher_id: tid,
            name: (teachers[t].prefix || '') + teachers[t].first_name + ' ' + teachers[t].last_name,
            department: teachers[t].department || '',
            substitute_count: this._countSubstituteThisWeek(tid, dateStr, assignments)
          });
        }

        // Sort: ครูที่สอนแทนน้อยกว่า ขึ้นก่อน
        available.sort(function(a, b) { return a.substitute_count - b.substitute_count; });

        return { success: true, data: available, date: dateStr, period: periodStr, day: dayOfWeek };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    _countSubstituteThisWeek: function (teacherId, dateStr, assignments) {
      var targetDate = new Date(dateStr + 'T00:00:00');
      var dayOfWeek = targetDate.getDay();
      var monday = new Date(targetDate);
      monday.setDate(monday.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
      var friday = new Date(monday);
      friday.setDate(friday.getDate() + 4);
      var monStr = formatDateForSheet(monday);
      var friStr = formatDateForSheet(friday);

      var count = 0;
      for (var i = 0; i < assignments.length; i++) {
        if (assignments[i].substitute_teacher_id === teacherId &&
            assignments[i].status !== 'cancelled' &&
            assignments[i].schedule_date >= monStr &&
            assignments[i].schedule_date <= friStr) {
          count++;
        }
      }
      return count;
    },

    assignSubstitute: function (data, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler', 'Director']);
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var sheets = getSheets();
        var assignmentId = generateId('ASG');

        sheets.assignments.appendRow([
          assignmentId,
          data.leave_id,
          data.original_teacher_id,
          data.substitute_teacher_id,
          data.schedule_date,
          data.period,
          data.class_name || '',
          data.subject || '',
          'assigned',
          false,
          requestorId,
          formatDateTimeForSheet(new Date())
        ]);
        _invalidateCache('assignments');
        logAction('ASSIGN_SUBSTITUTE', assignmentId + ' sub:' + data.substitute_teacher_id, requestorId);

        // แจ้ง Telegram ให้ครูสอนแทน
        try {
          var teachersResult = DataService.getTeachers();
          var subTeacher = null;
          (teachersResult.data || []).forEach(function(t) {
            if (t.teacher_id === data.substitute_teacher_id) subTeacher = t;
          });
          if (subTeacher && subTeacher.telegram_chat_id) {
            TelegramService.notifySubstituteAssignment_({
              substitute_teacher_id: data.substitute_teacher_id,
              schedule_date: data.schedule_date,
              period: data.period,
              subject: data.subject,
              class_name: data.class_name
            }, subTeacher.telegram_chat_id, assignmentId);
          }
        } catch (tgErr) { console.warn('Telegram notify sub error: ' + tgErr.message); }

        return { success: true, message: 'จัดครูสอนแทนสำเร็จ', assignmentId: assignmentId };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    cancelAssignment: function (assignmentId, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler', 'Director']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(SUBSTITUTE_ASSIGNMENTS_SHEET, assignmentId, 'assignment_id');
        if (row === -1) return { success: false, message: 'ไม่พบการจัดสอนแทน' };
        var map = getColumnMap(SUBSTITUTE_ASSIGNMENTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SUBSTITUTE_ASSIGNMENTS_SHEET);
        sheet.getRange(row, map.status + 1).setValue('cancelled');
        _invalidateCache('assignments');
        logAction('CANCEL_ASSIGNMENT', assignmentId, requestorId);
        return { success: true, message: 'ยกเลิกการจัดสอนแทนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateAssignmentTeacher: function (assignmentId, newSubTeacherId, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler', 'Director']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(SUBSTITUTE_ASSIGNMENTS_SHEET, assignmentId, 'assignment_id');
        if (row === -1) return { success: false, message: 'ไม่พบการจัดสอนแทน' };
        var map = getColumnMap(SUBSTITUTE_ASSIGNMENTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SUBSTITUTE_ASSIGNMENTS_SHEET);

        var currentStatus = sheet.getRange(row, map.status + 1).getValue();
        if (currentStatus !== 'assigned') {
          return { success: false, message: 'แก้ไขได้เฉพาะรายการที่จัดแล้วเท่านั้น' };
        }

        sheet.getRange(row, map.substitute_teacher_id + 1).setValue(newSubTeacherId);
        sheet.getRange(row, map.assigned_by + 1).setValue(requestorId);
        sheet.getRange(row, map.assigned_at + 1).setValue(formatDateTimeForSheet(new Date()));
        _invalidateCache('assignments');
        logAction('UPDATE_ASSIGNMENT', assignmentId + ' new_sub:' + newSubTeacherId, requestorId);

        // แจ้ง Telegram
        try {
          var teachersResult = DataService.getTeachers();
          var subTeacher = null;
          (teachersResult.data || []).forEach(function(t) {
            if (t.teacher_id === newSubTeacherId) subTeacher = t;
          });
          if (subTeacher && subTeacher.telegram_chat_id) {
            var schedDate = sheet.getRange(row, map.schedule_date + 1).getValue();
            var period = sheet.getRange(row, map.period + 1).getValue();
            var subject = sheet.getRange(row, map.subject + 1).getValue();
            var className = sheet.getRange(row, map.class_name + 1).getValue();
            TelegramService.notifySubstituteAssignment_({
              substitute_teacher_id: newSubTeacherId,
              schedule_date: schedDate,
              period: period,
              subject: subject,
              class_name: className
            }, subTeacher.telegram_chat_id, assignmentId);
          }
        } catch (tgErr) { console.warn('Telegram notify error: ' + tgErr.message); }

        return { success: true, message: 'แก้ไขครูสอนแทนสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    }
  };
})();
