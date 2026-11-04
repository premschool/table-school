var LeaveService = (function () {
  return {

    submitLeave: function (data, requestorId) {
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var sheets = getSheets();
        var periods = data.periods;
        if (typeof periods === 'object') periods = JSON.stringify(periods);

        // ตรวจลาซ้ำ: teacher_id + leave_date ที่ status ≠ rejected
        var allResult = this.getAllLeaves();
        var duplicate = (allResult.data || []).some(function(l) {
          return l.teacher_id === data.teacher_id && l.leave_date === data.leave_date && l.status !== 'rejected';
        });
        if (duplicate) {
          return { success: false, message: 'คุณมีใบลาในวันที่ ' + data.leave_date + ' แล้ว กรุณาแก้ไขใบลาเดิมแทน' };
        }

        var leaveId = generateId('LV');

        sheets.leaveRequests.appendRow([
          leaveId,
          data.teacher_id,
          data.leave_date,
          data.leave_type || 'ลากิจ',
          periods,
          data.reason || '',
          data.attachment_file_id || '',
          data.attachment_url || '',
          data.attachment_name || '',
          'pending',
          '',
          '',
          formatDateTimeForSheet(new Date())
        ]);
        _invalidateCache('leaves');
        logAction('SUBMIT_LEAVE', leaveId + ' date:' + data.leave_date, requestorId);

        // แจ้ง Telegram ถ้ามี
        try {
          var teachersResult = DataService.getTeachers();
          var teachers = teachersResult.data || [];
          var teacherName = requestorId;
          for (var i = 0; i < teachers.length; i++) {
            if (teachers[i].teacher_id === requestorId) {
              teacherName = (teachers[i].prefix || '') + teachers[i].first_name + ' ' + teachers[i].last_name;
              break;
            }
          }
          TelegramService.notifyLeaveSubmittedToAdmin_({
            teacher_name: teacherName,
            leave_date: data.leave_date,
            leave_type: data.leave_type,
            periods: periods,
            reason: data.reason
          });
        } catch (tgErr) { console.warn('Telegram notify error: ' + tgErr.message); }

        return { success: true, message: 'ส่งใบลาสำเร็จ', leaveId: leaveId };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    getLeavesByTeacher: function (teacherId) {
      try {
        var allResult = this.getAllLeaves();
        if (!allResult.success) return allResult;
        var filtered = (allResult.data || []).filter(function(l) {
          return l.teacher_id === teacherId;
        });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getAllLeaves: function () {
      var cacheKey = 'leaves_all';
      var cached = _chunkedCacheGet(cacheKey);
      if (cached) return cached;
      try {
        var data = getSheetData(LEAVE_REQUESTS_SHEET);
        if (data.length < 2) return { success: true, data: [] };
        var headers = data[0];
        var teachersResult = DataService.getTeachers();
        var tMap = new Map();
        (teachersResult.data || []).forEach(function(t) {
          tMap.set(t.teacher_id, (t.prefix || '') + t.first_name + ' ' + t.last_name);
        });

        var leaves = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) {
            var obj = _rowToObj(data[i], headers);
            obj.teacher_name = tMap.get(obj.teacher_id) || obj.teacher_id;
            leaves.push(obj);
          }
        }
        leaves.sort(function(a, b) {
          return (b.created_at || '').localeCompare(a.created_at || '');
        });
        var response = { success: true, data: leaves };
        _chunkedCachePut(cacheKey, response, CACHE_TTL);
        _registerCacheKey('leaves', cacheKey);
        return response;
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getPendingLeaves: function () {
      try {
        var result = this.getAllLeaves();
        if (!result.success) return result;
        var pending = (result.data || []).filter(function(l) {
          return l.status === 'pending';
        });
        return { success: true, data: pending };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    approveLeave: function (leaveId, approverId) {
      validateRequestorRole_(approverId, ['Admin', 'Director']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(LEAVE_REQUESTS_SHEET, leaveId, 'leave_id');
        if (row === -1) return { success: false, message: 'ไม่พบใบลา' };
        var map = getColumnMap(LEAVE_REQUESTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(LEAVE_REQUESTS_SHEET);

        // ตรวจสถานะปัจจุบัน
        var currentStatus = sheet.getRange(row, map.status + 1).getValue();
        if (currentStatus !== 'pending') {
          return { success: false, message: 'ใบลานี้ถูกดำเนินการแล้ว (สถานะ: ' + currentStatus + ')' };
        }

        sheet.getRange(row, map.status + 1).setValue('approved');
        sheet.getRange(row, map.approved_by + 1).setValue(approverId);
        sheet.getRange(row, map.approved_at + 1).setValue(formatDateTimeForSheet(new Date()));
        _invalidateCache('leaves');
        logAction('APPROVE_LEAVE', leaveId, approverId);
        return { success: true, message: 'อนุมัติใบลาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    rejectLeave: function (leaveId, reason, rejecterId) {
      validateRequestorRole_(rejecterId, ['Admin', 'Director']);
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(LEAVE_REQUESTS_SHEET, leaveId, 'leave_id');
        if (row === -1) return { success: false, message: 'ไม่พบใบลา' };
        var map = getColumnMap(LEAVE_REQUESTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(LEAVE_REQUESTS_SHEET);

        var currentStatus = sheet.getRange(row, map.status + 1).getValue();
        if (currentStatus !== 'pending') {
          return { success: false, message: 'ใบลานี้ถูกดำเนินการแล้ว' };
        }

        sheet.getRange(row, map.status + 1).setValue('rejected');
        sheet.getRange(row, map.approved_by + 1).setValue(rejecterId);
        sheet.getRange(row, map.approved_at + 1).setValue(formatDateTimeForSheet(new Date()));
        if (reason && map.reason !== undefined) {
          var existingReason = sheet.getRange(row, map.reason + 1).getValue();
          sheet.getRange(row, map.reason + 1).setValue(existingReason + ' [ปฏิเสธ: ' + reason + ']');
        }
        _invalidateCache('leaves');
        logAction('REJECT_LEAVE', leaveId + ' reason:' + (reason || ''), rejecterId);
        return { success: true, message: 'ปฏิเสธใบลาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    getApprovedAndUnassignedLeaves: function () {
      try {
        var leavesResult = this.getAllLeaves();
        var assignResult = AssignmentService.getAllAssignments();
        var leaves = (leavesResult.data || []).filter(function(l) { return l.status === 'approved'; });
        var assignments = assignResult.data || [];

        // สร้าง Set ของ assignment keys (leave_id + period)
        var assignedSet = {};
        assignments.forEach(function(a) {
          assignedSet[a.leave_id + '_' + a.period] = true;
        });

        var unassigned = [];
        for (var i = 0; i < leaves.length; i++) {
          var leave = leaves[i];
          var periods = leave.periods;
          if (typeof periods === 'string') {
            try { periods = JSON.parse(periods); } catch (e) { periods = [periods]; }
          }
          if (!Array.isArray(periods)) periods = [periods];

          var dayOfWeek = getDayOfWeekFromDate(leave.leave_date);

          // หาข้อมูลตารางสอนของครูในวันนั้น
          var schedResult = DataService.getSchedulesByTeacherAndDay(leave.teacher_id, dayOfWeek);
          var schedules = schedResult.data || [];
          var schedMap = {};
          schedules.forEach(function(s) { schedMap[String(s.period)] = s; });

          for (var p = 0; p < periods.length; p++) {
            var periodKey = String(periods[p]);
            if (!assignedSet[leave.leave_id + '_' + periodKey]) {
              var sched = schedMap[periodKey] || {};
              unassigned.push({
                leave_id: leave.leave_id,
                teacher_id: leave.teacher_id,
                teacher_name: leave.teacher_name,
                leave_date: leave.leave_date,
                day_of_week: dayOfWeek,
                period: periodKey,
                class_name: sched.class_name || '',
                subject: sched.subject_name || sched.subject_code || '',
                room: sched.room || ''
              });
            }
          }
        }
        return { success: true, data: unassigned };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    cancelLeave: function (leaveId, requestorId) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(LEAVE_REQUESTS_SHEET, leaveId, 'leave_id');
        if (row === -1) return { success: false, message: 'ไม่พบใบลา' };
        var map = getColumnMap(LEAVE_REQUESTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(LEAVE_REQUESTS_SHEET);

        // ตรวจ ownership
        var ownerId = sheet.getRange(row, map.teacher_id + 1).getValue();
        if (ownerId !== requestorId) {
          return { success: false, message: 'คุณไม่มีสิทธิ์ยกเลิกใบลานี้' };
        }

        // ตรวจสถานะ
        var currentStatus = sheet.getRange(row, map.status + 1).getValue();
        if (currentStatus !== 'pending') {
          return { success: false, message: 'ยกเลิกได้เฉพาะใบลาที่รออนุมัติเท่านั้น (สถานะปัจจุบัน: ' + currentStatus + ')' };
        }

        sheet.deleteRow(row);
        _invalidateCache('leaves');
        logAction('CANCEL_LEAVE', leaveId, requestorId);
        return { success: true, message: 'ยกเลิกใบลาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    updateLeave: function (leaveId, data, requestorId) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var row = findRowIndexById(LEAVE_REQUESTS_SHEET, leaveId, 'leave_id');
        if (row === -1) return { success: false, message: 'ไม่พบใบลา' };
        var map = getColumnMap(LEAVE_REQUESTS_SHEET);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(LEAVE_REQUESTS_SHEET);

        var ownerId = sheet.getRange(row, map.teacher_id + 1).getValue();
        if (ownerId !== requestorId) {
          return { success: false, message: 'คุณไม่มีสิทธิ์แก้ไขใบลานี้' };
        }
        var currentStatus = sheet.getRange(row, map.status + 1).getValue();
        if (currentStatus !== 'pending') {
          return { success: false, message: 'แก้ไขได้เฉพาะใบลาที่รออนุมัติเท่านั้น' };
        }

        // ตรวจลาซ้ำ (ถ้าเปลี่ยนวันที่)
        var oldDate = sheet.getRange(row, map.leave_date + 1).getValue();
        if (data.leave_date && data.leave_date !== String(oldDate)) {
          var allResult = this.getAllLeaves();
          var dup = (allResult.data || []).some(function(l) {
            return l.teacher_id === requestorId && l.leave_date === data.leave_date && l.status !== 'rejected' && l.leave_id !== leaveId;
          });
          if (dup) return { success: false, message: 'คุณมีใบลาในวันที่ ' + data.leave_date + ' แล้ว' };
        }

        if (data.leave_date) sheet.getRange(row, map.leave_date + 1).setValue(data.leave_date);
        if (data.leave_type) sheet.getRange(row, map.leave_type + 1).setValue(data.leave_type);
        if (data.periods) {
          var p = data.periods;
          if (typeof p === 'object') p = JSON.stringify(p);
          sheet.getRange(row, map.periods + 1).setValue(p);
        }
        if (data.reason !== undefined) sheet.getRange(row, map.reason + 1).setValue(data.reason);

        _invalidateCache('leaves');
        logAction('UPDATE_LEAVE', leaveId, requestorId);
        return { success: true, message: 'แก้ไขใบลาสำเร็จ' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    }
  };
})();
