var TimetableService = (function () {
  return {

    getScheduleByClass: function (className) {
      try {
        var result = DataService.getSchedules();
        if (!result.success) return result;
        var filtered = (result.data || []).filter(function(s) {
          return s.class_name === className;
        });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    // ตาราง ปตน.2-69 รวมแหล่งข้อมูลตามห้องที่แสดง โดยคงรายการที่ชนคาบไว้ครบ
    getPotnScheduleByClass: function (displayClass) {
      try {
        var aliases = {
          'ม1/1': ['ม1/1'],
          'ม1/2': ['ม1/2', 'ม1/2-3'],
          'ม1/3': ['ม1/3', 'ม1/2-3'],
          'ม2/1': ['ม2/1'],
          'ม2/2': ['ม2/2', 'ม2/2-3'],
          'ม2/3': ['ม2/3', 'ม2/2-3'],
          'ม3/1': ['ม3/1'],
          'ม3/2': ['ม3/2', 'ม3/2-3'],
          'ม3/3': ['ม3/3', 'ม3/2-3'],
          'ม4/1': ['ม4/1', 'ม4/1-3'],
          'ม4/2': ['ม4/2', 'ม4/2-3', 'ม4/1-3'],
          'ม4/3': ['ม4/3', 'ม4/2-3', 'ม4/1-3'],
          'ม5/1': ['ม5/1', 'ม5/1-3'],
          'ม5/2': ['ม5/2', 'ม5/2-3', 'ม5/1-3'],
          'ม5/3': ['ม5/3', 'ม5/2-3', 'ม5/1-3'],
          'ม6/1': ['ม6/1', 'ม6/1-3'],
          'ม6/2': ['ม6/2', 'ม6/2-3', 'ม6/1-3'],
          'ม6/3': ['ม6/3', 'ม6/2-3', 'ม6/1-3']
        };
        var key = String(displayClass || '').replace(/[.\s]/g, '');
        var sourceClasses = aliases[key];
        if (!sourceClasses) return { success: false, message: 'ไม่รองรับชั้นเรียนนี้', data: [] };

        var result = DataService.getSchedules();
        if (!result.success) return result;
        var sourceSet = {};
        sourceClasses.forEach(function (name) { sourceSet[name] = true; });
        var filtered = (result.data || []).filter(function (s) {
          var sourceName = String(s.class_name || '').replace(/[.\s]/g, '');
          return sourceSet[sourceName];
        });
        return { success: true, data: filtered };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    getScheduleByTeacher: function (teacherId) {
      try {
        return DataService.getSchedulesByTeacher(teacherId);
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    },

    /**
     * ⚡ Batch Save — บันทึกตารางเรียนทั้งห้อง
     * ลบข้อมูลเก่าของห้องนั้น แล้วแทรกใหม่ทั้งหมด (atomic)
     */
    saveScheduleBatch: function (data, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var className = data.class_name;
        var entries = data.entries || [];
        if (!className) return { success: false, message: 'ไม่ระบุห้องเรียน' };

        // ตรวจ conflicts ก่อนบันทึก
        var conflictResult = this.checkConflicts({ class_name: className, entries: entries });
        if (conflictResult.hasConflict) {
          return {
            success: false,
            message: 'พบการซ้ำซ้อน: ' + conflictResult.conflicts.map(function(c) { return c.message; }).join(', '),
            conflicts: conflictResult.conflicts
          };
        }

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SCHEDULES_SHEET);
        var allData = sheet.getDataRange().getValues();
        var map = _buildHeaderIndex(allData[0]);

        // ลบแถวเก่าของห้องนี้ (จากล่างขึ้นบน เพื่อไม่ให้ row shift)
        var rowsToDelete = [];
        for (var i = 1; i < allData.length; i++) {
          if (allData[i][map.class_name] === className) {
            rowsToDelete.push(i + 1);
          }
        }
        for (var d = rowsToDelete.length - 1; d >= 0; d--) {
          sheet.deleteRow(rowsToDelete[d]);
        }

        // แทรกข้อมูลใหม่
        var newRows = [];
        for (var e = 0; e < entries.length; e++) {
          var entry = entries[e];
          if (!entry.subject_id || !entry.teacher_id) continue;
          newRows.push([
            generateId('SCH'),
            entry.day_of_week,
            entry.period,
            entry.subject_id,
            className,
            entry.teacher_id,
            entry.room || ''
          ]);
        }
        if (newRows.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
        }

        _invalidateCache('schedules');
        logAction('SAVE_TIMETABLE', className + ' (' + newRows.length + ' entries)', requestorId);
        return { success: true, message: 'บันทึกตารางเรียน ' + className + ' สำเร็จ (' + newRows.length + ' รายการ)' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    /**
     * ⚡ Conflict Detection Engine
     * ตรวจ:
     * 1. ครูสอนซ้ำในคาบเดียวกัน (ต่างห้อง)
     * 2. ครูสอนซ้ำในห้องเดียวกัน (double booking)
     */
    checkConflicts: function (data) {
      try {
        var className = data.class_name;
        var entries = data.entries || [];
        var schedulesResult = DataService.getSchedules();
        var existingSchedules = (schedulesResult.data || []).filter(function(s) {
          return s.class_name !== className; // exclude current class
        });

        // สร้าง Teacher Name Map
        var teachersResult = DataService.getTeachers();
        var teacherNameMap = {};
        (teachersResult.data || []).forEach(function(t) {
          teacherNameMap[t.teacher_id] = (t.prefix || '') + t.first_name + ' ' + t.last_name;
        });

        // สร้าง Busy Map จากตารางที่มีอยู่
        var teacherBusy = {}; // key: "day_period_teacherId"
        var roomBusy = {};    // key: "day_period_room"
        for (var i = 0; i < existingSchedules.length; i++) {
          var s = existingSchedules[i];
          var key = s.day_of_week + '_' + s.period + '_' + s.teacher_id;
          teacherBusy[key] = { class_name: s.class_name, subject: s.subject_name };
          if (s.room) {
            var roomKey = s.day_of_week + '_' + s.period + '_' + s.room;
            roomBusy[roomKey] = { class_name: s.class_name, subject: s.subject_name };
          }
        }

        var conflicts = [];
        for (var e = 0; e < entries.length; e++) {
          var entry = entries[e];
          if (!entry.teacher_id) continue;
          var tName = teacherNameMap[entry.teacher_id] || entry.teacher_id;
          // ตรวจครูซ้ำ
          var busyKey = entry.day_of_week + '_' + entry.period + '_' + entry.teacher_id;
          if (teacherBusy[busyKey]) {
            var existing = teacherBusy[busyKey];
            conflicts.push({
              type: 'teacher_conflict',
              teacher_id: entry.teacher_id,
              day: entry.day_of_week,
              period: entry.period,
              message: tName + ' สอนซ้ำในคาบ ' + entry.period +
                       ' วัน' + entry.day_of_week + ' (สอน ' + existing.class_name + ' อยู่แล้ว)'
            });
          }
          // ตรวจห้องซ้ำ
          if (entry.room) {
            var entryRoomKey = entry.day_of_week + '_' + entry.period + '_' + entry.room;
            if (roomBusy[entryRoomKey]) {
              var existRoom = roomBusy[entryRoomKey];
              conflicts.push({
                type: 'room_conflict',
                room: entry.room,
                day: entry.day_of_week,
                period: entry.period,
                message: 'ห้อง ' + entry.room + ' ถูกใช้ในคาบ ' + entry.period +
                         ' วัน' + entry.day_of_week + ' (โดย ' + existRoom.class_name + ')'
              });
            }
            roomBusy[entryRoomKey] = { class_name: className, subject: entry.subject_id };
          }
          // เพิ่ม entry ปัจจุบันใน busy map
          teacherBusy[busyKey] = { class_name: className, subject: entry.subject_id };
        }

        return { hasConflict: conflicts.length > 0, conflicts: conflicts };
      } catch (e) {
        return { hasConflict: false, conflicts: [], error: e.message };
      }
    },

    clearClassSchedule: function (className, requestorId) {
      validateRequestorRole_(requestorId, ['Admin', 'Scheduler']);
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SCHEDULES_SHEET);
        var data = sheet.getDataRange().getValues();
        var map = _buildHeaderIndex(data[0]);
        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
          if (data[i][map.class_name] === className) rowsToDelete.push(i + 1);
        }
        for (var d = rowsToDelete.length - 1; d >= 0; d--) {
          sheet.deleteRow(rowsToDelete[d]);
        }
        _invalidateCache('schedules');
        logAction('CLEAR_TIMETABLE', className, requestorId);
        return { success: true, message: 'ล้างตารางเรียน ' + className + ' สำเร็จ (' + rowsToDelete.length + ' รายการ)' };
      } catch (e) {
        return { success: false, message: e.message };
      } finally { lock.releaseLock(); }
    },

    /**
     * สร้างรายชื่อห้องเรียนทั้งหมดจาก Subjects.class_names
     */
    getAllClassNames: function () {
      try {
        var result = DataService.getSubjects();
        var subjects = result.data || [];
        var classSet = {};
        for (var i = 0; i < subjects.length; i++) {
          var names = subjects[i].class_names;
          if (typeof names === 'string') {
            try { names = JSON.parse(names); } catch (e) { names = []; }
          }
          if (Array.isArray(names)) {
            names.forEach(function(cn) { if (cn) classSet[cn] = true; });
          }
        }
        var classList = Object.keys(classSet);
        classList.sort();
        return { success: true, data: classList };
      } catch (e) {
        return { success: false, message: e.message, data: [] };
      }
    }
  };
})();
