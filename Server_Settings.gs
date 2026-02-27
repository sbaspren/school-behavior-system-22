// =================================================================
// Server_Settings.gs - دوال الإعدادات
// الإصدار المصحح - بدون ثوابت خارجية
// =================================================================

// =================================================================
// ★ دالة تنظيف HTML عامة
// =================================================================
function stripHtml_(str) {
  return String(str || '').replace(/<[^>]*>/g, '').trim();
}

// =================================================================
// 1. جلب إعدادات المدرسة (الكليشة)
// =================================================================
function getSchoolSettings() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("إعدادات_المدرسة");

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: { letterhead: '' } };
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};

    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      if (key) {
        settings[key] = value || '';
      }
    }

    return { success: true, data: settings };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 2. حفظ إعدادات المدرسة (الكليشة)
// =================================================================
function saveSchoolSettings(settings) {
  try {
    // ★ التحقق من القيم المسموحة
    var validModes = ['image', 'text'];
    var validWhatsapp = ['per_stage', 'unified'];

    var letterheadMode = validModes.indexOf(settings.letterhead_mode) !== -1 ? settings.letterhead_mode : 'text';
    var whatsappMode = validWhatsapp.indexOf(settings.whatsapp_mode) !== -1 ? settings.whatsapp_mode : 'per_stage';

    // ★ تنظيف النصوص من HTML tags
    var cleanEduAdmin = sanitizeInput_(settings.edu_admin);
    var cleanEduDept = sanitizeInput_(settings.edu_dept);
    var cleanSchoolName = sanitizeInput_(settings.school_name);
    var cleanImageUrl = String(settings.letterhead_image_url || '').trim();

    // ★ التحقق من الروابط — يجب أن تبدأ بـ https:// إن وُجدت
    if (cleanImageUrl && cleanImageUrl.indexOf('https://') !== 0) {
      return { success: false, error: 'رابط صورة الكليشة يجب أن يبدأ بـ https://' };
    }
    // ★ التحقق حسب النوع
    if (letterheadMode === 'image' && !cleanImageUrl) {
      return { success: false, error: 'يرجى إدخال رابط صورة الكليشة' };
    }
    if (letterheadMode === 'text' && !cleanSchoolName) {
      return { success: false, error: 'يرجى إدخال اسم المدرسة على الأقل' };
    }

    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName("إعدادات_المدرسة");

    if (!sheet) {
      sheet = ss.insertSheet("إعدادات_المدرسة");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المفتاح', 'القيمة', 'الوصف', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 4).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }

    const now = new Date();
    const dataToSave = [
      ['letterhead_mode', letterheadMode, 'وضع الكليشة: image أو text', now],
      ['letterhead_image_url', cleanImageUrl, 'رابط صورة الكليشة الكاملة', now],
      ['edu_admin', cleanEduAdmin, 'الإدارة التعليمية', now],
      ['edu_dept', cleanEduDept, 'القسم / المكتب', now],
      ['school_name', cleanSchoolName, 'اسم المدرسة', now],
      ['letterhead', '', 'بيانات الكليشة (نص كامل - توافقية)', now],
      ['whatsapp_mode', whatsappMode, 'نمط الواتساب: per_stage أو unified', now]
    ];

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
    }

    sheet.getRange(2, 1, dataToSave.length, 4).setValues(dataToSave);

    logAuditAction_('تحديث إعدادات', 'تم تحديث إعدادات المدرسة');
    return { success: true, message: 'تم حفظ الإعدادات بنجاح' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 3. جلب هيكل المدرسة
// =================================================================
function getSchoolStructure() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("هيكل_المدرسة");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { 
        success: true, 
        structure: { 
          school_type: 'بنين', 
          secondary_system: 'فصلي', 
          stages: {} 
        },
        isEmpty: true 
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const structure = {
      school_type: 'بنين',
      secondary_system: 'فصلي',
      stages: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (key === 'school_type') {
        structure.school_type = value || 'بنين';
      } else if (key === 'secondary_system') {
        structure.secondary_system = value || 'فصلي';
      } else if (key === 'stage_config') {
        try {
          const stageData = JSON.parse(value);
          if (stageData.stageId && stageData.grades) {
            structure.stages[stageData.stageId] = stageData.grades;
          }
        } catch(e) {}
      }
    }
    
    return { success: true, structure: structure };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 4. حفظ هيكل المدرسة
// =================================================================
function saveSchoolStructure(structure) {
  try {
    // ★ التحقق من القيم المسموحة
    var validSchoolTypes = ['بنين', 'بنات'];
    var validSecondarySystems = ['فصلي', 'مسارات'];
    var validStageIds = ['kindergarten', 'primary', 'intermediate', 'secondary'];

    var schoolType = validSchoolTypes.indexOf(structure.school_type) !== -1 ? structure.school_type : 'بنين';
    var secondarySystem = validSecondarySystems.indexOf(structure.secondary_system) !== -1 ? structure.secondary_system : 'فصلي';
    var confirmedDeletion = structure.confirmedDeletion || false;

    // ★ التحقق من وجود فصل مفعّل واحد على الأقل
    var newEnabledStages = {}; // stageId → arabicName
    if (structure.stages) {
      for (var sid in structure.stages) {
        if (validStageIds.indexOf(sid) === -1) continue;
        var grades = structure.stages[sid];
        var hasEnabled = false;
        for (var g in grades) {
          if (grades[g] && grades[g].enabled && grades[g].classCount > 0) {
            hasEnabled = true;
            if (grades[g].classCount > 15) grades[g].classCount = 15;
          }
        }
        if (hasEnabled) {
          newEnabledStages[sid] = STAGE_ID_TO_ARABIC[sid];
        }
      }
    }
    if (Object.keys(newEnabledStages).length === 0) {
      return { success: false, error: 'يجب تفعيل مرحلة واحدة على الأقل مع صف وفصل' };
    }

    const ss = getSpreadsheet_();

    // ★ قراءة الهيكل القديم لمقارنة المراحل
    var oldEnabledStages = {};
    var oldSheet = ss.getSheetByName('هيكل_المدرسة');
    if (oldSheet && oldSheet.getLastRow() >= 2) {
      var oldData = oldSheet.getDataRange().getValues();
      for (var i = 1; i < oldData.length; i++) {
        if (oldData[i][0] !== 'stage_config') continue;
        try {
          var oldStage = JSON.parse(oldData[i][1]);
          if (!oldStage.stageId || !oldStage.grades) continue;
          var oldArabic = STAGE_ID_TO_ARABIC[oldStage.stageId];
          if (!oldArabic) continue;
          for (var og in oldStage.grades) {
            if (oldStage.grades[og] && oldStage.grades[og].enabled && oldStage.grades[og].classCount > 0) {
              oldEnabledStages[oldStage.stageId] = oldArabic;
              break;
            }
          }
        } catch(e) {}
      }
    }

    // ★ تحديد المراحل المُلغاة والمُضافة
    var removedStages = [];
    for (var oldSid in oldEnabledStages) {
      if (!newEnabledStages[oldSid]) {
        removedStages.push(oldEnabledStages[oldSid]);
      }
    }
    var addedStages = [];
    for (var newSid in newEnabledStages) {
      if (!oldEnabledStages[newSid]) {
        addedStages.push(newEnabledStages[newSid]);
      }
    }

    // ★ إذا هناك مراحل مُلغاة ولم يُؤكّد الحذف — أرجع قائمة للتأكيد
    if (removedStages.length > 0 && !confirmedDeletion) {
      var sheetsToDelete = [];
      removedStages.forEach(function(stage) {
        var studentSheetName = STAGE_ARABIC_TO_SHEET[stage];
        if (studentSheetName) sheetsToDelete.push(studentSheetName);
        var types = Object.keys(SHEET_REGISTRY);
        for (var t = 0; t < types.length; t++) {
          if (SHEET_REGISTRY[types[t]].perStage) {
            sheetsToDelete.push(SHEET_REGISTRY[types[t]].prefix + '_' + stage);
          }
        }
      });
      return {
        success: false,
        needsConfirmation: true,
        removedStages: removedStages,
        sheetsToDelete: sheetsToDelete
      };
    }

    // ★ حذف شيتات المراحل المُلغاة
    if (removedStages.length > 0 && confirmedDeletion) {
      removedStages.forEach(function(stage) {
        // حذف شيت الطلاب
        var studentSheetName = STAGE_ARABIC_TO_SHEET[stage];
        if (studentSheetName) {
          var studentSheet = ss.getSheetByName(studentSheetName);
          if (studentSheet && ss.getSheets().length > 1) {
            ss.deleteSheet(studentSheet);
            Logger.log('🗑️ حذف شيت طلاب: ' + studentSheetName);
          }
        }
        // حذف سجلات per-stage
        var types = Object.keys(SHEET_REGISTRY);
        for (var t = 0; t < types.length; t++) {
          if (!SHEET_REGISTRY[types[t]].perStage) continue;
          var logName = SHEET_REGISTRY[types[t]].prefix + '_' + stage;
          var logSheet = ss.getSheetByName(logName);
          if (logSheet && ss.getSheets().length > 1) {
            ss.deleteSheet(logSheet);
            Logger.log('🗑️ حذف سجل: ' + logName);
          }
          // البحث في الأسماء البديلة أيضاً
          var aliases = SHEET_ALIASES[logName] || [];
          for (var a = 0; a < aliases.length; a++) {
            var aliasSheet = ss.getSheetByName(aliases[a]);
            if (aliasSheet && ss.getSheets().length > 1) {
              ss.deleteSheet(aliasSheet);
              Logger.log('🗑️ حذف سجل (اسم بديل): ' + aliases[a]);
            }
          }
        }
      });
    }

    // ★ حفظ الهيكل الجديد
    let sheet = ss.getSheetByName("هيكل_المدرسة");
    if (!sheet) {
      sheet = ss.insertSheet("هيكل_المدرسة");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المفتاح', 'القيمة', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 3).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }

    const now = new Date();
    const dataToSave = [
      ['school_type', schoolType, now],
      ['secondary_system', secondarySystem, now]
    ];

    if (structure.stages) {
      for (const stageId in structure.stages) {
        if (validStageIds.indexOf(stageId) === -1) continue;
        const stageData = {
          stageId: stageId,
          grades: structure.stages[stageId]
        };
        dataToSave.push(['stage_config', JSON.stringify(stageData), now]);
      }
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clear();
    }

    if (dataToSave.length > 0) {
      sheet.getRange(2, 1, dataToSave.length, 3).setValues(dataToSave);
    }

    // ★ مسح كاش STUDENTS_SHEETS لإجبار إعادة البناء
    resetStudentsSheetsCache_();

    // ★ إنشاء شيتات المراحل المُضافة + سجلاتها
    if (addedStages.length > 0) {
      // إعادة بناء STUDENTS_SHEETS من الهيكل الجديد
      ensureStudentsSheetsLoaded_();
      ensureAllSheets_();
      // مسح كاش sheets_initialized لإجبار الإعادة
      CacheService.getScriptCache().remove('sheets_initialized');
    }

    return {
      success: true,
      message: 'تم حفظ الهيكل بنجاح',
      addedStages: addedStages,
      removedStages: confirmedDeletion ? removedStages : []
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. جلب جميع الطلاب (من شيتات المراحل)
// =================================================================
function getAllStudents() {
  try {
    var sheets = getAllStudentsSheets_();
    var students = [];
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var stage = sheets[s].stage;
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      
      var headers = data[0];
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0] && !row[1]) continue;
        
        var student = {};
        for (var j = 0; j < headers.length; j++) {
          var val = row[j];
          // تحويل التاريخ إلى نص (Date objects لا تُنقل عبر google.script.run)
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy/MM/dd');
          }
          student[headers[j]] = val || '';
        }
        student['المرحلة'] = stage;
        // تنظيف اسم الصف
        if (student['الصف']) student['الصف'] = cleanGradeName_(student['الصف']);
        students.push(student);
      }
    }
    
    return { success: true, students: students };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 6. إضافة طالب يدوياً (يحفظ في شيت المرحلة المناسب)
// =================================================================
function addStudentManually(studentData) {
  try {
    var ss = getSpreadsheet_();
    var grade = cleanGradeName_(studentData.grade || '');
    var stage = studentData.stage || detectStageFromGrade_(grade);
    if (!stage) {
      return { success: false, error: 'يرجى اختيار الصف لتحديد المرحلة' };
    }
    
    // كشف التكرار: البحث في ყველა الشيتات للتأكد من عدم وجود نفس رقم الطالب
    var allSheets = getAllStudentsSheets_();
    for (var s = 0; s < allSheets.length; s++) {
      var currentSheet = allSheets[s].sheet;
      var data = currentSheet.getDataRange().getValues();
      if (data.length < 2) continue;
      
      var headers = data[0];
      var idCol = headers.indexOf('رقم الطالب');
      if (idCol === -1) idCol = headers.indexOf('رقم_الطالب');
      if (idCol === -1) idCol = 0;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]).trim() === String(studentData.id).trim()) {
           return { success: false, error: 'الطالب موجود مسبقاً برقم الهوية ' + studentData.id + ' في مسار ' + allSheets[s].stage };
        }
      }
    }
    
    var sheetName = STUDENTS_SHEETS[stage];
    if (!sheetName) {
      return { success: false, error: 'المرحلة غير مفعّلة: ' + stage };
    }
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.setRightToLeft(true);
      sheet.appendRow(['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'تاريخ الإضافة']);
      sheet.getRange(1, 1, 1, 6).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    sheet.appendRow([
      sanitizeInput_(String(studentData.id)),
      sanitizeInput_(studentData.name),
      grade,
      studentData.class || '',
      sanitizeInput_(studentData.mobile || ''),
      new Date()
    ]);
    
    return { success: true, message: 'تم إضافة الطالب بنجاح' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. حذف طالب (يبحث في جميع شيتات المراحل)
// =================================================================
function deleteStudent(studentId) {
  try {
    var authCheck = checkUserPermission('admin');
    if (!authCheck.hasPermission) {
      return { success: false, error: 'غير مصرح: ' + authCheck.reason };
    }
    var sheets = getAllStudentsSheets_();
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var idCol = headers.indexOf('رقم الطالب');
      if (idCol === -1) idCol = headers.indexOf('رقم_الطالب');
      if (idCol === -1) idCol = 0;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === String(studentId)) {
          sheet.deleteRow(i + 1);
          return { success: true, message: 'تم حذف الطالب بنجاح' };
        }
      }
    }
    
    return { success: false, error: 'الطالب غير موجود' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 8. جلب المعلمين والمواد
// =================================================================
function getTeachers() {
  try {
    const ss = getSpreadsheet_();
    
    const teachersSheet = ss.getSheetByName("المعلمين");
    let teachers = [];
    
    if (teachersSheet && teachersSheet.getLastRow() > 1) {
      const data = teachersSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue;
        
        var civilId = String(row[0]).trim(); // العمود A = المعرف = السجل المدني
        
        teachers.push({
          id: civilId,
          civil_id: civilId,
          name: String(row[2] || ''),
          mobile: String(row[3] || ''),
          subjects: row[4] ? String(row[4]).split(',').map(s => s.trim()) : [],
          assigned_classes: row[5] ? String(row[5]).split(',').map(s => s.trim()) : [],
          status: String(row[7] || 'active')
        });
      }
    }
    
    const subjectsSheet = ss.getSheetByName("المواد");
    let subjects = [];
    
    if (subjectsSheet && subjectsSheet.getLastRow() > 1) {
      const data = subjectsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          subjects.push({
            id: String(data[i][0]),
            name: String(data[i][1] || '')
          });
        }
      }
    }
    
    return { success: true, teachers: teachers, subjects: subjects };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 9. إضافة معلم (المعرف = السجل المدني)
// =================================================================
function addTeacher(teacherData) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName("المعلمين");
    
    if (!sheet) {
      sheet = ss.insertSheet("المعلمين");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المعرف', 'السجل المدني', 'الاسم', 'الجوال', 'المواد', 'الفصول المسندة', 'الصلاحيات', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 10).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    var civilId = String(teacherData.civil_id || '').trim();
    if (!civilId) return { success: false, error: 'السجل المدني مطلوب' };
    
    // كشف التكرار: المعرف = السجل المدني
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
      for (var i = 1; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === civilId) {
          return { success: false, error: 'المعلم موجود مسبقاً (السجل المدني: ' + civilId + ')' };
        }
      }
    }
    
    const now = new Date();
    
    sheet.appendRow([
      civilId,
      civilId,
      sanitizeInput_(teacherData.name),
      sanitizeInput_(teacherData.mobile || ''),
      (teacherData.subjects || []).map(function(s) { return sanitizeInput_(s); }).join(','),
      (teacherData.assigned_classes || []).join(','),
      '',
      'active',
      now,
      now
    ]);

    return { success: true, message: 'تم إضافة المعلم بنجاح', id: civilId };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 10. تحديث معلم (البحث بالسجل المدني = المعرف)
// =================================================================
function updateTeacher(teacherData) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المعلمين");
    
    if (!sheet) {
      return { success: false, error: 'شيت المعلمين غير موجود' };
    }
    
    var civilId = String(teacherData.id || teacherData.civil_id || '').trim();
    if (!civilId) return { success: false, error: 'السجل المدني مطلوب' };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === civilId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'المعلم غير موجود' };
    }
    
    const now = new Date();
    sheet.getRange(rowIndex, 1, 1, 10).setValues([[
      civilId,
      civilId,
      sanitizeInput_(teacherData.name),
      sanitizeInput_(teacherData.mobile || ''),
      (teacherData.subjects || []).map(function(s) { return sanitizeInput_(s); }).join(','),
      (teacherData.assigned_classes || []).join(','),
      '',
      'active',
      data[rowIndex - 1][8],
      now
    ]]);

    return { success: true, message: 'تم تحديث المعلم بنجاح' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 11. حذف معلم (البحث بالسجل المدني = المعرف)
// =================================================================
function deleteTeacher(teacherId) {
  try {
    var authCheck = checkUserPermission('admin');
    if (!authCheck.hasPermission) {
      return { success: false, error: 'غير مصرح: ' + authCheck.reason };
    }
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المعلمين");
    
    if (!sheet) {
      return { success: false, error: 'شيت المعلمين غير موجود' };
    }
    
    var civilId = String(teacherId || '').trim();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === civilId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف المعلم بنجاح' };
      }
    }
    
    return { success: false, error: 'المعلم غير موجود' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 12. إضافة مادة
// =================================================================
function addSubject(subjectData) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName("المواد");
    
    if (!sheet) {
      sheet = ss.insertSheet("المواد");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المعرف', 'اسم المادة', 'تاريخ الإنشاء']);
      sheet.getRange(1, 1, 1, 3).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    const now = new Date();
    const id = 'SUB_' + now.getTime();
    
    sheet.appendRow([id, subjectData.name, now]);
    
    return { success: true, message: 'تم إضافة المادة بنجاح', id: id };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 13. حذف مادة
// =================================================================
function deleteSubject(subjectId) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المواد");
    
    if (!sheet) {
      return { success: false, error: 'شيت المواد غير موجود' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(subjectId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف المادة بنجاح' };
      }
    }
    
    return { success: false, error: 'المادة غير موجودة' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 14. معالجة ملف Excel المرفوع
// =================================================================
function processUploadedFile(base64Data, fileName, importType) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
    var tempFile = DriveApp.createFile(blob);
    var tempSS = SpreadsheetApp.open(tempFile);
    
    // للطلاب: البيانات في Sheet2 (الورقة الثانية)
    // للمعلمين: البيانات في الورقة الأولى
    var tempSheet;
    if (importType === 'students' && tempSS.getSheets().length > 1) {
      tempSheet = tempSS.getSheets()[1]; // Sheet2
    } else {
      tempSheet = tempSS.getSheets()[0];
    }
    
    var allData = tempSheet.getDataRange().getValues();
    
    // البحث عن صف العناوين (قد لا يكون الصف الأول)
    var headerRow = 0;
    for (var r = 0; r < Math.min(10, allData.length); r++) {
      var rowStr = allData[r].join(' ').toLowerCase();
      if (rowStr.includes('اسم') || rowStr.includes('name') || rowStr.includes('طالب')) {
        headerRow = r;
        break;
      }
    }
    
    var headers = allData[headerRow] || [];
    // البيانات تبدأ بعد صف العناوين
    var data = allData.slice(headerRow);
    
    var columns = {};
    
    headers.forEach(function(header, index) {
      var h = String(header || '').trim();
      
      if (importType === 'teachers') {
        if (h.includes('سجل') || h.includes('هوية') || h.includes('civil')) columns.civil_id = index;
        if (h.includes('اسم') || h.includes('name')) columns.name = index;
        if (h.includes('جوال') || h.includes('هاتف') || h.includes('mobile')) columns.mobile = index;
      } else {
        if (h.includes('رقم') && (h.includes('طالب') || h.includes('هوية'))) columns.studentId = index;
        if (h.includes('اسم')) columns.name = index;
        if (h.includes('رقم') && h.includes('صف')) columns.grade = index;
        if (h === 'الصف') columns.grade = index;
        if (h === 'الفصل') columns.classVal = index;
        if (h.includes('جوال') || h.includes('هاتف')) columns.mobile = index;
      }
    });
    
    tempFile.setTrashed(true);
    
    return {
      success: true,
      headers: headers,
      preview: data.slice(0, 10),
      totalRows: data.length - 1,
      columns: columns,
      rawData: data
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 15. استيراد المعلمين
// =================================================================
function importTeachersWithMapping(fileData, mapping) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName("المعلمين");
    
    if (!sheet) {
      sheet = ss.insertSheet("المعلمين");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المعرف', 'السجل المدني', 'الاسم', 'الجوال', 'المواد', 'الفصول المسندة', 'الصلاحيات', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 10).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    const rawData = fileData.rawData || [];
    const now = new Date();
    let imported = 0;
    
    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      const civil_id = mapping.civil_id >= 0 ? String(row[mapping.civil_id] || '').trim() : '';
      const name = mapping.name >= 0 ? String(row[mapping.name] || '').trim() : '';
      const mobile = mapping.mobile >= 0 ? String(row[mapping.mobile] || '').trim() : '';
      
      if (!civil_id || !name) continue;
      
      // المعرف = السجل المدني — تنظيف النصوص
      sheet.appendRow([civil_id, civil_id, sanitizeInput_(name), sanitizeInput_(mobile), '', '', '', 'active', now, now]);
      imported++;
    }

    return { success: true, message: 'تم استيراد ' + imported + ' معلم بنجاح' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 16. استيراد الطلاب - قراءة مباشرة من xlsx بدون SpreadsheetApp
// =================================================================
function importStudentsFromExcel(base64Data, fileName) {
  try {
    // 1. فك ضغط xlsx (هو ملف zip)
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/zip', 'temp.zip');
    var files = Utilities.unzip(blob);
    
    // 2. قراءة shared strings
    var shared = [];
    for (var f = 0; f < files.length; f++) {
      if (files[f].getName().indexOf('sharedStrings') > -1) {
        var ssXml = XmlService.parse(files[f].getDataAsString());
        var ns = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        var siElements = ssXml.getRootElement().getChildren('si', ns);
        for (var s = 0; s < siElements.length; s++) {
          var texts = siElements[s].getDescendants();
          var fullText = '';
          for (var t = 0; t < texts.length; t++) {
            if (texts[t].getType() === XmlService.ContentTypes.TEXT) {
              fullText += texts[t].getValue();
            }
          }
          shared.push(fullText);
        }
        break;
      }
    }
    
    // 3. قراءة Sheet2 (بيانات الطلاب)
    var sheetData = null;
    for (var f = 0; f < files.length; f++) {
      if (files[f].getName().indexOf('sheet2') > -1) {
        sheetData = files[f];
        break;
      }
    }
    
    if (!sheetData) {
      // fallback: Sheet1
      for (var f = 0; f < files.length; f++) {
        if (files[f].getName().indexOf('sheet1') > -1) {
          sheetData = files[f];
          break;
        }
      }
    }
    
    if (!sheetData) {
      return { success: false, error: 'لم يتم العثور على ورقة البيانات' };
    }
    
    // 4. تحليل بيانات الورقة
    var sheetXml = XmlService.parse(sheetData.getDataAsString());
    var ns = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    var rows = sheetXml.getRootElement().getChild('sheetData', ns).getChildren('row', ns);
    
    // 5. تحويل إلى مصفوفة
    var allData = [];
    for (var r = 0; r < rows.length; r++) {
      var cells = rows[r].getChildren('c', ns);
      var rowData = [];
      
      for (var c = 0; c < cells.length; c++) {
        var cell = cells[c];
        var ref = cell.getAttribute('r').getValue(); // مثل B5, C5
        var colLetter = ref.replace(/[0-9]/g, '');
        var colIndex = colLetterToIndex_(colLetter);
        
        // ملء الخلايا الفارغة
        while (rowData.length < colIndex) rowData.push('');
        
        var valEl = cell.getChild('v', ns);
        var val = valEl ? valEl.getText() : '';
        var typAttr = cell.getAttribute('t');
        var typ = typAttr ? typAttr.getValue() : '';
        
        if (typ === 's' && val) {
          var idx = parseInt(val);
          val = (idx < shared.length) ? shared[idx] : val;
        }
        
        rowData.push(val);
      }
      
      allData.push(rowData);
    }
    
    // 6. البحث عن صف العناوين
    var headerRow = -1;
    for (var r = 0; r < Math.min(10, allData.length); r++) {
      var rowStr = allData[r].join(' ');
      if (rowStr.indexOf('اسم') > -1 && (rowStr.indexOf('طالب') > -1 || rowStr.indexOf('جوال') > -1)) {
        headerRow = r;
        break;
      }
    }
    
    if (headerRow === -1) {
      return { success: false, error: 'لم يتم العثور على صف العناوين (اسم الطالب، الجوال...)' };
    }
    
    var headers = allData[headerRow];
    
    // 7. اكتشاف الأعمدة
    var colMap = { studentId: -1, name: -1, grade: -1, classVal: -1, mobile: -1 };
    
    for (var h = 0; h < headers.length; h++) {
      var header = String(headers[h] || '').trim();
      if (header.indexOf('رقم') > -1 && header.indexOf('طالب') > -1) colMap.studentId = h;
      else if (header.indexOf('اسم') > -1) colMap.name = h;
      else if (header.indexOf('رقم') > -1 && header.indexOf('صف') > -1) colMap.grade = h;
      else if (header === 'الصف') colMap.grade = h;
      else if (header === 'الفصل') colMap.classVal = h;
      else if (header.indexOf('جوال') > -1 || header.indexOf('هاتف') > -1) colMap.mobile = h;
    }
    
    // fallback
    if (colMap.name === -1) colMap.name = 4;
    if (colMap.studentId === -1) colMap.studentId = 5;
    if (colMap.grade === -1) colMap.grade = 3;
    if (colMap.classVal === -1) colMap.classVal = 2;
    if (colMap.mobile === -1) colMap.mobile = 1;
    
    // 8. تجميع الطلاب حسب المرحلة
    var ss = getSpreadsheet_();
    var byStage = {};
    var now = new Date();
    
    for (var i = headerRow + 1; i < allData.length; i++) {
      var row = allData[i];
      var name = colMap.name < row.length ? String(row[colMap.name] || '').trim() : '';
      var studentId = colMap.studentId < row.length ? String(row[colMap.studentId] || '').trim() : '';
      var gradeRaw = colMap.grade < row.length ? String(row[colMap.grade] || '').trim() : '';
      var classVal = colMap.classVal < row.length ? String(row[colMap.classVal] || '').trim() : '';
      var mobile = colMap.mobile < row.length ? String(row[colMap.mobile] || '').trim() : '';
      
      if (!name || name.indexOf('اسم') > -1) continue;
      
      var grade = cleanGradeName_(gradeRaw);
      var stage = detectStageFromGrade_(gradeRaw);
      if (!stage) continue;
      
      if (!byStage[stage]) byStage[stage] = [];
      byStage[stage].push([studentId, sanitizeInput_(name), grade, classVal, sanitizeInput_(mobile), now]);
    }
    
    // 9. حفظ في الشيتات المناسبة (مع منع التكرار)
    var totalNew = 0;
    var totalUpdated = 0;
    var totalSkipped = 0;
    var stagesList = Object.keys(byStage);
    
    if (stagesList.length === 0) {
      return { success: false, error: 'لم يتم العثور على طلاب في الملف' };
    }
    
    for (var s = 0; s < stagesList.length; s++) {
      var stage = stagesList[s];
      var sheetName = STUDENTS_SHEETS[stage];
      if (!sheetName) continue;
      
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.setRightToLeft(true);
        sheet.appendRow(['رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'رقم الجوال', 'تاريخ الإضافة']);
        sheet.getRange(1, 1, 1, 6).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
      
      // ★ قراءة أرقام الطلاب الموجودين حالياً
      var existingIds = {};
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
        for (var e = 0; e < existingData.length; e++) {
          var existId = String(existingData[e][0] || '').trim();
          if (existId) {
            existingIds[existId] = e + 2; // رقم الصف (1-indexed, +1 للعنوان)
          }
        }
      }
      
      // ★ فصل الطلاب الجدد عن الموجودين
      var newRows = [];
      var importRows = byStage[stage];
      
      for (var r = 0; r < importRows.length; r++) {
        var studentId = String(importRows[r][0] || '').trim();
        
        if (studentId && existingIds[studentId]) {
          // ★ الطالب موجود - تحديث بياناته (الاسم، الصف، الفصل، الجوال)
          var rowNum = existingIds[studentId];
          sheet.getRange(rowNum, 2, 1, 4).setValues([[importRows[r][1], importRows[r][2], importRows[r][3], importRows[r][4]]]);
          totalUpdated++;
        } else if (studentId) {
          // ★ طالب جديد - إضافة
          newRows.push(importRows[r]);
          totalNew++;
        } else {
          totalSkipped++;
        }
      }
      
      // ★ كتابة الطلاب الجدد فقط دفعة واحدة
      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
      }
    }
    
    // ★ رسالة تفصيلية
    var msg = '';
    if (totalNew > 0) msg += 'تم إضافة ' + totalNew + ' طالب جديد';
    if (totalUpdated > 0) msg += (msg ? '، و' : '') + 'تحديث بيانات ' + totalUpdated + ' طالب';
    if (totalNew === 0 && totalUpdated === 0) msg = 'جميع الطلاب موجودون مسبقاً، لا توجد إضافات جديدة';
    msg += ' (' + stagesList.join(' و ') + ')';
    
    return { success: true, message: msg, added: totalNew, updated: totalUpdated, skipped: totalSkipped };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// دالة تحويل حرف العمود إلى رقم (A=0, B=1, ..., Z=25, AA=26)
function colLetterToIndex_(letter) {
  var result = 0;
  for (var i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

// === التوافق مع الدوال القديمة ===
function importStudentsAuto(fileData) {
  return { success: false, error: 'استخدم importStudentsFromExcel' };
}
function importStudentsWithMapping(fileData, mapping) {
  return { success: false, error: 'استخدم importStudentsFromExcel' };
}

// =================================================================
// 20. حفظ معلمين محللين من المتصفح (SheetJS) — المعرف = السجل المدني
// =================================================================
function saveTeachersParsedData(jsonStr, updateExisting) {
  try {
    var teachers = JSON.parse(jsonStr);
    if (!teachers || teachers.length === 0) {
      return { success: false, error: 'لا توجد بيانات' };
    }
    
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("المعلمين");
    
    if (!sheet) {
      sheet = ss.insertSheet("المعلمين");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المعرف', 'السجل المدني', 'الاسم', 'الجوال', 'المواد', 'الفصول المسندة', 'الصلاحيات', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 10).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    // قراءة البيانات الحالية لكشف التكرار (العمود A = السجل المدني = المعرف)
    var existingData = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [[]];
    var civilIdMap = {}; // civil_id → row index (1-based)
    for (var i = 1; i < existingData.length; i++) {
      var cid = String(existingData[i][0] || '').trim();
      if (cid) civilIdMap[cid] = i + 1;
    }
    
    var now = new Date();
    var addedCount = 0;
    var updatedCount = 0;
    var newRows = [];
    
    for (var t = 0; t < teachers.length; t++) {
      var teacher = teachers[t];
      var identity = String(teacher.identity || '').trim();
      var name = String(teacher.name || '').trim();
      if (!identity || !name) continue;
      
      var mobile = String(teacher.mobile || '').trim();
      
      if (civilIdMap[identity]) {
        // معلم موجود — تحديث الجوال فقط
        if (updateExisting) {
          var rowNum = civilIdMap[identity];
          if (mobile) sheet.getRange(rowNum, 4).setValue(sanitizeInput_(mobile));
          sheet.getRange(rowNum, 10).setValue(now);
          updatedCount++;
        }
      } else {
        // معلم جديد — المعرف = السجل المدني — تنظيف النصوص
        newRows.push([identity, identity, sanitizeInput_(name), sanitizeInput_(mobile), '', '', '', 'active', now, now]);
        civilIdMap[identity] = true; // منع التكرار داخل نفس الملف
        addedCount++;
      }
    }
    
    // كتابة الجدد دفعة واحدة
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
    }
    
    var parts = [];
    if (addedCount > 0) parts.push('تم إضافة ' + addedCount + ' معلم جديد');
    if (updatedCount > 0) parts.push('تم تحديث ' + updatedCount + ' معلم موجود');
    
    return {
      success: true,
      added: addedCount,
      updated: updatedCount,
      message: parts.join(' | ') || 'لم يتم إجراء تغييرات'
    };
    
  } catch (e) {
    Logger.log('خطأ حفظ المعلمين: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}