// =================================================================
// Server_StaffInput.gs - نظام إدخال الهيئة الإدارية
// (التأخر والاستئذان) مع نظام الروابط الفريدة
// =================================================================

// ★ SPREADSHEET_URL مُعرّف في Config.js — لا تكرّره هنا

// ★ ترتيب الصفوف (يُستخدم لفرز الفصول ديناميكياً)
var GRADE_ORDER_MAP = { 'الأول': 1, 'الثاني': 2, 'الثالث': 3, 'الرابع': 4, 'الخامس': 5, 'السادس': 6 };
var LETTER_ORDER_MAP = { 'أ': 1, 'ب': 2, 'ج': 3, 'د': 4, 'هـ': 5, 'و': 6, 'ز': 7, 'ح': 8, 'ط': 9, 'ي': 10 };

function sortClassNames_(a, b) {
  // استخراج الرتبة من اسم الفصل
  var getOrder = function(name) {
    var gradeOrder = 99, letterOrder = 99;
    for (var g in GRADE_ORDER_MAP) {
      if (name.indexOf(g) >= 0) { gradeOrder = GRADE_ORDER_MAP[g]; break; }
    }
    for (var l in LETTER_ORDER_MAP) {
      if (name.indexOf(l) >= 0) { letterOrder = LETTER_ORDER_MAP[l]; break; }
    }
    return gradeOrder * 100 + letterOrder;
  };
  return getOrder(a) - getOrder(b);
}

// =================================================================
// 1. توليد رمز فريد للموظف
// =================================================================
function generateStaffToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var token = '';
  for (var i = 0; i < 8; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'S' + token; // S للموظفين
}

// ★ createStaffLink + createGuardLink — حُذفتا (كود ميت — الربط يتم عبر sendLinkToPersonWithStage → createUserLink)
// =================================================================
// 4. التحقق من رمز الموظف
// =================================================================
function getStaffByToken(token) {
  try {
    if (!token) {
      return { success: false, error: 'الرمز غير موجود' };
    }
    
    var ss = getSpreadsheet_();
    // ★ إصلاح: التوكن يُحفظ في "المستخدمين" (عبر createUserLink)
    // لذلك نبحث فيه أولاً، ثم "الهيئة_الإدارية" كبديل
    var staffSheet = ss.getSheetByName("المستخدمين") || ss.getSheetByName("الهيئة_الإدارية");
    
    if (!staffSheet) {
      return { success: false, error: 'شيت الهيئة الإدارية غير موجود' };
    }
    
    var data = staffSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idCol = findColumnIndex_(headers, ['المعرف', 'الرقم', 'id']);
    var nameCol = findColumnIndex_(headers, ['الاسم', 'اسم الموظف', 'name']);
    var roleCol = findColumnIndex_(headers, ['الدور', 'المسمى', 'الوظيفة', 'role']);
    var permissionsCol = findColumnIndex_(headers, ['الصلاحيات', 'permissions']);
    var tokenCol = findColumnIndex_(headers, ['رمز_الرابط', 'الرمز', 'token']);
    
    if (tokenCol === -1) {
      return { success: false, error: 'عمود الرمز غير موجود' };
    }
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][tokenCol]).trim() === String(token).trim()) {
        var permissions = data[i][permissionsCol] ? 
          String(data[i][permissionsCol]).split(',').map(function(p) { return p.trim(); }) : [];
        
        return {
          success: true,
          staff: {
            id: String(data[i][idCol] || ''),
            name: String(data[i][nameCol] || ''),
            role: String(data[i][roleCol] || ''),
            permissions: permissions,
            isGuard: String(data[i][roleCol] || '').includes('حارس')
          }
        };
      }
    }
    
    return { success: false, error: 'رابط غير صالح أو منتهي الصلاحية' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. جلب بيانات الموظف للنموذج
// =================================================================
function getStaffFormData(token) {
  return getStaffByToken(token);
}

// =================================================================
// 6. جلب جميع الطلاب (للهيئة الإدارية) - من شيتات المراحل
// =================================================================
function getAllStudentsForStaff() {
  try {
    var sheets = getAllStudentsSheets_();
    var students = {};

    // ★ بناء ديناميكي من المراحل المفعّلة
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var stage = sheets[s].stage;
      if (!students[stage]) students[stage] = {};

      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = data[0];

      var idCol = findColumnIndex_(headers, ['رقم الطالب', 'المعرف', 'الرقم']);
      var nameCol = findColumnIndex_(headers, ['اسم الطالب', 'الاسم']);
      var classCol = findColumnIndex_(headers, ['الفصل', 'فصل']);
      var gradeCol = findColumnIndex_(headers, ['الصف', 'رقم الصف', 'المرحلة']);
      var phoneCol = findColumnIndex_(headers, ['جوال ولي الأمر', 'الجوال', 'رقم الجوال']);

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var studentId = idCol > -1 ? String(row[idCol] || '').trim() : '';
        var studentName = nameCol > -1 ? String(row[nameCol] || '').trim() : '';
        var grade = gradeCol > -1 ? cleanGradeName_(row[gradeCol]) : '';
        var className = classCol > -1 ? String(row[classCol] || '').trim() : '';
        var phone = phoneCol > -1 ? String(row[phoneCol] || '').trim() : '';

        if (!studentName) continue;

        var fullClassName = (grade + ' ' + className).trim();

        if (!students[stage][fullClassName]) {
          students[stage][fullClassName] = [];
        }

        students[stage][fullClassName].push({
          id: studentId,
          name: studentName,
          phone: phone || 'غير متوفر'
        });
      }
    }

    // ★ ترتيب الفصول ديناميكياً
    var sortedStudents = {};
    var stageKeys = Object.keys(students);
    stageKeys.forEach(function(stage) {
      sortedStudents[stage] = {};
      var classes = Object.keys(students[stage]);
      classes.sort(sortClassNames_);
      classes.forEach(function(cls) {
        sortedStudents[stage][cls] = students[stage][cls];
      });
    });

    return { success: true, students: sortedStudents };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. حفظ الاستئذان
// =================================================================
function saveStaffPermission(data) {
  try {
    if (!data.stage) {
      return { success: false, error: 'المرحلة مطلوبة' };
    }
    var stage = data.stage;
    // استخدام نفس شيت الاستئذان الموحد (سجل_الاستئذان_[المرحلة])
    var sheet = getPermissionSheet(stage);
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    
    var savedCount = 0;
    
    if (data.students && data.students.length > 0) {
      // ★ فصل الصف والفصل تلقائياً
      var parsed = parseClassName_(data.className, data.gradeName, data.classLetter);
      var gradeName = sanitizeInput_(parsed.grade);
      var classLetter = sanitizeInput_(parsed.section);
      var safeReason = sanitizeInput_(data.reason || '');
      var safeGuardian = sanitizeInput_(data.guardian || '');
      var safeResponsible = sanitizeInput_(data.responsible || data.staffName || 'إداري');
      var safeStaffName = sanitizeInput_(data.staffName || 'إداري');

      // ★ تجميع الصفوف دفعة واحدة (بدلاً من appendRow في حلقة)
      var rows = [];
      data.students.forEach(function(student) {
        var studentName = sanitizeInput_(student.name || (typeof student === 'string' ? student : ''));
        var studentId = sanitizeInput_(student.id || '');
        var phone = sanitizeInput_(student.phone || '');

        rows.push([
          studentId, studentName, gradeName, classLetter, phone,
          timeStr, safeReason, safeGuardian, safeResponsible,
          hijriDate, safeStaffName, now, '', 'لا'
        ]);
        savedCount++;
      });

      // ★ كتابة دفعة واحدة — أسرع بكثير من appendRow × N
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }

      // تسجيل النشاط
      logStaffActivity_(data.staffName, 'استئذان', (gradeName + ' ' + classLetter).trim(), savedCount, stage);
    }
    
    return { 
      success: true, 
      message: 'تم تسجيل ' + savedCount + ' استئذان بنجاح',
      count: savedCount
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 8. حفظ التأخر
// =================================================================
function saveStaffLatecomers(data) {
  try {
    if (!data.stage) {
      return { success: false, error: 'المرحلة مطلوبة' };
    }
    var stage = data.stage;
    // استخدام نفس شيت التأخر الموحد (سجل_التأخر_[المرحلة])
    var sheet = getLateSheet(stage);
    
    var now = new Date();
    var hijriDate = getHijriDate_(now);
    
    var savedCount = 0;
    
    if (data.students && data.students.length > 0) {
      // ★ فصل الصف والفصل تلقائياً
      var parsed = parseClassName_(data.className, data.gradeName, data.classLetter);
      var gradeName = sanitizeInput_(parsed.grade);
      var classLetter = sanitizeInput_(parsed.section);
      var safeStaffName = sanitizeInput_(data.staffName || 'إداري');

      // ★ تجميع الصفوف دفعة واحدة
      var rows = [];
      data.students.forEach(function(student) {
        var studentName = sanitizeInput_(student.name || (typeof student === 'string' ? student : ''));
        var studentId = sanitizeInput_(student.id || '');
        var phone = sanitizeInput_(student.phone || '');

        rows.push([
          studentId, studentName, gradeName, classLetter, phone,
          'تأخر صباحي', '', hijriDate, safeStaffName, now, 'لا'
        ]);
        savedCount++;
      });

      // ★ كتابة دفعة واحدة
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }

      // تسجيل النشاط
      logStaffActivity_(data.staffName, 'تأخر', (gradeName + ' ' + classLetter).trim(), savedCount, stage);
    }
    
    return { 
      success: true, 
      message: 'تم تسجيل ' + savedCount + ' تأخر بنجاح',
      count: savedCount
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// ★ دالة مساعدة: التحقق من صلاحية المرحلة (whitelist)
// =================================================================
function validateStage_(stage) {
  if (!stage || typeof stage !== 'string') return false;
  ensureStudentsSheetsLoaded_();
  var validStages = STUDENTS_SHEETS ? Object.keys(STUDENTS_SHEETS) : [];
  return validStages.indexOf(stage) !== -1;
}

// =================================================================
// ★ دالة مساعدة: فصل اسم الفصل إلى صف + حرف فصل
// مثال: "الأول أ" → { grade: "الأول", section: "أ" }
// =================================================================
function parseClassName_(className, gradeName, classLetter) {
  // إذا أُرسلت القيم منفصلة استخدمها مباشرة
  if (gradeName && classLetter) return { grade: gradeName, section: classLetter };
  if (!className) return { grade: gradeName || '', section: classLetter || '' };

  var cn = String(className).trim();
  // حرف الفصل عادةً هو آخر كلمة (حرف واحد أو رقم)
  var parts = cn.split(/\s+/);
  if (parts.length >= 2) {
    var last = parts[parts.length - 1];
    // إذا الجزء الأخير حرف واحد أو رقم → هو حرف الفصل
    if (last.length <= 2 || /^\d+$/.test(last)) {
      return { grade: parts.slice(0, -1).join(' '), section: last };
    }
  }
  return { grade: cn, section: classLetter || '' };
}

// =================================================================
// 9. جلب سجلات الاستئذان للحارس (من الشيت التراكمي — فلترة اليوم)
// =================================================================
function getPermissionRecordsForGuard(stage, token) {
  try {
    // ★ التحقق من هوية الحارس عبر التوكن
    if (!token) return { success: false, error: 'التوكن مطلوب' };
    var auth = getStaffByToken(token);
    if (!auth.success) return { success: false, error: 'رابط غير صالح' };
    if (!auth.staff.isGuard) return { success: false, error: 'غير مصرح — للحراس فقط' };

    // ★ التحقق من صلاحية المرحلة ضد whitelist
    if (!validateStage_(stage)) {
      return { success: false, error: 'مرحلة غير صالحة' };
    }
    var ss = getSpreadsheet_();
    var sheetName = 'سجل_الاستئذان_' + stage;
    var sheet = findSheet_(ss, sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // ★ البحث بأسماء الأعمدة (آمن من تغيير الترتيب)
    var nameCol = findColumnIndex_(headers, ['اسم_الطالب', 'اسم الطالب']);
    var gradeCol = findColumnIndex_(headers, ['الصف']);
    var classCol = findColumnIndex_(headers, ['الفصل']);
    var reasonCol = findColumnIndex_(headers, ['السبب']);
    var timeCol = findColumnIndex_(headers, ['وقت_الخروج', 'وقت الخروج']);
    var confirmCol = findColumnIndex_(headers, ['وقت_التأكيد', 'وقت التأكيد']);
    var inputDateCol = findColumnIndex_(headers, ['وقت_الإدخال', 'وقت الإدخال']);
    var receiverCol = findColumnIndex_(headers, ['المستلم']);
    
    // ★ فلترة سجلات اليوم فقط
    var today = new Date();
    var todayStr = today.toDateString();
    var records = [];
    
    for (var i = 1; i < data.length; i++) {
      var inputDate = inputDateCol >= 0 ? new Date(data[i][inputDateCol]) : null;
      if (!inputDate || inputDate.toDateString() !== todayStr) continue;
      
      var confirmTime = confirmCol >= 0 ? String(data[i][confirmCol] || '').trim() : '';
      var className = (gradeCol >= 0 ? String(data[i][gradeCol] || '') : '') + ' ' + (classCol >= 0 ? String(data[i][classCol] || '') : '');
      
      records.push({
        rowIndex: i + 1, // ★ الصف الفعلي في الشيت (1-indexed + header)
        studentName: nameCol >= 0 ? String(data[i][nameCol] || '') : '',
        className: className.trim(),
        reason: reasonCol >= 0 ? String(data[i][reasonCol] || '') : '',
        time: timeCol >= 0 ? String(data[i][timeCol] || '') : '',
        receiver: receiverCol >= 0 ? String(data[i][receiverCol] || '') : '',
        status: confirmTime || 'غير مؤكد',
        isConfirmed: !!confirmTime && confirmTime !== 'غير مؤكد'
      });
    }
    
    return { success: true, records: records };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 10. تأكيد خروج الطالب (للحارس) — يكتب في الشيت التراكمي
// =================================================================
function confirmStudentExit(rowIndex, stage, token) {
  try {
    // ★ التحقق من هوية الحارس عبر التوكن
    if (!token) return { success: false, error: 'التوكن مطلوب' };
    var auth = getStaffByToken(token);
    if (!auth.success) return { success: false, error: 'رابط غير صالح' };
    if (!auth.staff.isGuard) return { success: false, error: 'غير مصرح — للحراس فقط' };

    // ★ التحقق من صلاحية المرحلة ضد whitelist
    if (!validateStage_(stage)) {
      return { success: false, error: 'مرحلة غير صالحة' };
    }
    var ss = getSpreadsheet_();
    var sheetName = 'سجل_الاستئذان_' + stage;
    var sheet = findSheet_(ss, sheetName);

    if (!sheet) {
      return { success: false, error: 'شيت الاستئذان غير موجود' };
    }
    // rowIndex هنا 1-indexed مباشرة — نتحقق أنه >= 2 (لا يمس الترويسة)
    if (typeof rowIndex !== 'number' || !isFinite(rowIndex) || rowIndex < 2 || rowIndex !== Math.floor(rowIndex) || rowIndex > sheet.getLastRow()) {
      return { success: false, error: 'rowIndex غير صالح' };
    }

    // ★ البحث عن عمود وقت_التأكيد بالاسم (بدل رقم ثابت)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var confirmCol = -1;
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim().replace(/\s+/g, '_') === 'وقت_التأكيد') {
        confirmCol = h + 1; // 1-indexed for getRange
        break;
      }
    }
    
    if (confirmCol < 0) {
      return { success: false, error: 'عمود وقت التأكيد غير موجود' };
    }
    
    var timeStr = new Date().toLocaleTimeString('ar-SA');
    sheet.getRange(rowIndex, confirmCol).setValue(timeStr).setHorizontalAlignment('center');
    
    return { success: true, message: 'تم تأكيد الخروج' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 11. جلب سجلات اليوم (للعرض)
// =================================================================
function getTodaysStaffEntries() {
  try {
    ensureStudentsSheetsLoaded_();
    var ss = getSpreadsheet_();
    var today = new Date();
    var todayStr = today.toDateString();
    var stages = Object.keys(STUDENTS_SHEETS);
    var entries = {};
    stages.forEach(function(s) { entries[s] = []; });
    
    // ★ القراءة من الشيتات التراكمية (نفس مكان الحفظ)
    stages.forEach(function(stage) {
      // 1) الاستئذان
      var permSheet = findSheet_(ss, 'سجل_الاستئذان_' + stage);
      if (permSheet && permSheet.getLastRow() > 1) {
        var pData = permSheet.getDataRange().getValues();
        var pH = pData[0];
        var pNameCol = findColumnIndex_(pH, ['اسم_الطالب', 'اسم الطالب']);
        var pTimeCol = findColumnIndex_(pH, ['وقت_الخروج', 'وقت الخروج']);
        var pDateCol = findColumnIndex_(pH, ['وقت_الإدخال', 'وقت الإدخال']);
        
        for (var i = 1; i < pData.length; i++) {
          var d = pDateCol >= 0 ? new Date(pData[i][pDateCol]) : null;
          if (!d || d.toDateString() !== todayStr) continue;
          entries[stage].push({
            name: pNameCol >= 0 ? String(pData[i][pNameCol] || '') : '',
            type: 'استئذان',
            time: pTimeCol >= 0 ? String(pData[i][pTimeCol] || '') : ''
          });
        }
      }
      
      // 2) التأخر
      var lateSheet = findSheet_(ss, 'سجل_التأخر_' + stage);
      if (lateSheet && lateSheet.getLastRow() > 1) {
        var lData = lateSheet.getDataRange().getValues();
        var lH = lData[0];
        var lNameCol = findColumnIndex_(lH, ['اسم_الطالب', 'اسم الطالب']);
        var lTimeCol = findColumnIndex_(lH, ['وقت_الإدخال', 'وقت الإدخال']);
        
        for (var j = 1; j < lData.length; j++) {
          var d2 = lTimeCol >= 0 ? new Date(lData[j][lTimeCol]) : null;
          if (!d2 || d2.toDateString() !== todayStr) continue;
          entries[stage].push({
            name: lNameCol >= 0 ? String(lData[j][lNameCol] || '') : '',
            type: 'تأخر',
            time: lTimeCol >= 0 ? Utilities.formatDate(d2, Session.getScriptTimeZone(), 'HH:mm') : ''
          });
        }
      }
    });
    
    return { success: true, entries: entries };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ★ sendStaffLinkWhatsApp — حُذفت (كود ميت — استُبدلت بـ sendLinkToPersonWithStage)
// =================================================================
// الدوال المساعدة
// =================================================================

function findColumnIndex_(headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx !== -1) return idx;
  }
  for (var h = 0; h < headers.length; h++) {
    for (var n = 0; n < possibleNames.length; n++) {
      if (String(headers[h]).includes(possibleNames[n])) {
        return h;
      }
    }
  }
  return -1; // ★ إصلاح: كانت 0 — تسبب في قراءة العمود الأول بالخطأ
}

function formatPhone_(phone) {
  phone = String(phone).replace(/\D/g, '');
  if (phone.startsWith('05')) {
    phone = '966' + phone.substring(1);
  } else if (phone.startsWith('5')) {
    phone = '966' + phone;
  }
  return phone;
}

// ★ getDailySheetName_ + rebuildDailyReport_ — حُذفتا (كود ميت — السجلات تُقرأ من الشيتات التراكمية)

function logStaffActivity_(staffName, type, className, count, stage) {
  try {
    var ss = getSpreadsheet_();
    var logSheet = ss.getSheetByName('سجل_النشاطات');

    if (!logSheet) {
      logSheet = ss.insertSheet('سجل_النشاطات');
      logSheet.setRightToLeft(true);
      logSheet.appendRow(['التاريخ', 'الوقت', 'المستخدم', 'النوع', 'التفاصيل', 'العدد', 'المرحلة']);
      logSheet.getRange(1, 1, 1, 7).setBackground('#1e3a5f').setFontColor('#ffffff');
    }

    logSheet.appendRow([
      new Date().toLocaleDateString('ar-SA'),
      new Date().toLocaleTimeString('ar-SA'),
      sanitizeInput_(staffName),
      type,
      'الفصل: ' + sanitizeInput_(className),
      count,
      stage || ''
    ]);
  } catch (e) {
    Logger.log('خطأ في logStaffActivity_: ' + e.toString());
  }
}

// ★ getStaffWithLinks — حُذفت (كود ميت — استُبدلت بـ getLinksTabData في Server_TeacherInput.gs)

// =================================================================
// ★ دوال بيانات الواجهات حسب الأدوار
// =================================================================

// بيانات صفحة الوكيل (جميع الميزات)
function getWakeelPageData_(token) {
  var staffResult = getStaffByToken(token);
  if (!staffResult.success) return staffResult;
  var stagesConfig = getEnabledStagesConfig_();
  return {
    success: true,
    staff: staffResult.staff,
    gradeMap: stagesConfig.stages || {}
  };
}

// بيانات صفحة الموجه الطلابي (استئذان + ملاحظات تربوية + سلوك متمايز)
function getCounselorPageData_(token) {
  var staffResult = getStaffByToken(token);
  if (!staffResult.success) return staffResult;
  var stagesConfig = getEnabledStagesConfig_();
  return {
    success: true,
    staff: staffResult.staff,
    gradeMap: stagesConfig.stages || {}
  };
}

// بيانات صفحة الإداري (تأخر صباحي فقط)
function getAdminPageData_(token) {
  var staffResult = getStaffByToken(token);
  if (!staffResult.success) return staffResult;
  var stagesConfig = getEnabledStagesConfig_();
  return {
    success: true,
    staff: staffResult.staff,
    gradeMap: stagesConfig.stages || {}
  };
}