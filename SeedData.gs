// =================================================================
// ⚠️ ملف تطويري فقط — يجب حذفه من بيئة الإنتاج
// SeedData.gs - بيانات وهمية شاملة للاختبار
// ★ شغّل seedAllTestData() مرة واحدة من المحرر
// ★ شغّل clearAllTestData() لحذف البيانات الوهمية
// =================================================================

var TEST_PHONE = '966546545556';
var TEST_ID_PREFIX_INT = '9990';  // بادئة أرقام طلاب المتوسط الوهميين
var TEST_ID_PREFIX_SEC = '9991';  // بادئة أرقام طلاب الثانوي الوهميين

// =================================================================
// ★ الدالة الرئيسية - شغّلها من المحرر
// =================================================================
function seedAllTestData() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('★ بدء إدخال البيانات الوهمية');
  Logger.log('═══════════════════════════════════════');
  
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var now = new Date();
  var hijriToday = getHijriDate_(now);
  var miladi = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy/MM/dd');
  var timeNow = now; // Date object - ليتعرف عليه filterTodayRecords_
  var dayName = getDayNameAr_(now);
  
  // تواريخ سابقة (للأرشيف والتقارير)
  var dates = [];
  for (var d = 1; d <= 5; d++) {
    var past = new Date(now.getTime() - d * 24 * 60 * 60 * 1000);
    dates.push({
      hijri: getHijriDate_(past),
      miladi: Utilities.formatDate(past, 'Asia/Riyadh', 'yyyy/MM/dd'),
      time: past, // Date object
      day: getDayNameAr_(past)
    });
  }
  
  // 1. إضافة الطلاب الوهميين
  var students = seedStudents_(ss);
  Logger.log('✅ تم إضافة ' + students.length + ' طالب وهمي');
  
  // 2. بيانات اليوم + أيام سابقة لكل مرحلة
  var stages = ['متوسط', 'ثانوي'];
  for (var s = 0; s < stages.length; s++) {
    var stage = stages[s];
    var stageStudents = students.filter(function(st) { return st.stage === stage; });
    
    seedViolations_(ss, stage, stageStudents, hijriToday, miladi, timeNow, dates);
    seedTardiness_(ss, stage, stageStudents, hijriToday, timeNow, dates);
    seedPermissions_(ss, stage, stageStudents, hijriToday, timeNow, dates);
    seedDailyAbsence_(ss, stage, stageStudents, hijriToday, dayName, timeNow, dates);
    seedCumulativeAbsence_(ss, stage, stageStudents);
    seedNotes_(ss, stage, stageStudents, hijriToday, timeNow, dates);
    
    Logger.log('✅ اكتملت بيانات المرحلة: ' + stage);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('🎉 اكتملت جميع البيانات الوهمية!');
  Logger.log('═══════════════════════════════════════');
}

// =================================================================
// ★ حذف كل البيانات الوهمية
// =================================================================
function clearAllTestData() {
  Logger.log('★ بدء حذف البيانات الوهمية...');
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  // 1. حذف الطلاب الوهميين من شيتات الطلاب
  var stages = ['متوسط', 'ثانوي'];
  for (var s = 0; s < stages.length; s++) {
    var sheet = ss.getSheetByName(STUDENTS_SHEETS[stages[s]]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var id = String(data[i][0] || '');
      if (id.startsWith(TEST_ID_PREFIX_INT) || id.startsWith(TEST_ID_PREFIX_SEC)) {
        sheet.deleteRow(i + 1);
      }
    }
  }
  Logger.log('✅ حُذف الطلاب الوهميون');
  
  // 2. مسح كل شيتات السجلات (ما عدا الترويسة)
  var types = Object.keys(SHEET_REGISTRY);
  for (var t = 0; t < types.length; t++) {
    var reg = SHEET_REGISTRY[types[t]];
    if (!reg.perStage) continue;
    for (var s = 0; s < stages.length; s++) {
      var sheetName = reg.prefix + '_' + stages[s];
      var sheet = findSheet_(ss, sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        Logger.log('🗑️ مُسح: ' + sheetName);
      }
    }
  }
  
  Logger.log('🎉 اكتمل الحذف!');
}

// =================================================================
// بيانات الطلاب الوهميين
// =================================================================
function seedStudents_(ss) {
  var allStudents = [];
  
  var names = {
    'متوسط': {
      grades: ['الأول المتوسط', 'الثاني المتوسط', 'الثالث المتوسط'],
      students: [
        'سعود محمد الغامدي', 'فارس أحمد العسيري', 'عبدالله خالد الشهري',
        'ياسر سلمان القحطاني', 'بدر عبدالرحمن الزهراني', 'نواف حسين المالكي',
        'مشعل سعيد البيشي', 'تركي محمد الأحمري', 'حمد فهد الدوسري',
        'سلطان ناصر الحربي', 'ريان علي العتيبي', 'زياد وليد الشمراني',
        'هاشم طلال المطيري', 'أسامة بندر السبيعي', 'خالد يوسف الجهني',
        'عادل سعد الرشيدي', 'ماجد عمر الثبيتي', 'فيصل مبارك الخالدي',
        'محمد إبراهيم الحارثي', 'طلال عبدالعزيز القرني'
      ]
    },
    'ثانوي': {
      grades: ['الأول الثانوي', 'الثاني الثانوي', 'الثالث الثانوي'],
      students: [
        'عمر حسن الشيباني', 'أحمد عبدالله الخثعمي', 'راكان سعيد البارقي',
        'عبدالعزيز فهد النجادي', 'حسام محمد الحازمي', 'بسام علي الكلبي',
        'وليد ناصر العمري', 'يزن خالد المزيني', 'معاذ سلطان الوادعي',
        'سامي عبدالرحمن السلمي', 'ثامر أحمد الرحيلي', 'أيمن فيصل التميمي',
        'رائد محمد الهلالي', 'صالح إبراهيم الصاعدي', 'نايف طارق الفيفي',
        'مهند سعود الغامدي', 'عمار خالد الجعيد', 'إياد عادل الأسمري',
        'عبدالملك وليد الشهراني', 'غازي مسفر الريثي'
      ]
    }
  };
  
  var stages = ['متوسط', 'ثانوي'];
  for (var s = 0; s < stages.length; s++) {
    var stage = stages[s];
    var sheet = ss.getSheetByName(STUDENTS_SHEETS[stage]);
    if (!sheet) continue;
    
    var prefix = stage === 'متوسط' ? TEST_ID_PREFIX_INT : TEST_ID_PREFIX_SEC;
    var info = names[stage];
    var rows = [];
    
    for (var i = 0; i < info.students.length; i++) {
      var gradeIdx = Math.floor(i / 7); // ~7 طلاب لكل صف
      if (gradeIdx >= info.grades.length) gradeIdx = info.grades.length - 1;
      var cls = (i % 2) + 1; // فصل 1 أو 2
      var id = prefix + String(100 + i);
      
      var studentRow = [id, info.students[i], info.grades[gradeIdx], String(cls), TEST_PHONE];
      rows.push(studentRow);
      
      allStudents.push({
        id: id,
        name: info.students[i],
        grade: info.grades[gradeIdx],
        cls: String(cls),
        phone: TEST_PHONE,
        stage: stage
      });
    }
    
    // إضافة بعد آخر صف
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  }
  
  return allStudents;
}

// =================================================================
// المخالفات (5 اليوم + 10 سابقة لكل مرحلة)
// =================================================================
function seedViolations_(ss, stage, students, hijriToday, miladiToday, timeNow, dates) {
  var sheetName = 'سجل_المخالفات_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var violations = [
    { id: '101', text: 'التأخر عن الطابور الصباحي', type: 'سلوكية', degree: '1', points: '1' },
    { id: '102', text: 'عدم إحضار الكتب المدرسية', type: 'تعليمية', degree: '1', points: '1' },
    { id: '201', text: 'الشغب أثناء الحصة', type: 'سلوكية', degree: '2', points: '3' },
    { id: '202', text: 'استخدام الجوال', type: 'سلوكية', degree: '2', points: '3' },
    { id: '301', text: 'التنمر على الزملاء', type: 'سلوكية', degree: '3', points: '5' }
  ];
  
  var rows = [];
  
  // 5 مخالفات اليوم
  for (var i = 0; i < 5 && i < students.length; i++) {
    var v = violations[i % violations.length];
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      v.id, v.text, v.type, v.degree,
      hijriToday, miladiToday, '1', 'إنذار شفهي',
      v.points, '', '', 'مدير_النظام', timeNow
    ]);
  }
  
  // 10 مخالفات سابقة (للأرشيف والتقارير)
  for (var d = 0; d < dates.length; d++) {
    for (var i = 0; i < 2 && i < students.length; i++) {
      var idx = (d * 2 + i + 5) % students.length;
      var v = violations[(d + i) % violations.length];
      rows.push([
        students[idx].id, students[idx].name, students[idx].grade, students[idx].cls,
        v.id, v.text, v.type, v.degree,
        dates[d].hijri, dates[d].miladi, '1', 'إنذار كتابي',
        v.points, '', '', 'مدير_النظام', dates[d].time
      ]);
    }
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  📝 المخالفات ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// التأخر (8 اليوم + 15 سابقة)
// =================================================================
function seedTardiness_(ss, stage, students, hijriToday, timeNow, dates) {
  var sheetName = 'سجل_التأخر_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var lateTypes = ['تأخر صباحي', 'تأخر عن الحصة', 'تأخر صباحي'];
  var periods = ['', 'الثانية', 'الثالثة', 'الرابعة', 'الخامسة'];
  var rows = [];
  
  // 8 متأخرين اليوم
  for (var i = 0; i < 8 && i < students.length; i++) {
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      students[i].phone, lateTypes[i % 3], periods[i % 5],
      hijriToday, 'مدير_النظام', timeNow, 'لا'
    ]);
  }
  
  // 15 سجل سابق
  for (var d = 0; d < dates.length; d++) {
    for (var i = 0; i < 3 && i < students.length; i++) {
      var idx = (d * 3 + i + 8) % students.length;
      rows.push([
        students[idx].id, students[idx].name, students[idx].grade, students[idx].cls,
        students[idx].phone, lateTypes[(d + i) % 3], periods[(d + i) % 5],
        dates[d].hijri, 'مدير_النظام', dates[d].time, 'نعم'
      ]);
    }
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  ⏰ التأخر ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// الاستئذان (3 اليوم + 10 سابقة)
// =================================================================
function seedPermissions_(ss, stage, students, hijriToday, timeNow, dates) {
  var sheetName = 'سجل_الاستئذان_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var reasons = ['مراجعة طبية', 'ظروف عائلية', 'موعد أسنان', 'مراجعة مستشفى', 'ظروف خاصة'];
  var receivers = ['والده', 'والدته', 'أخوه الأكبر', 'عمه', 'والده'];
  var times = ['09:30', '10:15', '11:00', '09:45', '10:30'];
  var rows = [];
  
  // 3 مستأذنين اليوم
  for (var i = 0; i < 3 && i < students.length; i++) {
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      students[i].phone, times[i], reasons[i], receivers[i],
      'وكيل_المدرسة', hijriToday, 'مدير_النظام', timeNow, '', 'لا'
    ]);
  }
  
  // 10 سجلات سابقة
  for (var d = 0; d < dates.length; d++) {
    for (var i = 0; i < 2 && i < students.length; i++) {
      var idx = (d * 2 + i + 3) % students.length;
      rows.push([
        students[idx].id, students[idx].name, students[idx].grade, students[idx].cls,
        students[idx].phone, times[(d + i) % 5], reasons[(d + i) % 5], receivers[(d + i) % 5],
        'وكيل_المدرسة', dates[d].hijri, 'مدير_النظام', dates[d].time, dates[d].time, 'نعم'
      ]);
    }
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  🚪 الاستئذان ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// الغياب اليومي (10 اليوم + 20 سابقة)
// =================================================================
function seedDailyAbsence_(ss, stage, students, hijriToday, dayName, timeNow, dates) {
  var sheetName = 'سجل_الغياب_اليومي_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var absTypes = ['يوم كامل', 'يوم كامل', 'حصة'];
  var excuseTypes = ['بعذر', 'بدون عذر', 'بدون عذر', 'بعذر', 'بدون عذر'];
  var statuses = ['معلق', 'معلق', 'معتمد', 'مرفوض', 'معلق'];
  var rows = [];

  // 10 غائبين اليوم
  for (var i = 0; i < 10 && i < students.length; i++) {
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      students[i].phone, absTypes[i % 3], '',
      hijriToday, dayName, 'مدير_النظام', timeNow,
      statuses[i % 5], excuseTypes[i % 5], 'لا', 'غائب', '', ''
    ]);
  }

  // 20 سجل سابق (4 لكل يوم × 5 أيام)
  for (var d = 0; d < dates.length; d++) {
    for (var i = 0; i < 4 && i < students.length; i++) {
      var idx = (d * 4 + i + 10) % students.length;
      rows.push([
        students[idx].id, students[idx].name, students[idx].grade, students[idx].cls,
        students[idx].phone, absTypes[i % 3], '',
        dates[d].hijri, dates[d].day, 'مدير_النظام', dates[d].time,
        'معتمد', excuseTypes[(d + i) % 5], 'نعم', 'غائب', '', ''
      ]);
    }
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  📋 الغياب اليومي ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// الغياب التراكمي (ملخص لكل طالب)
// =================================================================
function seedCumulativeAbsence_(ss, stage, students) {
  var sheetName = 'سجل_الغياب_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var rows = [];
  for (var i = 0; i < students.length; i++) {
    var excused = Math.floor(Math.random() * 5);
    var unexcused = Math.floor(Math.random() * 3);
    var late = Math.floor(Math.random() * 4);
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      excused, unexcused, late, new Date()
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  📊 الغياب التراكمي ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// الملاحظات التربوية (4 اليوم + 10 سابقة)
// =================================================================
function seedNotes_(ss, stage, students, hijriToday, timeNow, dates) {
  var sheetName = 'سجل_الملاحظات_التربوية_' + stage;
  var sheet = findSheet_(ss, sheetName);
  if (!sheet) return;
  
  var noteTypes = ['إيجابية', 'سلبية', 'إيجابية', 'ملاحظة عامة', 'سلبية'];
  var details = [
    'تفوق في مادة الرياضيات وحصل على الدرجة الكاملة',
    'كثير الحديث أثناء الشرح ويشتت انتباه زملائه',
    'ساعد زميله في فهم الدرس وأظهر روح التعاون',
    'يحتاج متابعة في مستوى القراءة والكتابة',
    'لم يلتزم بالزي المدرسي لمدة أسبوع متواصل'
  ];
  var teachers = ['أ. محمد الشهري', 'أ. خالد القحطاني', 'أ. سعيد العسيري', 'أ. فهد المالكي'];
  var rows = [];
  
  // 4 ملاحظات اليوم
  for (var i = 0; i < 4 && i < students.length; i++) {
    rows.push([
      students[i].id, students[i].name, students[i].grade, students[i].cls,
      students[i].phone, noteTypes[i], details[i],
      teachers[i % 4], hijriToday, timeNow, 'لا'
    ]);
  }
  
  // 10 ملاحظات سابقة
  for (var d = 0; d < dates.length; d++) {
    for (var i = 0; i < 2 && i < students.length; i++) {
      var idx = (d * 2 + i + 4) % students.length;
      rows.push([
        students[idx].id, students[idx].name, students[idx].grade, students[idx].cls,
        students[idx].phone, noteTypes[(d + i) % 5], details[(d + i) % 5],
        teachers[(d + i) % 4], dates[d].hijri, dates[d].time, 'نعم'
      ]);
    }
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('  📝 الملاحظات ' + stage + ': ' + rows.length + ' سجل');
}

// =================================================================
// أدوات مساعدة
// =================================================================
// getHijriDate_() و getDayNameAr_() → مركزية في Config.gs