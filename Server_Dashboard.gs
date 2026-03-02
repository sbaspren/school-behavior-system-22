// =================================================================
// Server_Dashboard.gs — بيانات لوحة المتابعة
// دالة واحدة تجمع كل الإحصائيات والبيانات المطلوبة
// =================================================================

// ★ ثوابت حدود المخالفات والغياب (بدل الأرقام السحرية)
var VIOLATION_DEGREE = {
  NEEDS_TAHOOD: 2,      // درجة تعهد سلوكي
  NEEDS_ISHAR: 3,       // درجة إشعار ولي أمر
  NEEDS_MAHDAR: 4,      // درجة محضر لجنة
  CRITICAL: 4            // الدرجة الحرجة
};
var ABSENCE_THRESHOLD = {
  NEEDS_TAHOOD: 3,       // عدد مرات الغياب لتعهد حضور
  NEEDS_WARNING: 5,      // عدد مرات الغياب لإنذار
  NEEDS_DEPRIVATION: 7   // عدد مرات الغياب لإشعار حرمان
};
var DASHBOARD_LIMITS = {
  MAX_VIOLATIONS_NO_ACTION: 5,  // أقصى عدد مخالفات بدون إجراء في اللوحة
  MAX_NOTES_PENDING: 5,         // أقصى عدد ملاحظات معلقة
  MAX_NEEDS_PRINTING: 15,       // أقصى عدد نماذج تحتاج طباعة
  MAX_RECENT_ACTIVITY: 6,       // أقصى عدد نشاطات حديثة
  MAX_RECENT_PER_SECTION: 3     // أقصى عدد من كل قسم
};

function getDashboardData() {
  try {
    ensureStudentsSheetsLoaded_();
    var stages = Object.keys(STUDENTS_SHEETS);
    var stageStats = {};
    for (var si = 0; si < stages.length; si++) {
      stageStats[stages[si]] = { absence: 0, tardiness: 0, permissions: 0, violations: 0, notes: 0 };
    }
    var result = {
      success: true,
      today: { absence: 0, tardiness: 0, permissions: 0, permissionsOut: 0, permissionsWaiting: 0, violations: 0, notes: 0, pendingExcuses: 0 },
      stageStats: stageStats,
      pending: { violationsNoAction: [], notesPending: [], notSent: { absence: 0, tardiness: 0, violations: 0 }, notSentByStage: {} },
      absenceByClass: {},
      recentActivity: [],
      semesterTotals: { violations: 0, absence: 0, permissions: 0, tardiness: 0 },
      needsPrinting: [],
      hijriOffset: 0,
      hijriDate: getHijriDateFull_(new Date())
    };

    var ss = getSpreadsheet_();

    for (var s = 0; s < stages.length; s++) {
      var stage = stages[s];
      
      // ★ تتبع "لم يُبلّغ" حسب المرحلة
      if (!result.pending.notSentByStage[stage]) {
        result.pending.notSentByStage[stage] = { absence: 0, tardiness: 0, violations: 0 };
      }

      // ── 1. الغياب اليومي ──
      try {
        var absResult = getTodayAbsenceRecords(stage);
        var absRecs = (absResult && absResult.records) ? absResult.records : [];
        
        // فصل سجلات "لا يوجد غائب" عن سجلات الغياب الفعلي
        var realAbsRecs = [];
        for (var a = 0; a < absRecs.length; a++) {
          var studentId = String(absRecs[a]['رقم_الطالب'] || '').trim();
          if (studentId === 'NO_ABSENCE') {
            // سجل تأكيد حضور — يدل على أن المعلم أدخل ولا يوجد غائب
            var naGrade = String(absRecs[a]['الصف'] || '').trim();
            var naCls = String(absRecs[a]['الفصل'] || '').trim();
            var naKey = naGrade + '-' + naCls;
            if (!result.absenceByClass[stage]) result.absenceByClass[stage] = {};
            if (result.absenceByClass[stage][naKey] === undefined) {
              result.absenceByClass[stage][naKey] = 0; // تم الإدخال: 0 غياب
            }
          } else {
            realAbsRecs.push(absRecs[a]);
          }
        }
        
        result.today.absence += realAbsRecs.length;
        result.stageStats[stage].absence += realAbsRecs.length;

        // عدد بدون إرسال
        for (var a = 0; a < realAbsRecs.length; a++) {
          if (realAbsRecs[a]['تم_الإرسال'] !== 'نعم') { result.pending.notSent.absence++; result.pending.notSentByStage[stage].absence++; }
        }

        // تجميع حسب الفصل
        if (!result.absenceByClass[stage]) result.absenceByClass[stage] = {};
        for (var a2 = 0; a2 < realAbsRecs.length; a2++) {
          var grade = String(realAbsRecs[a2]['الصف'] || '').trim();
          var cls = String(realAbsRecs[a2]['الفصل'] || '').trim();
          var key = grade + '-' + cls;
          result.absenceByClass[stage][key] = (result.absenceByClass[stage][key] || 0) + 1;
        }
      } catch(e) { Logger.log('Dashboard absence error: ' + e); }

      // ── 2. التأخر ──
      try {
        var tardResult = getTodayLateRecords(stage);
        var tardRecs = (tardResult && tardResult.records) ? tardResult.records : [];
        result.today.tardiness += tardRecs.length;
        result.stageStats[stage].tardiness += tardRecs.length;
        for (var t = 0; t < tardRecs.length; t++) {
          if (tardRecs[t]['تم_الإرسال'] !== 'نعم') { result.pending.notSent.tardiness++; result.pending.notSentByStage[stage].tardiness++; }
        }
      } catch(e) { Logger.log('Dashboard tardiness error: ' + e); }

      // ── 3. الاستئذان ──
      try {
        var permResult = getTodayPermissionRecords(stage);
        var permRecs = (permResult && permResult.records) ? permResult.records : [];
        result.today.permissions += permRecs.length;
        result.stageStats[stage].permissions += permRecs.length;
        for (var p = 0; p < permRecs.length; p++) {
          var confirmed = permRecs[p]['تم_الخروج'] || permRecs[p]['حالة_الخروج'] || '';
          if (confirmed === 'نعم' || confirmed === 'تم') {
            result.today.permissionsOut++;
          } else {
            result.today.permissionsWaiting++;
          }
        }
      } catch(e) { Logger.log('Dashboard permissions error: ' + e); }

      // ── 4. المخالفات ──
      try {
        var violRecs = getViolationRecords(stage);
        // مخالفات اليوم
        var todayViols = filterTodayViolations_(violRecs);
        result.today.violations += todayViols.length;
        result.stageStats[stage].violations += todayViols.length;

        for (var v = 0; v < todayViols.length; v++) {
          if (todayViols[v]['تم الإرسال'] !== 'نعم') { result.pending.notSent.violations++; result.pending.notSentByStage[stage].violations++; }
        }

        // مخالفات بدون إجراء (كل المخالفات وليس فقط اليوم)
        for (var v2 = 0; v2 < violRecs.length; v2++) {
          var proc = String(violRecs[v2]['الإجراءات'] || '').trim();
          var forms = String(violRecs[v2]['النماذج المحفوظة'] || '').trim();
          if (!proc && !forms) {
            result.pending.violationsNoAction.push({
              name: violRecs[v2]['اسم الطالب'] || '',
              violation: violRecs[v2]['نص المخالفة'] || '',
              grade: violRecs[v2]['الصف'] || '',
              cls: violRecs[v2]['الفصل'] || '',
              degree: violRecs[v2]['الدرجة'] || '',
              date: violRecs[v2]['التاريخ الهجري'] || '',
              stage: stage
            });
          }
        }

        // إجمالي الفصل
        result.semesterTotals.violations += violRecs.length;

        // ── مخالفات تحتاج توثيق (طباعة نماذج) ──
        for (var vp = 0; vp < violRecs.length; vp++) {
          var vr = violRecs[vp];
          var deg = parseInt(vr['الدرجة']) || 0;
          var savedForms = String(vr['النماذج المحفوظة'] || '').trim();
          if (deg >= VIOLATION_DEGREE.NEEDS_TAHOOD && !savedForms) {
            var neededForms = [];
            if (deg >= VIOLATION_DEGREE.NEEDS_TAHOOD) neededForms.push('تعهد سلوكي');
            if (deg >= VIOLATION_DEGREE.NEEDS_ISHAR) neededForms.push('إشعار ولي أمر');
            if (deg >= VIOLATION_DEGREE.NEEDS_MAHDAR) neededForms.push('محضر لجنة');
            result.needsPrinting.push({
              type: 'مخالفة',
              name: vr['اسم الطالب'] || '',
              studentId: vr['رقم الطالب'] || '',
              detail: vr['نص المخالفة'] || '',
              degree: deg,
              grade: vr['الصف'] || '',
              cls: vr['الفصل'] || '',
              date: vr['التاريخ الهجري'] || '',
              neededForms: neededForms,
              stage: stage,
              rowIndex: vr.rowIndex || 0,
              section: 'violations'
            });
          }
        }

        // آخر التحويلات (من المعلمين)
        for (var v3 = todayViols.length - 1; v3 >= Math.max(0, todayViols.length - DASHBOARD_LIMITS.MAX_RECENT_PER_SECTION); v3--) {
          var actionField = todayViols[v3]['التصرف المتخذ'] || todayViols[v3]['التصرف_المتخذ'] || '';
          result.recentActivity.push({
            type: 'مخالفة',
            teacher: todayViols[v3]['المستخدم'] || 'الوكيل',
            detail: todayViols[v3]['نص المخالفة'] || '',
            student: todayViols[v3]['اسم الطالب'] || '',
            cls: (todayViols[v3]['الصف'] || '') + ' ' + (todayViols[v3]['الفصل'] || ''),
            time: todayViols[v3]['وقت الإدخال'] || '',
            stage: stage,
            section: 'violations',
            actionTaken: actionField !== ''
          });
        }
      } catch(e) { Logger.log('Dashboard violations error: ' + e); }

      // ── 5. الملاحظات التربوية ──
      try {
        var notesResult = getTodayEducationalNotesRecords(stage);
        var notesRecs = (notesResult && notesResult.records) ? notesResult.records : [];
        result.today.notes += notesRecs.length;
        result.stageStats[stage].notes += notesRecs.length;

        // الملاحظات المعلقة (بدون إرسال)
        for (var n = 0; n < notesRecs.length; n++) {
          if (notesRecs[n]['تم_الإرسال'] !== 'نعم') {
            result.pending.notesPending.push({
              name: notesRecs[n]['اسم_الطالب'] || notesRecs[n]['اسم الطالب'] || '',
              type: notesRecs[n]['نوع_الملاحظة'] || notesRecs[n]['نوع الملاحظة'] || '',
              detail: notesRecs[n]['التفاصيل'] || '',
              teacher: notesRecs[n]['المعلم/المسجل'] || notesRecs[n]['المعلم_المسجل'] || '',
              cls: (notesRecs[n]['الصف'] || '') + ' ' + (notesRecs[n]['الفصل'] || ''),
              stage: stage
            });
          }

          // آخر التحويلات
          if (result.recentActivity.length < DASHBOARD_LIMITS.MAX_RECENT_ACTIVITY + 2) {
            var noteActionField = notesRecs[n]['تم_الإرسال'] || '';
            result.recentActivity.push({
              type: 'ملاحظة',
              teacher: notesRecs[n]['المعلم/المسجل'] || notesRecs[n]['المعلم_المسجل'] || '',
              detail: notesRecs[n]['نوع_الملاحظة'] || notesRecs[n]['نوع الملاحظة'] || '',
              student: notesRecs[n]['اسم_الطالب'] || notesRecs[n]['اسم الطالب'] || '',
              cls: (notesRecs[n]['الصف'] || '') + ' ' + (notesRecs[n]['الفصل'] || ''),
              time: notesRecs[n]['وقت_الإدخال'] || '',
              stage: stage,
              section: 'educational-notes',
              actionTaken: noteActionField === 'نعم'
            });
          }
        }
      } catch(e) { Logger.log('Dashboard notes error: ' + e); }

      // ── 6. إجماليات الفصل (تأخر + استئذان + غياب) ──
      try {
        var lateSheet = getLateSheet(stage);
        if (lateSheet && lateSheet.getLastRow() > 1) result.semesterTotals.tardiness += lateSheet.getLastRow() - 1;
      } catch(e) {}
      try {
        var permSheet = getPermissionSheet(stage);
        if (permSheet && permSheet.getLastRow() > 1) result.semesterTotals.permissions += permSheet.getLastRow() - 1;
      } catch(e) {}
      try {
        var absSS2 = getSpreadsheet_();
        var absSheetName2 = 'سجل_الغياب_اليومي_' + stage;
        var absSheet = findSheet_(absSS2, absSheetName2);
        if (absSheet && absSheet.getLastRow() > 1) result.semesterTotals.absence += absSheet.getLastRow() - 1;
      } catch(e) {}
    }

    // ── 7. أعذار بانتظار الاعتماد ──
    try {
      var absSS = getSpreadsheet_();
      for (var s2 = 0; s2 < stages.length; s2++) {
        var absSheetName = 'سجل_الغياب_اليومي_' + stages[s2];
        var absSheetPending = findSheet_(absSS, absSheetName);
        if (absSheetPending && absSheetPending.getLastRow() > 1) {
          var absData = absSheetPending.getDataRange().getValues();
          var absHeaders = absData[0];
          var statusCol = -1;
          var excuseCol = -1;
          for (var h = 0; h < absHeaders.length; h++) {
            var hName = String(absHeaders[h]).trim().replace(/\s+/g, '_');
            if (hName === 'حالة_الاعتماد') statusCol = h;
            if (hName === 'نوع_العذر') excuseCol = h;
          }
          if (statusCol > -1) {
            for (var r = 1; r < absData.length; r++) {
              var status = String(absData[r][statusCol] || '').trim();
              var excuse = excuseCol > -1 ? String(absData[r][excuseCol] || '').trim() : '';
              if (excuse && status !== 'معتمد' && status !== 'مقبول' && status !== 'مرفوض') {
                result.today.pendingExcuses++;
              }
            }
          }
        }
      }
    } catch(e) { Logger.log('Dashboard excuses error: ' + e); }

    // ── 7b. أعذار أولياء الأمور المعلقة من شيت اعذار_اولياء_الامور ──
    try {
      var peSheet = findSheet_(getSpreadsheet_(), 'اعذار_اولياء_الامور');
      if (peSheet && peSheet.getLastRow() >= 2) {
        var peData = peSheet.getDataRange().getValues();
        var peHeaders = peData[0].map(function(h) { return String(h).trim().replace(/\s+/g, '_'); });
        var peStatusIdx = peHeaders.indexOf('الحالة');
        if (peStatusIdx >= 0) {
          for (var pe = 1; pe < peData.length; pe++) {
            if (String(peData[pe][peStatusIdx]).trim() === 'معلق') {
              result.today.pendingExcuses++;
            }
          }
        }
      }
    } catch(e) { Logger.log('Dashboard parent excuses error: ' + e); }

    // ── 8. غياب يحتاج توثيق (3 مرات فأكثر بدون عذر) ──
    try {
      var absSS3 = getSpreadsheet_();
      for (var s3 = 0; s3 < stages.length; s3++) {
        var absSheetName3 = 'سجل_الغياب_اليومي_' + stages[s3];
        var absSheet3 = findSheet_(absSS3, absSheetName3);
        if (!absSheet3 || absSheet3.getLastRow() < 2) continue;
        var absData3 = absSheet3.getDataRange().getValues();
        var absH3 = absData3[0];
        var colMap3 = {};
        for (var h3 = 0; h3 < absH3.length; h3++) { colMap3[String(absH3[h3]).trim().replace(/\s+/g,'_')] = h3; }

        var idCol = colMap3['رقم_الطالب']; if (idCol === undefined) continue;
        var nameCol = colMap3['اسم_الطالب'];
        var gradeCol = colMap3['الصف'];
        var clsCol = colMap3['الفصل'];
        var excCol = colMap3['نوع_العذر'];
        var statusCol3 = colMap3['حالة_الاعتماد'];

        // عدّ غياب بدون عذر لكل طالب
        var absCount = {};
        for (var r3 = 1; r3 < absData3.length; r3++) {
          var row3 = absData3[r3];
          var sid = String(row3[idCol] || '').trim();
          if (!sid || sid === 'NO_ABSENCE') continue;
          var excuse3 = excCol !== undefined ? String(row3[excCol] || '').trim() : '';
          var status3 = statusCol3 !== undefined ? String(row3[statusCol3] || '').trim() : '';
          // بدون عذر أو عذر مرفوض
          if (!excuse3 || excuse3 === 'بدون عذر' || status3 === 'مرفوض') {
            if (!absCount[sid]) {
              absCount[sid] = {
                count: 0,
                name: nameCol !== undefined ? String(row3[nameCol] || '') : '',
                grade: gradeCol !== undefined ? String(row3[gradeCol] || '') : '',
                cls: clsCol !== undefined ? String(row3[clsCol] || '') : ''
              };
            }
            absCount[sid].count++;
          }
        }

        // الطلاب الذين وصلوا للحد الأدنى فأكثر
        for (var sid2 in absCount) {
          var ac = absCount[sid2];
          if (ac.count >= ABSENCE_THRESHOLD.NEEDS_TAHOOD) {
            var neededForms = ['تعهد حضور'];
            if (ac.count >= ABSENCE_THRESHOLD.NEEDS_WARNING) neededForms.push('إنذار غياب');
            if (ac.count >= ABSENCE_THRESHOLD.NEEDS_DEPRIVATION) neededForms.push('إشعار حرمان');
            result.needsPrinting.push({
              type: 'غياب',
              name: ac.name,
              studentId: sid2,
              detail: ac.count + ' مرة غياب بدون عذر',
              degree: ac.count,
              grade: ac.grade,
              cls: ac.cls,
              date: '',
              neededForms: neededForms,
              stage: stages[s3],
              section: 'absence'
            });
          }
        }
      }
    } catch(e) { Logger.log('Dashboard absence docs error: ' + e); }

    // ترتيب النشاط حسب الوقت (الأحدث أولاً)
    result.recentActivity.sort(function(a, b) {
      return (b.time || '').localeCompare(a.time || '');
    });
    // أخذ أحدث النشاطات
    result.recentActivity = result.recentActivity.slice(0, DASHBOARD_LIMITS.MAX_RECENT_ACTIVITY);

    // تقليل حجم البيانات
    result.pending.violationsNoAction = result.pending.violationsNoAction.slice(0, DASHBOARD_LIMITS.MAX_VIOLATIONS_NO_ACTION);
    result.pending.notesPending = result.pending.notesPending.slice(0, DASHBOARD_LIMITS.MAX_NOTES_PENDING);
    result.needsPrinting = result.needsPrinting.slice(0, DASHBOARD_LIMITS.MAX_NEEDS_PRINTING);

    return result;

  } catch (e) {
    Logger.log('❌ getDashboardData error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ── مساعد: فلترة مخالفات اليوم ──
function filterTodayViolations_(records) {
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  var todayParts = todayStr.split('/');
  
  return records.filter(function(rec) {
    var dateVal = rec['التاريخ الميلادي'] || rec['وقت الإدخال'] || '';
    if (!dateVal) return false;
    var str = String(dateVal);
    // مقارنة بالتاريخ
    return str.indexOf(todayParts[0] + '/' + todayParts[1] + '/' + todayParts[2]) > -1 ||
           str.indexOf(todayParts[2] + '/' + todayParts[1] + '/' + todayParts[0]) > -1;
  });
}

// =================================================================
// ★ بيانات التقويم الدراسي — مركزية على الخادم (إصلاح #12)
// يمكن تحديثها هنا بدلاً من تعديل كود العميل
// =================================================================
function getSchoolCalendarData() {
  return {
    events: [
      {d:24,m:8,  label:'بداية العام الدراسي', type:'event'},
      {d:23,m:9,  label:'إجازة اليوم الوطني', type:'national', holiday:true},
      {d:5, m:10, label:'يوم المعلم العالمي', type:'event'},
      {d:12,m:10, label:'إجازة إضافية', type:'holiday', holiday:true},
      {d:16,m:10, label:'يوم الغذاء العالمي', type:'event'},
      {d:16,m:11, label:'اليوم العالمي للتسامح', type:'event'},
      {d:20,m:11, label:'اليوم العالمي للطفل', type:'event'},
      {d:21,m:11, label:'بداية إجازة الخريف', type:'holiday', holiday:true},
      {d:3, m:12, label:'اليوم العالمي لذوي الإعاقة', type:'event'},
      {d:11,m:12, label:'إجازة إضافية', type:'holiday', holiday:true},
      {d:18,m:12, label:'اليوم العالمي للغة العربية', type:'event'},
      {d:9, m:1,  label:'بداية إجازة منتصف العام', type:'holiday', holiday:true},
      {d:18,m:1,  label:'بداية الفصل الثاني', type:'event'},
      {d:24,m:1,  label:'اليوم الدولي للتعليم', type:'event'},
      {d:22,m:2,  label:'يوم التأسيس السعودي', type:'national', holiday:true},
      {d:6, m:3,  label:'بداية إجازة عيد الفطر', type:'holiday', holiday:true},
      {d:28,m:3,  label:'نهاية إجازة عيد الفطر', type:'event'},
      {d:11,m:3,  label:'يوم العلم السعودي', type:'national'},
      {d:7, m:4,  label:'اليوم العالمي للصحة', type:'event'},
      {d:22,m:5,  label:'بداية إجازة عيد الأضحى', type:'holiday', holiday:true},
      {d:1, m:6,  label:'نهاية إجازة عيد الأضحى', type:'event'},
      {d:25,m:6,  label:'بداية إجازة نهاية العام', type:'holiday', holiday:true}
    ],
    semesters: [
      { name: 'الفصل الأول',  start: [2025, 7, 24], end: [2026, 0, 8], weeks: 18 },
      { name: 'الفصل الثاني', start: [2026, 0, 18], end: [2026, 5, 25], weeks: 18 }
    ]
  };
}