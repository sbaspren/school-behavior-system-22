// =================================================================
// DiagnosticAbsence.gs - أداة تشخيص بيانات الغياب اليومي
// انسخ هذا الكود في محرر Apps Script وشغّل الدالة:
// diagnoseDailyAbsence()
// ثم انسخ النتيجة من Logger (عرض > السجلات)
// =================================================================

function diagnoseDailyAbsence() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var output = [];

  output.push('╔══════════════════════════════════════════════════════════╗');
  output.push('║       تشخيص شيتات الغياب اليومي                        ║');
  output.push('║       ' + new Date().toLocaleString('ar-SA') + '       ║');
  output.push('╚══════════════════════════════════════════════════════════╝');
  output.push('');

  // ========== 1) البحث عن جميع شيتات الغياب اليومي ==========
  var dailySheets = [];
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (name.indexOf('سجل_الغياب_اليومي') !== -1) {
      dailySheets.push(allSheets[i]);
    }
  }

  output.push('═══ [1] الشيتات المكتشفة ═══');
  output.push('عدد شيتات الغياب اليومي: ' + dailySheets.length);
  for (var i = 0; i < dailySheets.length; i++) {
    output.push('  ' + (i+1) + '. ' + dailySheets[i].getName() + ' (صفوف: ' + dailySheets[i].getLastRow() + ', أعمدة: ' + dailySheets[i].getLastColumn() + ')');
  }
  output.push('');

  // ========== 2) الأعمدة المتوقعة (17 عمود) ==========
  var expectedHeaders = [
    'رقم_الطالب',      // 1
    'اسم_الطالب',      // 2
    'الصف',            // 3
    'الفصل',           // 4
    'رقم_الجوال',      // 5
    'نوع_الغياب',      // 6
    'الحصة',           // 7
    'التاريخ_هجري',    // 8
    'اليوم',           // 9
    'المسجل',          // 10
    'وقت_الإدخال',     // 11
    'حالة_الاعتماد',   // 12
    'نوع_العذر',       // 13
    'تم_الإرسال',      // 14
    'حالة_التأخر',     // 15
    'وقت_الحضور',      // 16
    'ملاحظات'          // 17
  ];

  // ========== 3) تحليل كل شيت ==========
  for (var s = 0; s < dailySheets.length; s++) {
    var sheet = dailySheets[s];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    output.push('');
    output.push('╔══════════════════════════════════════════════════════════╗');
    output.push('║  تحليل: ' + sheetName);
    output.push('╚══════════════════════════════════════════════════════════╝');

    if (lastRow < 1) {
      output.push('  ⚠ الشيت فارغ تماماً');
      continue;
    }

    // ---------- 3أ) تحليل الهيدرز ----------
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    output.push('');
    output.push('─── الهيدرز (عدد الأعمدة: ' + lastCol + ') ───');

    for (var h = 0; h < headers.length; h++) {
      var headerVal = String(headers[h] || '').trim();
      var expectedVal = h < expectedHeaders.length ? expectedHeaders[h] : '(عمود إضافي)';
      var match = (headerVal === expectedVal) ? '✓' : '✗ متوقع: ' + expectedVal;
      output.push('  عمود ' + (h+1) + ': [' + headerVal + '] ' + match);
    }

    if (lastCol < 17) {
      output.push('  ⚠ ناقص ' + (17 - lastCol) + ' أعمدة! المتوقع 17 عمود');
      for (var m = lastCol; m < 17; m++) {
        output.push('    ← ناقص: عمود ' + (m+1) + ' (' + expectedHeaders[m] + ')');
      }
    } else if (lastCol > 17) {
      output.push('  ⚠ يوجد ' + (lastCol - 17) + ' أعمدة زائدة!');
    } else {
      output.push('  ✓ عدد الأعمدة صحيح (17)');
    }

    if (lastRow < 2) {
      output.push('  ⚠ لا توجد بيانات (هيدرز فقط)');
      continue;
    }

    // ---------- 3ب) قراءة جميع البيانات ----------
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var totalRows = data.length;

    output.push('');
    output.push('─── إحصائيات عامة ───');
    output.push('  إجمالي الصفوف: ' + totalRows);

    // ---------- 3ج) تحليل عدد الأعمدة لكل صف ----------
    var colCountMap = {};
    var shortRows = [];
    for (var r = 0; r < data.length; r++) {
      // حساب آخر عمود غير فارغ في الصف
      var lastFilledCol = 0;
      for (var c = data[r].length - 1; c >= 0; c--) {
        if (data[r][c] !== '' && data[r][c] !== null && data[r][c] !== undefined) {
          lastFilledCol = c + 1;
          break;
        }
      }
      if (!colCountMap[lastFilledCol]) colCountMap[lastFilledCol] = 0;
      colCountMap[lastFilledCol]++;

      if (lastFilledCol < 17 && lastFilledCol > 0) {
        if (shortRows.length < 5) { // أول 5 صفوف قصيرة فقط
          shortRows.push({ row: r + 2, cols: lastFilledCol, name: data[r][1], recorder: data[r][9] });
        }
      }
    }

    output.push('');
    output.push('─── توزيع عدد الأعمدة المملوءة لكل صف ───');
    var colCounts = Object.keys(colCountMap).sort(function(a,b){ return Number(a)-Number(b); });
    for (var cc = 0; cc < colCounts.length; cc++) {
      var cnt = colCounts[cc];
      var status = Number(cnt) === 17 ? ' ✓' : (Number(cnt) < 17 ? ' ⚠ ناقص!' : ' ⚠ زائد!');
      output.push('  ' + cnt + ' أعمدة: ' + colCountMap[cnt] + ' صف' + status);
    }

    if (shortRows.length > 0) {
      output.push('');
      output.push('─── أمثلة صفوف ناقصة الأعمدة ───');
      for (var sr = 0; sr < shortRows.length; sr++) {
        output.push('  صف ' + shortRows[sr].row + ': أعمدة=' + shortRows[sr].cols +
                     ', اسم=[' + shortRows[sr].name + '], مسجل=[' + shortRows[sr].recorder + ']');
      }
    }

    // ---------- 3د) تحليل المصادر (عمود المسجل) ----------
    var recorderMap = {};
    var recorderCol = 9; // index 0-based للعمود 10

    for (var r = 0; r < data.length; r++) {
      var recorder = String(data[r][recorderCol] || '').trim();
      if (!recorder) recorder = '(فارغ)';
      if (!recorderMap[recorder]) {
        recorderMap[recorder] = { count: 0, sampleRow: r + 2 };
      }
      recorderMap[recorder].count++;
    }

    output.push('');
    output.push('─── المصادر (عمود المسجل - عمود 10) ───');
    var recorders = Object.keys(recorderMap).sort(function(a,b){ return recorderMap[b].count - recorderMap[a].count; });
    for (var rc = 0; rc < recorders.length; rc++) {
      var rec = recorders[rc];
      var source = identifySource_(rec);
      output.push('  [' + rec + '] → ' + source + ' | عدد: ' + recorderMap[rec].count + ' صف (مثال: صف ' + recorderMap[rec].sampleRow + ')');
    }

    // ---------- 3هـ) تحليل قيم نوع_الغياب (عمود 6) ----------
    var absTypeMap = {};
    var absTypeCol = 5; // index 0-based

    for (var r = 0; r < data.length; r++) {
      var absType = String(data[r][absTypeCol] || '').trim();
      if (!absType) absType = '(فارغ)';
      if (!absTypeMap[absType]) absTypeMap[absType] = 0;
      absTypeMap[absType]++;
    }

    output.push('');
    output.push('─── قيم نوع_الغياب (عمود 6) ───');
    var absTypes = Object.keys(absTypeMap).sort(function(a,b){ return absTypeMap[b] - absTypeMap[a]; });
    for (var at = 0; at < absTypes.length; at++) {
      var expected = (absTypes[at] === 'يوم كامل' || absTypes[at] === 'حصة') ? ' ✓' : ' ⚠ غير معياري!';
      output.push('  [' + absTypes[at] + ']: ' + absTypeMap[absTypes[at]] + ' صف' + expected);
    }

    // ---------- 3و) تحليل حالة_الاعتماد (عمود 12) ----------
    var statusMap = {};
    var statusCol = 11;
    for (var r = 0; r < data.length; r++) {
      var st = String(data[r][statusCol] || '').trim();
      if (!st) st = '(فارغ)';
      if (!statusMap[st]) statusMap[st] = 0;
      statusMap[st]++;
    }

    output.push('');
    output.push('─── حالة_الاعتماد (عمود 12) ───');
    var statuses = Object.keys(statusMap).sort(function(a,b){ return statusMap[b] - statusMap[a]; });
    for (var st = 0; st < statuses.length; st++) {
      output.push('  [' + statuses[st] + ']: ' + statusMap[statuses[st]] + ' صف');
    }

    // ---------- 3ز) تحليل حالة_التأخر (عمود 15) ----------
    var lateStatusMap = {};
    var lateCol = 14;
    for (var r = 0; r < data.length; r++) {
      var val = (data[r].length > lateCol) ? String(data[r][lateCol] || '').trim() : '(عمود غير موجود)';
      if (!val) val = '(فارغ)';
      if (!lateStatusMap[val]) lateStatusMap[val] = 0;
      lateStatusMap[val]++;
    }

    output.push('');
    output.push('─── حالة_التأخر (عمود 15) ───');
    var lateStatuses = Object.keys(lateStatusMap).sort(function(a,b){ return lateStatusMap[b] - lateStatusMap[a]; });
    for (var ls = 0; ls < lateStatuses.length; ls++) {
      output.push('  [' + lateStatuses[ls] + ']: ' + lateStatusMap[lateStatuses[ls]] + ' صف');
    }

    // ---------- 3ح) تحليل الملاحظات (عمود 17) ----------
    var notesMap = {};
    var notesCol = 16;
    for (var r = 0; r < data.length; r++) {
      var val = (data[r].length > notesCol) ? String(data[r][notesCol] || '').trim() : '(عمود غير موجود)';
      if (!val) val = '(فارغ)';
      // تقصير القيمة الطويلة
      var shortVal = val.length > 40 ? val.substring(0, 40) + '...' : val;
      if (!notesMap[shortVal]) notesMap[shortVal] = 0;
      notesMap[shortVal]++;
    }

    output.push('');
    output.push('─── الملاحظات (عمود 17) ───');
    var notes = Object.keys(notesMap).sort(function(a,b){ return notesMap[b] - notesMap[a]; });
    for (var n = 0; n < Math.min(notes.length, 15); n++) { // أول 15 فقط
      output.push('  [' + notes[n] + ']: ' + notesMap[notes[n]] + ' صف');
    }
    if (notes.length > 15) {
      output.push('  ... و ' + (notes.length - 15) + ' قيم أخرى');
    }

    // ---------- 3ط) عينة تفصيلية (أول 3 صفوف من كل مصدر) ----------
    output.push('');
    output.push('─── عينة تفصيلية (أول صف من كل مصدر) ───');
    var sampledRecorders = {};

    for (var r = 0; r < data.length; r++) {
      var recorder = String(data[r][recorderCol] || '').trim() || '(فارغ)';
      if (sampledRecorders[recorder]) continue;
      sampledRecorders[recorder] = true;

      output.push('');
      output.push('  ◆ مصدر: [' + recorder + '] — صف ' + (r + 2));
      for (var c = 0; c < Math.max(data[r].length, 17); c++) {
        var val = c < data[r].length ? data[r][c] : '(غير موجود)';
        var headerName = c < expectedHeaders.length ? expectedHeaders[c] : 'عمود_' + (c+1);
        var displayVal = '';
        if (val instanceof Date) {
          displayVal = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else {
          displayVal = String(val || '');
          if (displayVal.length > 50) displayVal = displayVal.substring(0, 50) + '...';
        }
        output.push('    عمود ' + (c+1) + ' [' + headerName + ']: ' + displayVal);
      }
    }

    // ---------- 3ي) تحليل التواريخ ----------
    var dateCol = 7; // التاريخ_هجري
    var dateMap = {};
    for (var r = 0; r < data.length; r++) {
      var dateVal = String(data[r][dateCol] || '').trim();
      if (!dateVal) dateVal = '(فارغ)';
      if (!dateMap[dateVal]) dateMap[dateVal] = 0;
      dateMap[dateVal]++;
    }

    output.push('');
    output.push('─── التواريخ الهجرية (عمود 8) ───');
    var dates = Object.keys(dateMap).sort();
    for (var d = 0; d < dates.length; d++) {
      output.push('  [' + dates[d] + ']: ' + dateMap[dates[d]] + ' صف');
    }

    // ---------- 3ك) مطابقة مصدر + عدد أعمدة ----------
    output.push('');
    output.push('─── المصدر × عدد الأعمدة المملوءة ───');
    var crossMap = {};
    for (var r = 0; r < data.length; r++) {
      var recorder = String(data[r][recorderCol] || '').trim() || '(فارغ)';
      var lastFilled = 0;
      for (var c = data[r].length - 1; c >= 0; c--) {
        if (data[r][c] !== '' && data[r][c] !== null && data[r][c] !== undefined) {
          lastFilled = c + 1;
          break;
        }
      }
      var key = recorder + ' → ' + lastFilled + ' أعمدة';
      if (!crossMap[key]) crossMap[key] = 0;
      crossMap[key]++;
    }

    var crossKeys = Object.keys(crossMap).sort(function(a,b){ return crossMap[b] - crossMap[a]; });
    for (var ck = 0; ck < crossKeys.length; ck++) {
      output.push('  ' + crossKeys[ck] + ': ' + crossMap[crossKeys[ck]] + ' صف');
    }
  }

  // ========== 4) ملخص نهائي ==========
  output.push('');
  output.push('╔══════════════════════════════════════════════════════════╗');
  output.push('║                    ملخص التشخيص                         ║');
  output.push('╚══════════════════════════════════════════════════════════╝');
  output.push('');
  output.push('المصادر المتوقعة (6):');
  output.push('  مؤقت 1: مزامنة تلقائية (SyncToApp) → المسجل يحتوي اسم المعلم + ملاحظات=مزامنة تلقائية');
  output.push('  مؤقت 2: استيراد نور → المسجل=مستورد من نور / ملاحظات=استيراد نور');
  output.push('  دائم 3: واجهة الوكيل → المسجل=اسم الوكيل');
  output.push('  دائم 4: واجهة المعلم → المسجل=اسم المعلم');
  output.push('  دائم 5: استيراد منصة → المسجل=استيراد منصة / ملاحظات=منصة');
  output.push('  دائم 6: تسجيل يدوي → المسجل=يدوي');
  output.push('');
  output.push('═══ انتهى التشخيص ═══');

  // طباعة النتيجة
  var fullOutput = output.join('\n');
  Logger.log(fullOutput);

  // حفظ في شيت جديد للسهولة
  var diagSheet = ss.getSheetByName('_تشخيص_الغياب');
  if (diagSheet) ss.deleteSheet(diagSheet);
  diagSheet = ss.insertSheet('_تشخيص_الغياب');
  diagSheet.setRightToLeft(true);
  diagSheet.getRange(1, 1).setValue(fullOutput);
  diagSheet.setColumnWidth(1, 800);
  // تنسيق الخلية
  diagSheet.getRange(1, 1)
    .setFontFamily('Courier New')
    .setFontSize(10)
    .setVerticalAlignment('top')
    .setWrap(true);

  return fullOutput;
}

// =================================================================
// تحديد المصدر من قيمة المسجل
// =================================================================
function identifySource_(recorder) {
  if (!recorder || recorder === '(فارغ)') return '❓ غير معروف';

  var r = recorder.toLowerCase ? recorder.toLowerCase() : String(recorder);

  if (r === 'يدوي' || r === 'manual') return '📝 مصدر 6: تسجيل يدوي';
  if (r.indexOf('مستورد') !== -1 && r.indexOf('نور') !== -1) return '📊 مصدر 2: استيراد نور';
  if (r.indexOf('استيراد') !== -1 && r.indexOf('منصة') !== -1) return '🌐 مصدر 5: استيراد منصة';
  if (r.indexOf('مزامنة') !== -1) return '🔄 مصدر 1: مزامنة تلقائية (SyncToApp)';
  if (r.indexOf('مستورد') !== -1 && r.indexOf('رابط') !== -1) return '🔄 مصدر 1: استيراد من رابط الشيت';
  if (r.indexOf('وكيل') !== -1) return '👤 مصدر 3: واجهة الوكيل';

  // إذا كان اسم شخص (ليس أي من القيم المعروفة أعلاه)
  return '👨‍🏫 مصدر 3/4: معلم أو وكيل (' + recorder + ')';
}

// =================================================================
// إصلاح بيانات الغياب اليومي — شغّل مرة واحدة
// =================================================================
function repairDailyAbsenceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var log = [];
  var totalFixes = 0;

  log.push('╔══════════════════════════════════════════════════════════╗');
  log.push('║       إصلاح بيانات الغياب اليومي                       ║');
  log.push('║       ' + new Date().toLocaleString('ar-SA') + '       ║');
  log.push('╚══════════════════════════════════════════════════════════╝');
  log.push('');

  // أسماء الحصص المعروفة
  var PERIOD_NAMES = ['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة'];
  // أسماء أيام الأسبوع
  var DAY_NAMES = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
  // قيم ملاحظات من SyncToApp المزاحة
  var SHIFTED_NOTES = ['مزامنة تلقائية', 'مستورد من رابط الغياب'];

  for (var si = 0; si < allSheets.length; si++) {
    var sheet = allSheets[si];
    var name = sheet.getName();
    if (name.indexOf('سجل_الغياب_اليومي') === -1) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;

    log.push('═══ إصلاح: ' + name + ' (' + (lastRow - 1) + ' صف) ═══');

    // ضمان 17 عمود في الهيدرز
    var expectedHeaders = [
      'رقم_الطالب', 'اسم_الطالب', 'الصف', 'الفصل', 'رقم_الجوال',
      'نوع_الغياب', 'الحصة', 'التاريخ_هجري', 'اليوم', 'المسجل',
      'وقت_الإدخال', 'حالة_الاعتماد', 'نوع_العذر', 'تم_الإرسال',
      'حالة_التأخر', 'وقت_الحضور', 'ملاحظات'
    ];

    if (lastCol < 17) {
      // إضافة الأعمدة الناقصة في الهيدرز
      for (var h = lastCol; h < 17; h++) {
        sheet.getRange(1, h + 1).setValue(expectedHeaders[h]);
      }
      log.push('  ✓ أضيفت ' + (17 - lastCol) + ' أعمدة ناقصة للهيدرز');
      lastCol = 17;
    }

    // قراءة كل البيانات
    var data = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 17)).getValues();
    var fixes = { shifted: 0, absType: 0, recorder: 0, approval: 0, gradeClass: 0, hijriDate: 0, entryDate: 0 };

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var changed = false;

      // ── 1) إصلاح إزاحة الأعمدة (SyncToApp) ──
      var lateStatus = String(row[14] || '').trim();  // عمود 15 (0-indexed: 14)
      if (SHIFTED_NOTES.indexOf(lateStatus) !== -1) {
        // قيمة الملاحظات موجودة في عمود حالة_التأخر → نقلها
        if (!row[16] || String(row[16]).trim() === '') {
          row[16] = lateStatus;  // نقل إلى ملاحظات (عمود 17)
        }
        row[14] = 'غائب';   // حالة_التأخر
        row[15] = '';         // وقت_الحضور
        fixes.shifted++;
        changed = true;
      }

      // ── 2) توحيد نوع_الغياب (عمود 6, 0-indexed: 5) ──
      var absType = String(row[5] || '').trim();
      var periodCol = String(row[6] || '').trim();  // عمود الحصة

      if (absType && absType !== 'يوم كامل' && absType !== 'حصة') {
        // تحويل 'غياب بدون عذر' و 'غائب' → 'يوم كامل'
        if (absType === 'غياب بدون عذر' || absType === 'غياب بعذر' || absType === 'غائب') {
          row[5] = 'يوم كامل';
          fixes.absType++;
          changed = true;
        }
        // تحويل 'غائب عن حصة' → 'حصة'
        else if (absType === 'غائب عن حصة') {
          row[5] = 'حصة';
          fixes.absType++;
          changed = true;
        }
        // اسم حصة (الأولى، الرابعة...) → 'حصة' + نقل الاسم لعمود الحصة
        else if (PERIOD_NAMES.indexOf(absType) !== -1) {
          row[5] = 'حصة';
          if (!periodCol) row[6] = absType;  // نقل اسم الحصة
          fixes.absType++;
          changed = true;
        }
      }

      // ── 3) إصلاح المسجل = اسم يوم (إزاحة أعمدة) ──
      var recorder = String(row[9] || '').trim();  // عمود 10 (0-indexed: 9)
      if (DAY_NAMES.indexOf(recorder) !== -1) {
        row[9] = 'غير معروف (إزاحة أعمدة)';
        fixes.recorder++;
        changed = true;
      }

      // ── 4) إصلاح timestamp في حالة_الاعتماد ──
      var approval = row[11];  // عمود 12 (0-indexed: 11)
      if (approval instanceof Date) {
        row[11] = 'معلق';
        fixes.approval++;
        changed = true;
      } else {
        var approvalStr = String(approval || '').trim();
        // تحقق إذا هي تاريخ مكتوب كنص
        if (approvalStr.match(/^\d{4}|^[A-Z][a-z]{2}\s|^GMT|^Mon|^Tue|^Wed|^Thu|^Fri|^Sat|^Sun/)) {
          row[11] = 'معلق';
          fixes.approval++;
          changed = true;
        }
      }

      // ── 5) إصلاح الصف/الفصل المقلوبين ──
      var gradeVal = String(row[2] || '').trim();   // عمود 3 (الصف)
      var classVal = String(row[3] || '').trim();   // عمود 4 (الفصل)
      // إذا الفصل يحتوي اسم صف كامل (مثل "الأول متوسط 1")
      if (classVal && classVal.match(/(متوسط|ثانوي|ابتدائي)/) && !gradeVal) {
        var parts = classVal.split(/\s+/);
        if (parts.length > 1) {
          row[2] = parts.slice(0, -1).join(' ');  // الصف
          row[3] = parts[parts.length - 1];        // رقم الفصل
          fixes.gradeClass++;
          changed = true;
        }
      }

      // ── 6) ضمان حالة_التأخر ليست فارغة للصفوف العادية ──
      var studentId = String(row[0] || '').trim();
      var currentLate = String(row[14] || '').trim();
      if (studentId && studentId !== 'NO_ABSENCE' && !currentLate) {
        row[14] = 'غائب';
        changed = true;
      }

      // ── 7) توحيد التاريخ الهجري (Date object أو أرقام عربية ← نص غربي) ──
      var hijriVal = row[7];
      if (hijriVal instanceof Date) {
        // Sheets حوّل النص إلى Date — نستخرج day/month/year
        row[7] = hijriVal.getDate() + '/' + (hijriVal.getMonth() + 1) + '/' + hijriVal.getFullYear();
        fixes.hijriDate++;
        changed = true;
      } else if (hijriVal) {
        var hijriStr = String(hijriVal).trim();
        var westernHijri = repairArabicToWesternNumerals_(hijriStr);
        if (westernHijri !== hijriStr) {
          row[7] = westernHijri;
          fixes.hijriDate++;
          changed = true;
        }
      }

      // ── 8) توحيد وقت_الإدخال (نص ← كائن Date) ──
      var entryTime = row[10];
      if (entryTime && !(entryTime instanceof Date)) {
        var parsedDate = repairParseDateString_(entryTime);
        if (parsedDate) {
          row[10] = parsedDate;
          fixes.entryDate++;
          changed = true;
        }
      }

      if (changed) {
        data[r] = row;
      }
    }

    // ★ تنسيق عمود التاريخ الهجري كنص عادي (يمنع Sheets من تحويله لـ Date)
    sheet.getRange(2, 8, data.length, 1).setNumberFormat('@');

    // كتابة البيانات المصلحة دفعة واحدة
    sheet.getRange(2, 1, data.length, Math.max(data[0].length, 17)).setValues(data);

    var sheetTotal = fixes.shifted + fixes.absType + fixes.recorder + fixes.approval + fixes.gradeClass + fixes.hijriDate + fixes.entryDate;
    totalFixes += sheetTotal;

    log.push('  إزاحة أعمدة (SyncToApp): ' + fixes.shifted + ' صف');
    log.push('  توحيد نوع_الغياب: ' + fixes.absType + ' صف');
    log.push('  إصلاح المسجل: ' + fixes.recorder + ' صف');
    log.push('  إصلاح حالة_الاعتماد: ' + fixes.approval + ' صف');
    log.push('  إصلاح الصف/الفصل: ' + fixes.gradeClass + ' صف');
    log.push('  توحيد التاريخ الهجري: ' + fixes.hijriDate + ' صف');
    log.push('  توحيد وقت_الإدخال: ' + fixes.entryDate + ' صف');
    log.push('  المجموع: ' + sheetTotal + ' إصلاح');
    log.push('');
  }

  log.push('╔══════════════════════════════════════════════════════════╗');
  log.push('║  إجمالي الإصلاحات: ' + totalFixes);
  log.push('╚══════════════════════════════════════════════════════════╝');

  var fullLog = log.join('\n');
  Logger.log(fullLog);

  // حفظ التقرير
  var reportSheet = ss.getSheetByName('_تقرير_الإصلاح');
  if (reportSheet) ss.deleteSheet(reportSheet);
  reportSheet = ss.insertSheet('_تقرير_الإصلاح');
  reportSheet.setRightToLeft(true);
  reportSheet.getRange(1, 1).setValue(fullLog);
  reportSheet.setColumnWidth(1, 800);
  reportSheet.getRange(1, 1)
    .setFontFamily('Courier New')
    .setFontSize(10)
    .setVerticalAlignment('top')
    .setWrap(true);

  return fullLog;
}

// =================================================================
// دوال مساعدة للإصلاح
// =================================================================

/**
 * تحويل الأرقام العربية إلى غربية وإزالة "هـ" والمسافات الزائدة
 * مثال: "٦‏/٩‏/١٤٤٧ هـ" → "6/9/1447"
 */
function repairArabicToWesternNumerals_(str) {
  if (!str) return str;
  var arabicMap = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  var result = String(str);
  for (var k in arabicMap) {
    result = result.split(k).join(arabicMap[k]);
  }
  // إزالة "هـ" وحروف Unicode غير المرئية (LRM, RLM, etc.)
  result = result.replace(/\s*هـ\s*/g, '').replace(/[\u200e\u200f\u200b\u200c\u200d\u2066\u2067\u2068\u2069\u061c]/g, '').trim();
  return result;
}

/**
 * تحويل نص تاريخ إلى كائن Date
 * يدعم: "2026/02/23", "2026-02-23", "2026/02/05 17:02", "Feb 23, 2026"
 */
function repairParseDateString_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  var s = String(value).trim();
  if (!s) return null;

  // صيغة yyyy/MM/dd أو yyyy-MM-dd مع وقت اختياري
  var match = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[\s_T]+(\d{1,2}):(\d{1,2}))?/);
  if (match) {
    var y = parseInt(match[1], 10);
    var m = parseInt(match[2], 10) - 1;
    var d = parseInt(match[3], 10);
    // نحتفظ بالتاريخ فقط بدون وقت (بداية اليوم)
    return new Date(y, m, d);
  }

  // صيغة dd/MM/yyyy
  var match2 = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (match2) {
    var y2 = parseInt(match2[3], 10);
    var m2 = parseInt(match2[2], 10) - 1;
    var d2 = parseInt(match2[1], 10);
    return new Date(y2, m2, d2);
  }

  // محاولة عامة
  try {
    var parsed = new Date(s);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2020) {
      // إرجاع بداية اليوم فقط
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
  } catch (e) {}

  return null;
}
