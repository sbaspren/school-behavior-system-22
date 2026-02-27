// ═══════════════════════════════════════════════════════════════
// 🔬 ProjectDiag2.gs — تشخيص شامل محسّن v2
// شغّله: Run → projectDiag2
// النتيجة: شيت جديد في Google Drive + Execution Log
// ═══════════════════════════════════════════════════════════════

function projectDiag2() {
  var O = [];
  O.push('═══════════════════════════════════════════════════════');
  O.push('🔬 تقرير تشخيصي شامل v2 — ' + new Date().toLocaleString('ar-SA'));
  O.push('═══════════════════════════════════════════════════════');

  // ═══ 1. SHEET_REGISTRY كامل ═══
  O.push('\n\n【1】 SHEET_REGISTRY (كامل بدون اقتطاع)');
  O.push('─────────────────────────────────────────');
  try {
    if (typeof SHEET_REGISTRY !== 'undefined') {
      O.push(JSON.stringify(SHEET_REGISTRY, null, 2));
    } else {
      O.push('❌ غير موجود');
    }
  } catch(e) { O.push('خطأ: ' + e.message); }

  // ═══ 2. كل الثوابت العامة ═══
  O.push('\n\n【2】 كل الثوابت والمتغيرات العامة');
  O.push('─────────────────────────────────────────');
  var possibleVars = [
    'CONFIG', 'SETTINGS', 'SHEET_IDS', 'SPREADSHEET_ID', 'SS_ID', 'DB_ID',
    'MID_ID', 'HIGH_ID', 'MIDDLE_SCHOOL', 'HIGH_SCHOOL', 'MIDDLE_SCHOOL_ID',
    'HIGH_SCHOOL_ID', 'MAIN_SS', 'SCHOOL_SS', 'DATA_SS', 'STUDENTS_SS',
    'WHATSAPP_API', 'WA_API', 'API_URL', 'NOOR_API', 'WA_TOKEN', 'WA_INSTANCE',
    'SCHOOL_NAME', 'APP_CONFIG', 'STAGES', 'ENABLED_STAGES',
    'GRADE_MAP', 'CLASS_MAP', 'SECTION_MAP', 'CURRENT_SEMESTER',
    'SS_REGISTRY', 'DB_REGISTRY', 'SHEET_MAP', 'LETTERHEAD', 'LOGO'
  ];
  possibleVars.forEach(function(v) {
    try {
      var val = eval(v);
      if (val !== undefined) {
        var str = typeof val === 'object' ? JSON.stringify(val) : String(val);
        O.push('  ✅ ' + v + ' [' + typeof val + '] = ' + str.substring(0, 500));
      }
    } catch(e) {}
  });

  // ═══ 3. البحث عن Spreadsheet IDs في كل مكان ═══
  O.push('\n\n【3】 البحث عن IDs الشيتات');
  O.push('─────────────────────────────────────────');
  
  var foundIds = {};
  
  // من SHEET_REGISTRY
  try {
    if (typeof SHEET_REGISTRY !== 'undefined') {
      for (var k in SHEET_REGISTRY) {
        var entry = SHEET_REGISTRY[k];
        if (typeof entry === 'string' && entry.length > 20) {
          foundIds[k] = entry;
        } else if (typeof entry === 'object' && entry !== null) {
          // البحث عن id داخل الكائن
          ['id', 'ssId', 'sheetId', 'spreadsheetId', 'ss', 'db'].forEach(function(prop) {
            if (entry[prop] && typeof entry[prop] === 'string' && entry[prop].length > 20) {
              foundIds[k + '.' + prop] = entry[prop];
            }
          });
        }
      }
    }
  } catch(e) {}
  
  // من CONFIG
  try {
    if (typeof CONFIG !== 'undefined') {
      var configStr = JSON.stringify(CONFIG);
      var matches = configStr.match(/[a-zA-Z0-9_-]{30,}/g);
      if (matches) matches.forEach(function(m, i) { foundIds['CONFIG_match_' + i] = m; });
    }
  } catch(e) {}

  // من Script Properties
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    for (var pk in props) {
      if (props[pk].length > 25 && /^[a-zA-Z0-9_-]+$/.test(props[pk])) {
        foundIds['prop_' + pk] = props[pk];
      }
    }
  } catch(e) {}
  
  // من this (global scope)
  try {
    var globalFn = new Function('var ids={}; try { for(var k in this) { try { var v=this[k]; if(typeof v==="string" && v.length>25 && /^[a-zA-Z0-9_-]+$/.test(v)) ids[k]=v; } catch(e){} } } catch(e){} return ids;');
    var globalIds = globalFn();
    for (var gk in globalIds) foundIds['global_' + gk] = globalIds[gk];
  } catch(e) {}

  for (var fk in foundIds) {
    O.push('  📌 ' + fk + ' = ' + foundIds[fk]);
  }

  // ═══ 4. محاولة فتح Active Spreadsheet ═══
  O.push('\n\n【4】 Active Spreadsheet');
  O.push('─────────────────────────────────────────');
  try {
    var activeSS = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSS) {
      O.push('  ✅ ID: ' + activeSS.getId());
      O.push('  📋 اسم: ' + activeSS.getName());
      listSheets_(activeSS, O);
    } else {
      O.push('  ❌ لا يوجد Active Spreadsheet');
    }
  } catch(e) {
    O.push('  ❌ ' + e.message);
  }

  // ═══ 5. فتح كل شيت وُجد ═══
  O.push('\n\n【5】 تفاصيل كل الشيتات المكتشفة');
  O.push('─────────────────────────────────────────');
  var openedIds = {};
  for (var fid in foundIds) {
    var ssId = foundIds[fid];
    if (openedIds[ssId]) continue;
    openedIds[ssId] = true;
    
    O.push('\n  📗 ' + fid + ': ' + ssId);
    try {
      var ss = SpreadsheetApp.openById(ssId);
      O.push('     اسم: ' + ss.getName());
      listSheets_(ss, O);
    } catch(e) {
      O.push('     ⚠️ فشل الفتح: ' + e.message);
    }
  }

  // ═══ 6. كل الدوال الموجودة فعلاً ═══
  O.push('\n\n【6】 كل الدوال المكتشفة (الموجودة)');
  O.push('─────────────────────────────────────────');
  
  var allFns = [
    'doGet','doPost','include','getInitialData','onOpen','onEdit','onInstall',
    'getStudents','getStudentData','getStudentInfo','getStudentRecord','searchStudents','findStudent',
    'getViolations','addViolation','saveViolation','deleteViolation','removeViolation','getViolationRecords',
    'getAbsenceRecords','addAbsence','saveAbsence','deleteAbsence','recordAbsence','markAbsence','getAbsenceData','saveAbsenceRecords',
    'getTardinessRecords','addTardiness','saveTardiness','recordTardiness','getTardinessData',
    'getPermissions','addPermission','savePermission','getPermissionRecords',
    'getSettings','saveSettings','getSchoolStructure','saveSchoolStructure','getConfig','saveConfig',
    'sendWhatsApp','sendWhatsAppMessage','sendWAMessage','sendMessage','sendBulkWhatsApp',
    'getNoorPendingRecords','updateNoorStatus','getNoorStats','noorGetPending',
    'generateForm','printForm','getFormData','createForm','generatePDF',
    'getDashboardData','getDashboardStats','getReport','getStatistics',
    'getEducationalNotes','addEducationalNote','saveNote','getNotes',
    'getCommunicationRecords','addCommunication','saveCommunication',
    'getAcademicData','getGrades','saveGrades',
    'handleFormSubmit','processForm','onFormSubmit',
    'getSheetData','saveToSheet','appendRow','updateRow','deleteRow',
    'getStageData','getStageStudents','filterByStage',
    'createTrigger','deleteTrigger','setupTriggers',
    'backup','restore','exportData','importData',
    'testConnection','debugSheetNames','projectDiagnostic','projectDiag2',
    'doGetNoor','doGet_noor','handleNoorRequest'
  ];

  var existingFns = [];
  allFns.forEach(function(fn) {
    try {
      if (typeof eval(fn) === 'function') existingFns.push(fn);
    } catch(e) {}
  });
  
  O.push('  عدد الدوال المكتشفة: ' + existingFns.length);
  existingFns.forEach(function(fn) { O.push('  ✅ ' + fn + '()'); });

  // ═══ 7. Script Properties كاملة ═══
  O.push('\n\n【7】 Script Properties (كاملة)');
  O.push('─────────────────────────────────────────');
  try {
    var allProps = PropertiesService.getScriptProperties().getProperties();
    var pKeys = Object.keys(allProps);
    O.push('  عدد المفاتيح: ' + pKeys.length);
    pKeys.forEach(function(pk) {
      var pv = allProps[pk];
      // إخفاء tokens فقط
      if (/token|secret|password|api_key/i.test(pk)) pv = pv.substring(0, 10) + '***';
      O.push('  🔑 ' + pk + ' = ' + String(pv).substring(0, 300));
    });
  } catch(e) { O.push('  خطأ: ' + e.message); }

  // ═══ 8. User Properties ═══
  O.push('\n\n【8】 User Properties');
  O.push('─────────────────────────────────────────');
  try {
    var uProps = PropertiesService.getUserProperties().getProperties();
    var uKeys = Object.keys(uProps);
    O.push('  عدد المفاتيح: ' + uKeys.length);
    uKeys.forEach(function(uk) {
      O.push('  🔑 ' + uk + ' = ' + String(uProps[uk]).substring(0, 300));
    });
  } catch(e) { O.push('  خطأ: ' + e.message); }

  // ═══ 9. Triggers ═══
  O.push('\n\n【9】 المشغّلات (Triggers)');
  O.push('─────────────────────────────────────────');
  try {
    var triggers = ScriptApp.getProjectTriggers();
    O.push('  عدد: ' + triggers.length);
    triggers.forEach(function(t) {
      O.push('  ⏰ ' + t.getHandlerFunction() + ' | ' + t.getEventType() + ' | source: ' + t.getTriggerSource());
    });
  } catch(e) {}

  // ═══ 10. وصف doGet ═══
  O.push('\n\n【10】 كود doGet (أول 500 حرف)');
  O.push('─────────────────────────────────────────');
  try {
    if (typeof doGet === 'function') {
      O.push(doGet.toString().substring(0, 500));
    }
  } catch(e) {}

  // ═══ حفظ ═══
  var report = O.join('\n');
  Logger.log(report);

  // حفظ في شيت
  try {
    var tempSS = SpreadsheetApp.create('🔬 تقرير v2 — ' + new Date().toLocaleDateString('ar-SA'));
    var sheet = tempSS.getActiveSheet();
    // تقسيم على أسطر في خلايا منفصلة لسهولة القراءة
    var lines = report.split('\n');
    var data = lines.map(function(l) { return [l]; });
    sheet.getRange(1, 1, data.length, 1).setValues(data);
    sheet.setColumnWidth(1, 1000);
    Logger.log('\n📄 التقرير في: ' + tempSS.getUrl());
  } catch(e) {
    Logger.log('فشل حفظ الشيت: ' + e.message);
  }

  return report;
}

// ═══ دالة مساعدة: عرض أوراق الشيت ═══
function listSheets_(ss, O) {
  var sheets = ss.getSheets();
  O.push('     عدد الأوراق: ' + sheets.length);
  sheets.forEach(function(s) {
    var name = s.getName();
    var rows = s.getLastRow();
    var cols = s.getLastColumn();
    O.push('     📋 "' + name + '" — ' + rows + ' صف × ' + cols + ' عمود');
    if (rows > 0 && cols > 0) {
      try {
        var headers = s.getRange(1, 1, 1, Math.min(cols, 25)).getValues()[0]
          .filter(function(h) { return h !== ''; }).join(' | ');
        O.push('        أعمدة: ' + headers);
      } catch(e) {}
      // عينة من أول صف بيانات
      if (rows > 1) {
        try {
          var sample = s.getRange(2, 1, 1, Math.min(cols, 15)).getValues()[0]
            .map(function(v) { return String(v).substring(0, 30); }).join(' | ');
          O.push('        عينة: ' + sample);
        } catch(e) {}
      }
    }
  });
}