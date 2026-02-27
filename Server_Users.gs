// =================================================================
// Server_Users.gs - إدارة المستخدمين والصلاحيات
// الإصدار المصحح - بدون ثوابت خارجية
// =================================================================

// =================================================================
// ★ دالة تسجيل التدقيق المركزية — لكل العمليات الحساسة
// =================================================================
function logAuditAction_(action, details, extra) {
  try {
    var ss = getSpreadsheet_();
    var logSheet = ss.getSheetByName('سجل_النشاطات');
    if (!logSheet) {
      logSheet = ss.insertSheet('سجل_النشاطات');
      logSheet.setRightToLeft(true);
      logSheet.appendRow(['التاريخ', 'الوقت', 'المستخدم', 'النوع', 'التفاصيل', 'العدد', 'المرحلة']);
      logSheet.getRange(1, 1, 1, 7).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    var now = new Date();
    var user = '';
    try { user = Session.getActiveUser().getEmail() || 'غير معروف'; } catch(e2) { user = 'نظام'; }
    logSheet.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
      user,
      action,
      details,
      extra || '',
      ''
    ]);
  } catch(e) {
    Logger.log('logAuditAction_ error: ' + e.toString());
  }
}

// =================================================================
// 1. جلب جميع المستخدمين
// =================================================================
function getUsers() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المستخدمين");
    
    if (!sheet) {
      return { success: true, users: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, users: [] };
    }
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      users.push({
        id: String(row[0]),
        name: String(row[1] || ''),
        role: String(row[2] || ''),
        mobile: String(row[3] || ''),
        email: String(row[4] || ''),
        permissions: row[5] ? String(row[5]).split(',').map(p => p.trim()).filter(Boolean) : [],
        scope_type: String(row[6] || 'all'),
        scope_value: String(row[7] || ''),
        status: String(row[8] || 'active')
      });
    }
    
    return { success: true, users: users };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 2. إضافة مستخدم جديد
// =================================================================
function addUser(userData) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName("المستخدمين");
    
    if (!sheet) {
      sheet = ss.insertSheet("المستخدمين");
      sheet.setRightToLeft(true);
      sheet.appendRow(['المعرف', 'الاسم', 'الدور', 'الجوال', 'البريد الإلكتروني', 'الصلاحيات', 'نوع النطاق', 'قيمة النطاق', 'الحالة', 'تاريخ الإنشاء', 'تاريخ التحديث']);
      sheet.getRange(1, 1, 1, 11).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][3]) === String(userData.mobile)) {
        return { success: false, error: "رقم الجوال مسجل مسبقاً" };
      }
    }
    
    const id = 'USER_' + new Date().getTime();
    const now = new Date();
    
    // ★ الصلاحيات لم تعد مستخدمة — كل دور له واجهة محددة تلقائياً

    sheet.appendRow([
      id,
      sanitizeInput_(userData.name),
      sanitizeInput_(userData.role),
      sanitizeInput_(userData.mobile),
      sanitizeInput_(userData.email || ''),
      '',
      sanitizeInput_(userData.scope_type || 'all'),
      sanitizeInput_(userData.scope_value || ''),
      'active',
      now,
      now
    ]);
    
    logAuditAction_('إضافة مستخدم', 'تم إضافة: ' + sanitizeInput_(userData.name) + ' (دور: ' + sanitizeInput_(userData.role) + ')');
    return { success: true, id: id, message: 'تم إضافة المستخدم بنجاح' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 3. تحديث بيانات مستخدم
// =================================================================
function updateUser(userData) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المستخدمين");
    
    if (!sheet) {
      return { success: false, error: "ورقة المستخدمين غير موجودة" };
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: "المستخدم غير موجود" };
    }
    
    // ★ الصلاحيات لم تعد مستخدمة — كل دور له واجهة محددة تلقائياً

    const now = new Date();
    sheet.getRange(rowIndex, 2, 1, 10).setValues([[
      sanitizeInput_(userData.name),
      sanitizeInput_(userData.role),
      sanitizeInput_(userData.mobile),
      sanitizeInput_(userData.email || ''),
      '',
      sanitizeInput_(userData.scope_type || 'all'),
      sanitizeInput_(userData.scope_value || ''),
      userData.status || 'active',
      data[rowIndex - 1][9],
      now
    ]]);
    
    logAuditAction_('تحديث مستخدم', 'تم تحديث: ' + sanitizeInput_(userData.name) + ' (ID: ' + userData.id + ')');
    return { success: true, message: 'تم تحديث بيانات المستخدم بنجاح' };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 4. حذف مستخدم
// =================================================================
function deleteUser(userId) {
  try {
    var authCheck = checkUserPermission('admin');
    if (!authCheck.hasPermission) {
      return { success: false, error: 'غير مصرح: ' + authCheck.reason };
    }
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("المستخدمين");
    
    if (!sheet) {
      return { success: false, error: "ورقة المستخدمين غير موجودة" };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        var deletedName = String(data[i][1] || '');
        sheet.deleteRow(i + 1);
        logAuditAction_('حذف مستخدم', 'تم حذف: ' + deletedName + ' (ID: ' + userId + ')');
        return { success: true, message: 'تم حذف المستخدم بنجاح' };
      }
    }
    
    return { success: false, error: "المستخدم غير موجود" };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 5. جلب الصفوف والفصول للنطاق
// =================================================================
function getScopeOptions() {
  try {
    var sheets = getAllStudentsSheets_();
    var stages = {};
    var grades = {};
    var classes = {};
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s].sheet;
      var stage = sheets[s].stage;
      stages[stage] = true;
      
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = data[0];
      var gradeIdx1 = headers.indexOf('الصف');
      var gradeIdx2 = headers.indexOf('رقم الصف');
      var gradeCol = gradeIdx1 >= 0 ? gradeIdx1 : (gradeIdx2 >= 0 ? gradeIdx2 : 2);
      var classIdx = headers.indexOf('الفصل');
      var classCol = classIdx >= 0 ? classIdx : 3;
      
      for (var i = 1; i < data.length; i++) {
        var grade = cleanGradeName_(data[i][gradeCol]);
        var cls = String(data[i][classCol] || '').trim();
        if (grade) grades[grade] = true;
        if (cls) classes[cls] = true;
      }
    }
    
    return {
      success: true,
      stages: Object.keys(stages),
      grades: Object.keys(grades),
      classes: Object.keys(classes)
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 6. جلب أعضاء لجنة التوجيه والسلوك
// =================================================================
function getCommitteeMembers() {
  try {
    const result = getUsers();
    if (!result.success) {
      return { success: false, error: result.error };
    }
    
    const members = result.users
      .filter(u => u.status === 'active')
      .map(u => ({
        id: u.id,
        name: u.name,
        role: 'عضو',
        jobTitle: u.role
      }));
    
    return { success: true, members: members };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 7. التحقق من صلاحية المستخدم الحالي
// =================================================================
function checkUserPermission(permissionId) {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return { hasPermission: false, reason: 'لم يتم تسجيل الدخول' };
    }

    const result = getUsers();
    if (!result.success) {
      return { hasPermission: false, reason: 'خطأ في جلب المستخدمين' };
    }

    // إذا لم يوجد أي مستخدم في النظام بعد → السماح (أول إعداد)
    if (!result.users || result.users.length === 0) {
      return { hasPermission: true, reason: 'لا يوجد مستخدمين — وضع الإعداد الأولي' };
    }

    const currentUser = result.users.find(u => u.email === email && u.status === 'active');
    if (!currentUser) {
      return { hasPermission: false, reason: 'مستخدم غير مسجل أو غير نشط' };
    }
    
    const hasPermission = currentUser.permissions.includes(permissionId) || 
                          currentUser.permissions.includes('admin');
    
    return {
      hasPermission: hasPermission,
      user: {
        id: currentUser.id,
        name: currentUser.name,
        role: currentUser.role,
        permissions: currentUser.permissions,
        scope_type: currentUser.scope_type,
        scope_value: currentUser.scope_value
      }
    };
    
  } catch (e) {
    Logger.log('checkUserPermission error: ' + e.toString());
    return { hasPermission: false, reason: 'خطأ في فحص الصلاحيات' };
  }
}

// =================================================================
// 8. جلب المستخدم الحالي
// =================================================================
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return { success: false, error: 'لم يتم تسجيل الدخول' };
    }
    
    const result = getUsers();
    if (!result.success) {
      return { success: false, error: result.error };
    }
    
    const currentUser = result.users.find(u => u.email === email);
    if (!currentUser) {
      return { 
        success: true, 
        user: null, 
        message: 'المستخدم غير مسجل في النظام',
        email: email
      };
    }
    
    return { success: true, user: currentUser };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =================================================================
// 9. دوال الاختبار
// =================================================================
function testConnection() {
  return { success: true, message: "الاتصال يعمل!", time: new Date().toString() };
}

function testSpreadsheet() {
  try {
    const ss = getSpreadsheet_();
    return { success: true, name: ss.getName() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}