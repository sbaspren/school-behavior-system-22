# نظام الملف الشامل للطالب - ملف مرجعي لـ Claude
# آخر تحديث: 2026-03-01
# الغرض: يُقرأ في بداية أي محادثة جديدة لاستعادة السياق الكامل

## المشروع
- **الاسم**: نظام الملف الشامل للطالب (Comprehensive Student File System)
- **المطور**: سعيد (معلم ومطور في متوسطة وثانوية العرين بأبها)
- **التقنية**: Google Apps Script (server .gs + client HTML/JS)
- **الحجم**: 47,587 سطر، 65+ ملف
- **الريبو**: https://github.com/sbaspren/school-behavior-system-22
- **السبريدشيت**: https://docs.google.com/spreadsheets/d/1jgtMv8kj7Qrb4WVCuajjKnSGw-5d8mK0CUe1bRRv1q4/edit
- **اللغة**: عربي (RTL)، نظام تعليمي سعودي

## البنية المعمارية

### نقطة الدخول
- `Main.gs` (151 سطر): doGet() يوجّه بمعامل page إلى: teacher, staff, parent, guard, wakeel, counselor, admin, extension

### الإعدادات المركزية
- `Config.gs` (1,443 سطر): 
  - SHEET_REGISTRY: 9 أنواع أوراق (مخالفات، تأخر، استئذان، غياب، غياب_يومي، ملاحظات_تربوية، تواصل، تحصيل، سلوك_إيجابي)
  - NOOR_MAPPINGS: ربط المخالفات (101-710) مع أكواد نور
  - نظام المراحل الديناميكي من ورقة 'هيكل_المدرسة': طفولة مبكرة، ابتدائي، متوسط، ثانوي
  - أوراق الطلاب: طلاب_[المرحلة]
  - النظام الثانوي: فصلي/مقررات

### ملفات السيرفر (22 ملف .gs)
| الملف | الأسطر | الوظيفة |
|-------|--------|---------|
| Server_Data.gs | 371 | getInitialData(), getStudents_(), getRulesData_() |
| Server_Actions.gs | 592 | calculateRepeatLevel(), سجل المخالفات، التخزين المؤقت |
| Server_Absence.gs | 329 | processNoorAbsenceFile(), getAbsenceSheet() |
| Server_Absence_Daily.gs | 1,283 | تتبع الغياب اليومي، 17 عمود، تطبيع، تحقق |
| Server_Academic.gs | 1,053 | التحصيل (ملخص + تفصيلي)، استيراد نور |
| Server_Attendance.gs | 459 | أوراق التأخر والاستئذان |
| Server_Communication.gs | 500 | سجل التواصل |
| Server_Dashboard.gs | 438 | getDashboardData() - إحصائيات مجمّعة |
| Server_EducationalNotes.gs | 358 | الملاحظات التربوية |
| Server_ParentExcuse.gs | 452 | أعذار أولياء الأمور بنظام الرموز المؤقتة |
| Server_Print.gs | 11 | getPrintTemplateContent() |
| Server_Settings.gs | 1,153 | إعدادات المدرسة، الترويسة |
| Server_SMS.gs | 198 | تكامل Madar API |
| Server_StaffInput.gs | 600 | إدخال الموظفين بالرموز الفريدة |
| Server_TeacherInput.gs | 2,015 | نظام إدخال المعلمين (5 خطوات wizard) |
| Server_Templates.gs | 67 | قوالب الرسائل |
| Server_Users.gs | 372 | إدارة المستخدمين، الصلاحيات، سجل النشاطات |
| Server_WhatsApp.gs | 1,274 | واتساب (per_stage/unified)، QR، polling |
| Server noor.gs | 520 | Noor API v3 (مخالفات، تأخر، تعويض، متميز، غياب) |
| Server_Extension.gs | 148 | JSON endpoint للغياب |
| NoorAPI.gs | 15 | ملغي (مدمج في Server noor.gs) |

### ملفات الواجهة (30+ ملف .html)
| الملف | الأسطر | الوظيفة |
|-------|--------|---------|
| index.html | 150 | لوحة التحكم الرئيسية |
| CSS_Styles.html | 499 | نظام التصميم الموحد v4.0 |
| JS_Core.html | 580 | الوظائف الأساسية، التهيئة، التنقل |
| JS_Dashboard.html | 1,140 | لوحة التحكم (تقويم، أحداث، إحصائيات) |
| JS_Violations.html | 2,285 | المخالفات (اليوم/الكل، بطاقات/جدول، التعويضات) |
| JS_Absence.html | 2,878 | الغياب (اليوم/معتمد/أعذار/تقارير، 180 يوم) |
| JS_Attendance.html | 1,355 | التأخر والاستئذان (واجهة مدمجة) |
| JS_Settings.html | 3,322 | الإعدادات (8 تبويبات شاملة) |
| JS_EducationalNotes.html | 1,576 | الملاحظات التربوية |
| JS_GeneralForms.html | 1,336 | نماذج عامة |
| JS_WhatsApp.html | 818 | واجهة واتساب |
| JS_Academic.html | 688 | التحصيل الدراسي |
| JS_PrintShared.html | 422 | طباعة مشتركة مع الترويسة (شعارات Base64) |

### نماذج الإدخال (7 ملفات)
| الملف | الأسطر | المستخدم |
|-------|--------|----------|
| TeacherInputForm.html | 913 | المعلم (wizard 5 خطوات) |
| StaffInputForm.html | 342 | الموظف (تأخر/استئذان) |
| CounselorForm.html | 613 | المرشد (3 تبويبات) |
| WakeelForm.html | 522 | الوكيل (6 تبويبات شاملة) |
| ParentExcuseForm.html | 242 | ولي الأمر (حد 500 حرف) |
| GuardDisplay.html | 178 | الحارس (عرض الاستئذان) |
| AdminTardinessForm.html | 237 | إداري التأخر الصباحي |

### ملفات التشخيص والأدوات
- AuditDatabase.gs (618): تدقيق قاعدة البيانات
- DiagnosticAbsence.gs (648): تشخيص الغياب اليومي (17 عمود)
- ProjectDiag.gs (266): تشخيص شامل
- SeedData.gs (419): بيانات تجريبية
- TestDebug.gs (410): تشخيص الأقسام
- SyncToApp_UPDATED.gs (99): مزامنة الغياب المحدّثة

## الأنظمة الرئيسية

### نظام المخالفات
- 101-109: درجة 1 حضوري | 201-204: درجة 2 | 301-311: درجة 3
- 401-408: درجة 4 | 501-513: درجة 5
- 601-620: رقمي (درجات 1-5) | 701-710: هيئة تعليمية (درجات 4-5)
- نظام التكرار: calculateRepeatLevel() يحسب مستوى التكرار تلقائياً

### نظام الغياب اليومي (17 عمود)
التاريخ_هجري | التاريخ_ميلادي | اسم_الطالب | رقم_الهوية | الصف | الفصل | المرحلة | الحالة | نوع_الغياب | المبرر | مصدر_التسجيل | المسجّل | وقت_التسجيل | حالة_الإشعار | تاريخ_الإشعار | ملاحظات | معرف_فريد

### السلوك الإيجابي (14 نوع)
- 6 درجات: انضباط، خدمة مجتمعية، فعالية حوارية، حملة توعوية، تجارب ناجحة، برنامج/دورة
- 4 درجات: مهارات اتصال/قيادة/رقمية، إدارة وقت
- 2 درجتان: رسالة شكر، إذاعة، مقترح مجتمعي، تعاون

### الملاحظات التربوية
- سلبية (16 نوع): عدم حل الواجب، عدم الحفظ، عدم المشاركة، إلخ
- إيجابية (حسب المرحلة): 4 أنواع لكل مرحلة

### نظام نور (Server noor.gs v3)
- getNoorPendingRecords(stage, type): جلب السجلات غير المرسلة
- updateNoorStatus(): تحديث حالة الإرسال
- الأنواع: violations, tardiness, compensation, excellent, absence

### الواتساب
- وضعان: per_stage (مثيل لكل مرحلة) أو unified (مثيل واحد)
- QR code display، status polling، message queue

### الترويسة الرسمية
- شعار المملكة + شعار الوزارة (Base64 مضمّن)
- وضعان: image (رابط مخصص) أو text (تلقائي)
- 3 أعمدة: شعار يمين | معلومات المدرسة | شعار يسار

## الدوال المساعدة المهمة
- `cleanGradeName_()`: تنظيف أسماء الصفوف
- `detectStageFromGrade_()`: كشف المرحلة من الصف
- `normalizeArabicForMatch_()`: تطبيع النص العربي
- `findNoorMapping_()`: مطابقة ضبابية لأكواد نور
- `getHijriDateFull_()`: تحويل تاريخ هجري
- `escapeHtml()`: حماية XSS في كل الواجهات
- `sanitizeInput_()`: إزالة وسوم HTML
- `validateRowIndex_()`: منع الوصول للصف الأول

## الأمان
- مصادقة بالرموز (tokens) لكل النماذج الخارجية
- صلاحيات حسب النطاق (all/stage/grade/class)
- تنظيف المدخلات + التحقق من URLs (https فقط)
- سجل نشاطات (audit log)

## مشاكل معلّقة
- مشكلة رابط المعلم: يُنشأ ويُرسل بواتساب لكن يظهر خطأ Google Drive عند الفتح من الجوال (مؤجلة)

## ملاحظات للعمل مع سعيد
- يتواصل بالعربي
- يفضل الحلول المفتوحة المصدر
- خبرة عالية بـ Google Apps Script, HTML, JS
- يعمل على إضافات Chrome لنظام نور
- يهتم بالتفاصيل ودقة البيانات
- النظام في الإنتاج (production) لمدرسة حقيقية
