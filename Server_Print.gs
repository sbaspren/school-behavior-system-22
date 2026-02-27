// =================================================================
// PRINTING SERVICE - خدمة الطباعة
// =================================================================

// 🔥 دالة جلب محتوى ملف الطباعة (لربطه مع index.html)
function getPrintTemplateContent() {
  // دمج ملفين: النماذج + المحرك + المشاركة (لتضمين الكليشة)
  var shared = HtmlService.createHtmlOutputFromFile('JS_PrintShared').getContent();
  var forms = HtmlService.createHtmlOutputFromFile('PrintTemplates_Forms').getContent();
  var engine = HtmlService.createHtmlOutputFromFile('PrintTemplates_Engine').getContent();
  return shared + forms + engine;
}