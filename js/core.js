'use strict';
/* ══════════════════════════════════════════════════════════════
   brivatedocs — Core Utilities, Theme, i18n, DraftDB & Pkg
   ══════════════════════════════════════════════════════════════ */
function safeUrl(u) {
  const s = String(u || '').trim();
  if (s.startsWith('#')) return s;
  try {
    const candidate = /^[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}$/.test(s) ? `mailto:${s}`
      : /^[a-z][a-z0-9+.-]*:/i.test(s) ? s : `https://${s}`;
    const p = new URL(candidate);
    return ['http:', 'https:', 'mailto:'].includes(p.protocol) ? candidate : '#';
  } catch (_) { return '#'; }
}

function escHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function escXml(s) { return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

const MAX_FILE_BYTES = 100 * 1024 * 1024;
function assertFileSize(file) {
  if (file.size > MAX_FILE_BYTES) throw new UserError('الملف كبير جداً (الحد 100MB)');
}


/* ───────────────────────── ثوابت وأدوات ───────────────────────── */
const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  svg: 'http://www.w3.org/2000/svg'
};
const A = NS.a, P = NS.p, R = NS.r, W = NS.w, C = NS.c;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const CT_NOTES = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));

/* ── الوضع الليلي / النهاري (Dark Mode) ── */
let currentTheme = 'light';
function applyTheme(theme) {
  currentTheme = theme === 'dark' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', currentTheme);
  try { localStorage.setItem('doc-studio-theme', currentTheme); } catch (_) {}
  const icon = currentTheme === 'dark' ? '☀️' : '🌙';
  for (const b of [$('#themeToggleHome'), $('#themeToggleEditor')]) {
    if (b) b.textContent = icon;
  }
}

function setupTheme() {
  let saved = 'light';
  try { saved = localStorage.getItem('doc-studio-theme') || 'light'; } catch (_) {}
  applyTheme(saved);
  const toggle = () => applyTheme(currentTheme === 'dark' ? 'light' : 'dark');
  $('#themeToggleHome')?.addEventListener('click', toggle);
  $('#themeToggleEditor')?.addEventListener('click', toggle);
}

/* ── واجهة ثنائية اللغة (إنجليزي افتراضي / عربي فصيح) ── */
const translations = {
  'أدوات سريعة': 'Quick Tools',
  'بدون رفع لسيرفر، بدون حساب، وبلا حدود.': 'Zero uploads, zero accounts, zero limits.',
  'تحويل الصور إلى PDF': 'Images → PDF',
  'تقسيم PDF': 'Split PDF',
  'ادمج، قسّم، اضغط، احمِ، قارن المستندات، استخرج الصور، ونظّف البيانات المخفية — كل المعالجة تتم على جهازك.': 'Merge, split, compress, protect, compare docs, extract images, and clean hidden metadata — all processed locally.',
  '🗜 ضغط PDF': '🗜 Compress PDF',
  'ضغط PDF': 'Compress PDF',
  '🔒 حماية PDF بكلمة مرور': '🔒 Protect PDF',
  'حماية PDF بكلمة مرور': 'Protect PDF',
  '🔓 فك حماية PDF': '🔓 Unlock PDF',
  'فك حماية PDF': 'Unlock PDF',
  '📝 تعبئة نموذج PDF': '📝 Fill PDF Form',
  'تعبئة نموذج PDF': 'Fill PDF Form',
  'إنشاء عرض جديد': 'Create New Deck',
  'إنشاء مستند Word': 'New Word Doc',
  'أو ابدأ من الصفر:': 'Or start from scratch:',

  'اللغة': 'Language',
  'للتواصل': 'Contact',
  '↩ تراجع': '↩ Undo',
  '↪ إعادة': '↪ Redo',
  'جهات الاتصال': 'For contacts',
  '📝 محرر المستندات': '📝 Document Editor',
  'محرر المستندات': 'Document Editor',
  'حرّر مستنداتك. وخلّها على جهازك.': 'Edit documents. Keep them yours.',
  'افتح أو أنشئ ملفات Word و PowerPoint و PDF داخل متصفحك مباشرة.': 'Open or create Word, PowerPoint, and PDF files right in your browser.',
  'كل شيء يتم داخل متصفحك، ولا يغادر ملفك جهازك.': 'Everything happens inside your browser, and your file never leaves your device.',
  'اسحب الملف هنا أو اضغط للاختيار': 'Drag a file here or click to browse',
  'مستند Word (.docx) أو عرض PowerPoint (.pptx)': 'Word (.docx) or PowerPoint (.pptx)',
  '✨ أو ابدأ من الصفر:': '✨ Or start from scratch:',
  'قالب كلاسيكي (أبيض)': 'Classic Theme (White)',
  'قالب ليلي (داكن)': 'Night Theme (Dark)',
  'قالب رسمي (أزرق ملكي)': 'Formal Theme (Royal Blue)',
  'قالب دافئ (رملي)': 'Warm Theme (Sand)',
  'قالب زمردي (أخضر داكن)': 'Emerald Theme (Forest Green)',
  'قالب تنفيذي (بنفسجي داكن)': 'Executive Theme (Deep Purple)',
  'قالب ستوديو (تيراكوتا)': 'Studio Theme (Terracotta)',
  'قالب عصري (رمادي حديث)': 'Modern Theme (Cool Slate)',
  'قالب محضر اجتماع': 'Meeting Minutes Template',
  'قالب مقترح مشروع': 'Project Proposal Template',
  'قالب اتفاقية عمل': 'Service Agreement Template',
  'قالب ملخص دراسي': 'Study Notes Template',
  '➕ إنشاء عرض جديد': '➕ Create New Deck',
  '📄 إنشاء مستند Word': '📄 New Word Doc',
  'تحرير محلي وآمن': 'Private, local editing',
  'تبقى ملفاتك داخل المتصفح وعلى جهازك': 'Your files remain inside the browser and on your device',
  '🖼️ تحويل الصور إلى PDF': '🖼️ Images to PDF',
  '✂️ تقسيم PDF': '✂️ Split PDF',
  'Word و PowerPoint و PDF': 'Word, PowerPoint & PDF',
  'تحرير متكامل للمستندات والعروض التقديمية وصفحات PDF محلياً': 'Edit documents, slide decks, and PDF pages locally',
  'حفظ مباشر وعمل أوفلاين': 'Direct Save & Offline Ready',
  'احفظ تعديلاتك مباشرة على الملف الأصلي واستخدم الموقع حتى بدون إنترنت': 'Save directly back to your file and work offline anytime',
  'طريقة الاستخدام': 'How it works',
  '١. اختر ملفك أو أنشئ مستنداً ← ٢. عدّل ← ٣. حمّل النسخة الجديدة': '1. Choose or create a file → 2. Edit → 3. Download the new copy',
   '⬅ الرئيسية': '⬅ Home',
  'تحرير محلي': 'Local editing',
  '🔎 بحث واستبدال': '🔎 Find & Replace',
  '🎯 وضع التركيز': '🎯 Focus Mode',
  '🕒 النسخ': '🕒 Versions',
  '⬇ تحميل الملف': '⬇ Download file',
  'جارٍ العمل…': 'Working…',
  'يجب تفعيل جافاسكربت لكي يعمل الموقع.': 'JavaScript must be enabled for this website to work.',
  'الوضع الليلي / النهاري': 'Light / Dark Mode',
  'تبديل المظهر': 'Toggle Theme',
  'توجد مسودة غير محفوظة، هل ترغب في استكمالها؟': 'An unsaved draft was found. Would you like to resume it?',
  '↩ استعادة المسودة': '↩ Restore Draft',
  '🗑 تجاهل': '🗑 Discard',
  'تمت استعادة المسودة بنجاح': 'Draft restored successfully',
  'تم تجاهل المسودة السابقة': 'Previous draft discarded',

  /* التبويبات وشريط الأدوات PPTX */
  'الرئيسية': 'Home',
  'إدراج': 'Insert',
  'تصميم': 'Design',
  'عرض وتصدير': 'Show & Export',
  '➕ شريحة': '➕ New slide',
  '📄 نسخ': '📄 Duplicate',
  '🗑 حذف شريحة': '🗑 Delete slide',
  '⬆️ للأعلى': '⬆️ Move up',
  '⬇️ للأسفل': '⬇️ Move down',
  'التخطيط:': 'Layout:',
  'مقاس الشريحة:': 'Slide Size:',
  '16:9 شاشة عريضة': '16:9 Widescreen',
  '4:3 قياسي': '4:3 Standard',
  'تم تغيير مقاس الشريحة': 'Slide aspect ratio updated',
  '👁 إخفاء/إظهار': '👁 Hide/Show',
  'إخفاء': 'Hide',
  'إظهار': 'Show',
  'مخفية': 'Hidden',
  '⊞ عرض الشبكة': '⊞ Grid View',
  '📄 عرض عادي': '📄 Normal View',
  '➕ مربع نص': '➕ Text box',
  '🖼 صورة': '🖼 Image',
  '🖼+ صور لشرائح': '🖼+ Images to Slides',
  '✨ أيقونات': '✨ Icons',
  'مكتبة الأيقونات': 'Icon Library',
  'صح': 'Check',
  'خطأ': 'Cross',
  'نجمة': 'Star',
  'موقع': 'Location',
  'هاتف': 'Phone',
  'بريد': 'Email',
  'تنبيه': 'Alert',
  'قلب': 'Heart',
  'ترس': 'Gear',
  'سهم': 'Arrow',
  '▦ جدول': '▦ Table',
  '📊 رسم بياني': '📊 Chart',
  '🎬 فيديو/صوت': '🎬 Video/Audio',
  '▭ مستطيل': '▭ Rectangle',
  '◯ دائرة': '◯ Circle',
  '▲ مثلث': '▲ Triangle',
  '➡ سهم': '➡ Arrow',
  '╱ خط': '╱ Line',
  '📋 نسخ عنصر': '📋 Copy',
  '📥 لصق': '📥 Paste',
  '📄 تكرار عنصر': '📄 Duplicate Object',
  '🖌 نسخ التنسيق': '🖌 Format Painter',
  '🗑 حذف عنصر': '🗑 Delete',
  '⬆ للأمام': '⬆ Bring front',
  '⬇ للخلف': '⬇ Send back',
  '🔗 تجميع': '🔗 Group',
  '🔓 فك التجميع': '🔓 Ungroup',
  '📐 شبكة التصميم': '📐 Canvas Grid',
  '🎨 لون الخلفية': '🎨 Background',
  '🌈 تدرج الخلفية': '🌈 BG Gradient',
  'نوع التعبئة:': 'Fill Type:',
  'لون ثابت': 'Solid Color',
  'تدرج لوني': 'Gradient Fill',
  'لون 1': 'Color 1',
  'لون 2': 'Color 2',
  'الزاوية (°)': 'Angle (°)',
  '💉 قطارة الألوان': '💉 Eyedropper',
  'انقر على أي نقطة في الشريحة لالتقاط لونها': 'Click anywhere on the slide to pick a color',
  'تم التقاط اللون': 'Color picked',
  '🖼 صورة خلفية': '🖼 BG Image',
  'تطبيق على الكل': 'Apply to all',
  '🔢 ترقيم وتذييل الكل': '🔢 Numbers & Footer',
  'انتقال الشريحة:': 'Transition:',
  'بدون انتقال': 'No transition',
  'تلاشي (Fade)': 'Fade',
  'دفع (Push)': 'Push',
  'مسح (Wipe)': 'Wipe',
  'تكبير (Zoom)': 'Zoom transition',
  '▶ بدء العرض': '▶ Start Slideshow',
  '🖥 وضع مقدم العرض': '🖥 Presenter View',
  '🖼 تصدير PNG': '🖼 Export PNG',
  '📦 تصدير الكل ZIP': '📦 Export All ZIP',
  '📄 تصدير PDF': '📄 Export PDF',
  'تكبير': 'Zoom',
  '◀ السابق': '◀ Previous',
  'التالي ▶': 'Next ▶',
  '📝 ملاحظات المحاضر (Speaker Notes)': '📝 Speaker Notes',
  'اكتب ملاحظاتك لهذه الشريحة هنا…': 'Write your speaker notes for this slide here…',

  /* أدوات محرر Word (DOCX) */
  'نص عادي': 'Normal text',
  'عنوان 1': 'Heading 1',
  'عنوان 2': 'Heading 2',
  'عنوان 3': 'Heading 3',
  'يتوسطه خط': 'Strikethrough',
  'نص مرتفع': 'Superscript',
  'نص منخفض': 'Subscript',
  'تظليل النص': 'Highlight Color',
  'مسح التنسيق': 'Clear Formatting',
  'قائمة نقطية': 'Bulleted List',
  'قائمة رقمية': 'Numbered List',
  '• قائمة': '• Bullets',
  '1. ترقيم': '1. Numbering',
  'زيادة المسافة البادئة': 'Increase Indent',
  'إنقاص المسافة البادئة': 'Decrease Indent',
  '🔗 رابط': '🔗 Link',
  '📄 فاصل صفحة': '📄 Page Break',
  'فاصل صفحة': 'Page Break',
  'رأس وتذييل': 'Header & Footer',
  'نص رأس الصفحة…': 'Header text…',
  'نص تذييل الصفحة…': 'Footer text…',
  'ترقيم الصفحات': 'Page Numbers',
  'حفظ الرأس والتذييل': 'Save Header & Footer',
  'تم حفظ الرأس والتذييل': 'Header & Footer saved',
  'حجم الورقة:': 'Page Size:',
  'عمودي': 'Portrait',
  'أفقي': 'Landscape',
  'هوامش عادية': 'Normal Margins',
  'هوامش ضيقة': 'Narrow Margins',
  'هوامش واسعة': 'Wide Margins',
  'أدوات الجدول:': 'Table Tools:',
  'إخفاء/إظهار الحدود': 'Toggle Borders',
  'أدوات الصورة:': 'Image Tools:',
  'صغير 25%': 'Small 25%',
  'متوسط 50%': 'Medium 50%',
  'كبير 75%': 'Large 75%',
  'كامل 100%': 'Full 100%',
  'حذف الصورة': 'Delete Image',
  'الكلمات:': 'Words:',
  'الأحرف:': 'Chars:',
  'الفقرات:': 'Paragraphs:',
  'الصفحات:': 'Pages:',
  'تم إنشاء مستند Word بنجاح': 'Word document created successfully',
  'تم إدراج الرابط بنجاح': 'Hyperlink inserted successfully',
  'حدد نصاً أولاً لإضافة رابط': 'Select text first to add a link',
  'المشاريع المحفوظة مؤخراً:': 'Recent Saved Projects:',
  '🗑 مسح الكل': '🗑 Clear All',
  'مستند Word فارغ': 'Blank Word Doc',
  'قالب سيرة ذاتية (CV)': 'Resume / CV Template',
  'قالب خطاب رسمي': 'Formal Letter Template',
  'قالب تقرير جامعي': 'Academic Report Template',
  'قالب فاتورة': 'Invoice Template',
  '📑 فهرس تلقائي': '📑 Auto TOC',
  'فهرس المحتويات': 'Table of Contents',
  'لا توجد عناوين (Heading 1-3) لبناء الفهرس': 'No headings (Heading 1-3) found for TOC',
  'تم إدراج فهرس المحتويات': 'Table of Contents inserted',
  '💬 صندوق تنبيه': '💬 Callout Box',
  '➖ خط فاصل': '➖ Horizontal Line',
  'لون الورقة': 'Page Color',
  'إطار الصفحة': 'Page Border',
  'علامة مائية': 'Watermark',
  'اكتب نص العلامة المائية (اتركه فارغاً للإزالة):': 'Enter watermark text (leave empty to remove):',
  'تم تحديث العلامة المائية': 'Watermark updated',
  'دمج مع الخلية التالية': 'Merge Right Cell',
  'توسيط عمودي': 'Align Middle',
  'المحدد:': 'Selected:',
  'وقت القراءة:': 'Read time:',
  'دقيقة': 'min',
  '📝 استيراد Markdown': '📝 Import Markdown',
  'الصق نص Markdown هنا لتحويله تلقائياً إلى محتوى منسق:': 'Paste Markdown text here to convert into formatted content:',
  'تحويل وإدراج': 'Convert & Insert',
  'إلغاء': 'Cancel',
  'تم استيراد Markdown بنجاح': 'Markdown imported successfully',
  '🔄 تحويل إلى عرض PPTX': '🔄 Convert to PPTX',
  '📄 تصدير إلى Word': '📄 Export to Word',
  'تم تحويل المستند إلى عرض تقديمي': 'Document converted to presentation',
  'تم تصدير العرض إلى مستند Word': 'Presentation exported to Word document',
  '✨ شرائح ذكية': '✨ Smart Slides',
  'مكتبة الشرائح الذكية': 'Smart Slide Templates',
  'مقارنة عمودين': 'Two-Column Comparison',
  'أرقام وإحصائيات': 'Stats & KPIs',
  'خط زمني': 'Timeline',
  'فريق العمل': 'Team Members',
  'تم إدراج الشريحة الذكية': 'Smart slide inserted',
  '◆ معين': '◆ Diamond',
  '⬡ سداسي': '⬡ Hexagon',
  '★ نجمة': '★ Star',
  '💬 فقاعة': '💬 Callout',
  'استدارة الزوايا (%)': 'Corner Radius (%)',
  '🔒 قفل/فتح العنصر': '🔒 Lock/Unlock',
  'تم قفل العنصر': 'Object locked',
  'تم فك قفل العنصر': 'Object unlocked',
  '➕ سلسلة مقارنة': '➕ Add Series',
  'السلسلة 1': 'Series 1',
  'السلسلة 2': 'Series 2',
  'تشغيل تلقائي': 'Autoplay',
  '✏️ قلم': '✏️ Pen',
  '🧹 مسح': '🧹 Clear',
  'البحث والاستبدال': 'Find & Replace',
  'استبدال الكل': 'Replace All',
  'النص المطلوب البحث عنه': 'Text to find',
  'النص البديل': 'Replace with',
  'تم تصدير ملف PDF بنجاح': 'PDF exported successfully',
  'إغلاق': 'Close',
  '✏️ تعديل نص الملف': '✏️ Edit PDF Text',
  'استعادة النص الأصلي': 'Restore original text',
  'هذه الصفحة مصوّرة ولا تحتوي على نص أصلي قابل للتعديل': 'This page is scanned and contains no editable original text',
  'انقر على أي نص في الصفحة لتعديله مباشرة': 'Click any text on the page to edit it directly',
  '🔄 تحويل إلى Word': '🔄 Convert to Word',
  '🔄 تحويل إلى PDF': '🔄 Convert to PDF',
  'تم تحويل ملف PDF إلى مستند Word بنجاح': 'PDF converted to Word document successfully',
  'تم تحويل المستند إلى محرر PDF': 'Document opened in PDF editor',
  'جارٍ تحويل ملف PDF إلى Word…': 'Converting PDF to Word…',
  'تم تحديث نص الـ PDF بنجاح': 'PDF text updated',
  '✒️ توقيع رقمي': '✒️ Digital Sign',
  'إضافة توقيع رقمي / إلكتروني': 'Add Digital / Electronic Signature',
  '✍️ رسم التوقيع': '✍️ Draw Signature',
  '⌨️ كتابة التوقيع': '⌨️ Type Signature',
  '📤 رفع صورة / ختم': '📤 Upload Image / Stamp',
  'ارسم توقيعك بالماوس أو اللمس هنا': 'Draw your signature here with mouse or touch',
  'لون الحبر:': 'Ink Color:',
  'سمك الحبر:': 'Stroke Width:',
  '🧹 مسح التوقيع': '🧹 Clear Pad',
  'اكتب اسمك للتوقيع…': 'Type your name to sign…',
  'نمط الخط:': 'Signature Style:',
  'توقيع انسيابي': 'Cursive Script',
  'كلاسيكي رسمي': 'Classic Serif',
  'عصري عريض': 'Modern Bold',
  'خط رفيع مائل': 'Monoline Italic',
  'اختر صورة توقيع أو ختم رسمي': 'Choose signature or stamp image',
  'تفريغ الخلفية البيضاء تلقائياً': 'Auto-remove white background',
  'إرفاق ختم التوثيق الرقمي (الاسم + التاريخ + SHA-256)': 'Attach Digital Verification Seal (Name + Date + SHA-256)',
  'اسم الموقّع:': 'Signer Name:',
  'الصفة أو السبب (اختياري):': 'Title or Reason (optional):',
  'معتمد وموقّع إلكترونياً': 'Digitally Signed & Verified',
  'حفظ هذا التوقيع للاستخدام لاحقاً': 'Save this signature for future use',
  'التواقيع المحفوظة محلياً:': 'Locally Saved Signatures:',
  'إدراج التوقيع في الصفحة': 'Insert Signature on Page',
  'يرجى رسم التوقيع أو كتابته أو رفع صورة أولاً': 'Please draw, type, or upload a signature first',
  'تم إدراج التوقيع بنجاح — اسحبه أو غيّر حجمه بحرية': 'Signature inserted — drag or resize freely',
  'تم حذف التوقيع': 'Signature removed',
  '📄 تعميم على كل الصفحات': '📄 Apply to All Pages',
  'تم تعميم التوقيع على جميع الصفحات': 'Signature applied to all pages',
  'حذف التوقيع المحفوظ': 'Delete saved signature',
  '🗑 حذف الجدول': '🗑 Delete Table',
  'تم حذف الجدول': 'Table deleted',
  'حذف': 'Delete',
  'رفيع': 'Fine',
  'متوسط': 'Medium',
  '⏱ تشغيل تلقائي': '⏱ Autoplay',
  'ملائمة العرض': 'Fit Width',
  '🖍 تظليل': '🖍 Highlight',
  '▭ طمس/تغطية': '▭ Whiteout',
  '🖼 صورة PDF': '🖼 Add Image',
  '📄 صفحة فارغة': '📄 Blank Page',
  '✂️ استخراج صفحات': '✂️ Extract Pages',
  'تمت إضافة صفحة فارغة': 'Blank page added',
  'أدخل رقم الصفحة للانتقال إليها:': 'Enter page number to jump to:',
  'انقر للانتقال السريع إلى صفحة': 'Click to jump to page',
  'ابحث في صفحات PDF…': 'Search in PDF pages…',
  'نتائج البحث في PDF': 'PDF Search Results',
  'لم يتم العثور على نتائج في هذا الملف': 'No matches found in this PDF',
  '🖨 طباعة': '🖨 Print',
  '☑️ مهام': '☑️ Checklist',
  'مهمة جديدة': 'New task',
  '📅 التاريخ': '📅 Date',
  'تم لصق الصورة من الحافظة': 'Image pasted from clipboard',
  'تم استخراج الصفحات بنجاح': 'Pages extracted successfully',

  /* ميزات متقدمة إضافية */
  '📰 عمودين': '📰 Two Columns',
  '¹ حاشية سفلية': '¹ Footnote',
  '💬 تعليق': '💬 Comment',
  '💬 التعليقات': '💬 Comments',
  '✔ حل وإزالة': '✔ Resolve',
  'التفاف يمين': 'Wrap Right',
  'التفاف يسار': 'Wrap Left',
  '🎨 تنسيق رأس الجدول': '🎨 Style Header',
  '🎨 ثيم العرض الكامل…': '🎨 Deck Theme…',
  'كلاسيكي أبيض': 'Classic White',
  'ليلي فاخر': 'Midnight Dark',
  'أزرق ملكي': 'Royal Blue',
  'أخضر زمردي': 'Emerald Green',
  'رملي دافئ': 'Warm Sand',
  'تم تطبيق القالب على العرض بالكامل': 'Theme applied to entire deck',
  'محاذاة مع الشريحة وحركة الدخول': 'Slide Align & Entrance Animation',
  '↔ توسيط أفقي للشريحة': '↔ Center H on Slide',
  '↕ توسيط عمودي للشريحة': '↕ Center V on Slide',
  'حركة ظهور العنصر في العرض': 'Object Entrance Animation',
  'بدون حركة دخول': 'No Animation',
  'تلاشي ناعم (Fade)': 'Fade In',
  'انزلاق للأعلى (Fly In)': 'Fly In',
  'تكبير (Zoom In)': 'Zoom In',
  '💠 مخطط SmartArt…': '💠 SmartArt Diagram…',
  'هرم تنظيمي': 'Pyramid Diagram',
  'دورة مستمرة': 'Cycle Diagram',
  'خطوات متتالية': 'Process Steps',
  'تم إدراج المخطط الذكي': 'SmartArt diagram inserted',
  'تأثيرات وفلاتر الصورة': 'Image Filters & Effects',
  'أبيض وأسود': 'Grayscale',
  'تباين عالي': 'High Contrast',
  'إضاءة ناعمة': 'Soft Brightness',
  'أصلي': 'Original',
  'تمت استعادة النسخة بنجاح': 'Version restored successfully',

  /* لوحة الخصائص الجانبية */
  'خصائص العنصر': 'Object Properties',
  'انقر مرتين على أي نص للتعديل المباشر، أو حدد عنصراً لتعديل خصائصه.': 'Double-click any text to edit inline, or select an object to edit its properties.',
  'العنصر غير متاح.': 'Object is unavailable.',
  'محتوى النص': 'Text Content',
  'اكتب النص هنا…': 'Type text here…',
  'حفظ النص': 'Save Text',
  'تم تحديث النص': 'Text updated',
  'العرض (px)': 'Width (px)',
  'الارتفاع (px)': 'Height (px)',
  'التدوير (°)': 'Rotation (°)',
  'تدوير': 'Rotate',
  'تطبيق الأبعاد': 'Apply Dimensions',
  'أدوات الصورة': 'Image Tools',
  '🔄 تبديل الصورة': '🔄 Replace Image',
  '↔ قلب أفقي': '↔ Flip H',
  '↕ قلب عمودي': '↕ Flip V',
  '⭕ قص دائري': '⭕ Circle Crop',
  '⬜ إعادة تعيين القص': '⬜ Reset Crop',
  'تنسيق النص': 'Text Formatting',
  'تنسيق نص الخلية المحددة': 'Selected Cell Text Format',
  'تطبيق على الخلية': 'Apply to Cell',
  'تم تطبيق التنسيق على نص الخلية': 'Cell text formatting applied',
  'تباعد الأسطر': 'Line Spacing',
  'تنسيق المحدد:': 'Selection:',
  'لون': 'Color',
  'حجم': 'Size',
  'لون النص المظلل': 'Selected text color',
  'حجم النص المظلل': 'Selected text size',
  'حجم الخط': 'Font Size',
  'نوع الخط': 'Font family',
  'محاذاة': 'Alignment',
  'يسار': 'Left',
  'وسط': 'Center',
  'يمين': 'Right',
  'ضبط': 'Justify',
  'اتجاه النص': 'Text Direction',
  'اتجاه RTL': 'RTL Direction',
  'اتجاه LTR': 'LTR Direction',
  'التعداد': 'List Style',
  'بدون تعداد': 'No numbering',
  'تعداد نقطي': 'Bullets',
  'تعداد رقمي': 'Numbered',
  'عريض': 'Bold',
  'مائل': 'Italic',
  'تحته خط': 'Underline',
  'لون الخط': 'Font Color',
  'تطبيق تنسيق النص': 'Apply Text Style',
  'المظهر والحدود': 'Appearance & Border',
  'لون التعبئة': 'Fill Color',
  'بدون تعبئة': 'No Fill',
  'لون الحدود': 'Border Color',
  'سمك الحدود (px)': 'Border Width (px)',
  'الشفافية (%)': 'Opacity (%)',
  'تأثير': 'Effect',
  'ظل خفيف': 'Drop Shadow',
  'تطبيق المظهر': 'Apply Appearance',
  'إدارة الجدول': 'Table Management',
  '➕ صف': '➕ Row',
  '➖ صف': '➖ Row',
  '➕ عمود': '➕ Column',
  '➖ عمود': '➖ Column',
  'لون الخلية المحددة': 'Selected Cell Fill',
  'تطبيق لون الخلية': 'Apply Cell Fill',
  'بيانات الرسم البياني': 'Chart Data',
  'نوع الرسم': 'Chart Type',
  'أعمدة': 'Bar',
  'خطي': 'Line',
  'دائري': 'Pie',
  'التسمية': 'Label',
  'القيمة': 'Value',
  '➕ إضافة صف/بيان': '➕ Add Row/Data',
  'تحديث الرسم': 'Update Chart',
  'رابط تشعبي (Hyperlink)': 'Hyperlink',
  'حفظ الرابط': 'Save Link',
  'إزالة الرابط': 'Remove link',
  'محاذاة العناصر المحددة': 'Align Selected Objects',
  'توسيط أفقي': 'Align Center',
  'أعلى': 'Top',
  'أسفل': 'Bottom',
  'توزيع أفقي': 'Distribute H',
  'توزيع عمودي': 'Distribute V',
  '🔴 ليزر': '🔴 Laser',
  '✕ إنهاء العرض': '✕ Exit Show',
  'لون الخلفية': 'Background Color',
  'لون النص': 'Text Color',

  /* التنبيهات والرسائل الديناميكية */
  'ابحث عن نص:': 'Find text:',
  'لا يحتوي هذا الملف على شرائح قابلة للعرض.': 'This file contains no displayable slides.',
  'تم فتح الملف بنجاح': 'File opened successfully',
  'تم إنشاء العرض التقديمي بنجاح': 'Presentation created successfully',
  'تم حفظ ملاحظات الشريحة': 'Slide notes saved',
  'تم إخفاء الشريحة من العرض': 'Slide hidden from slideshow',
  'تم إظهار الشريحة': 'Slide shown',
  'تمت إضافة الترقيم والتذييل لجميع الشرائح': 'Slide numbers & footer added to all slides',
  'تم تغيير تخطيط الشريحة': 'Slide layout updated',
  'تم ضبط الانتقال': 'Slide transition updated',
  'تم تطبيق الأبعاد': 'Dimensions applied',
  'تم تطبيق تنسيق النص': 'Text formatting applied',
  'تم تطبيق المظهر': 'Appearance applied',
  'انقر على خلية داخل الجدول أولاً': 'Click a table cell first',
  'تم تلوين الخلية': 'Cell fill applied',
  'تم تحديث الرسم البياني': 'Chart updated',
  'تم ربط العنصر بالرابط': 'Hyperlink saved',
  'تمت إزالة الرابط': 'Hyperlink removed',
  'تم تبديل الصورة بنجاح': 'Image replaced successfully',
  'انقر الآن على العنصر الهدف لتطبيق التنسيق عليه': 'Click the target object to apply copied formatting',
  'تم نسخ التنسيق بنجاح': 'Formatting copied successfully',
  'تمت محاذاة العناصر': 'Objects aligned',
  'حدد 3 عناصر على الأقل للتوزيع المتساوي': 'Select at least 3 objects to distribute evenly',
  'تم توزيع المسافات بالتساوي': 'Objects distributed evenly',
  'تعذر تحديد تخطيط الشريحة': 'Could not determine slide layout',
  'تمت إضافة شريحة جديدة': 'New slide added',
  'تم نسخ الشريحة': 'Slide duplicated',
  'لا يمكن حذف آخر شريحة': 'Cannot delete the last slide',
  'تم حذف الشريحة': 'Slide deleted',
  'تم تغيير ترتيب الشرائح': 'Slide order updated',
  'تمت إضافة مربع نص': 'Text box added',
  'تمت إضافة الشكل': 'Shape added',
  'تمت إضافة الأيقونة': 'Icon added',
  'تمت إضافة خط': 'Line added',
  'تمت إضافة جدول (انقر مرتين على أي خلية لتعديلها)': 'Table added (double-click any cell to edit)',
  'تمت إضافة الرسم البياني (عدّل أرقامه من القائمة الجانبية)': 'Chart added (edit values in the right panel)',
  'تمت إضافة الصورة': 'Image added',
  'تم إدراج الوسائط بنجاح': 'Media inserted successfully',
  'تم نسخ العنصر': 'Object copied',
  'تم لصق العنصر': 'Object pasted',
  'تم تكرار العنصر': 'Object duplicated',
  'تم حذف العنصر': 'Object deleted',
  'تم تقديم العنصر للأمام': 'Brought to front',
  'تم إرسال العنصر للخلف': 'Sent to back',
  'تم تجميع العناصر': 'Objects grouped',
  'تم فك التجميع': 'Objects ungrouped',
  'تم تغيير لون الخلفية': 'Background color updated',
  'تم تطبيق التدرج اللوني للخلفية': 'Background gradient applied',
  'تم تعيين صورة الخلفية': 'Background image set',
  'الشريحة الحالية ليس لها خلفية مخصصة': 'Current slide has no custom background',
  'تم تطبيق الخلفية على جميع الشرائح': 'Background applied to all slides',
  'جميع الشرائح مخفية!': 'All slides are hidden!',
  'اسمح بالنوافذ المنبثقة لفتح وضع مقدم العرض': 'Allow pop-ups to open Presenter View',
  'جارٍ فتح الملف…': 'Opening file…',
  'جارٍ تصدير الشريحة كصورة PNG…': 'Exporting slide as PNG…',
  'جارٍ تصدير جميع الشرائح في ملف ZIP…': 'Exporting all slides as ZIP…',
  'تم تصدير الشريحة كـ PNG': 'Slide exported as PNG',
  'تم تصدير جميع الشرائح كملف ZIP': 'All slides exported as ZIP',
  'تعذر تصدير الصورة': 'Failed to export image',
  'تعذر تصدير ملف ZIP': 'Failed to export ZIP archive',
  'اسمح بالنوافذ المنبثقة لتصدير PDF': 'Allow pop-ups to export PDF',
  'تمت إضافة جدول': 'Table added',
  'جارٍ تجهيز الصورة…': 'Preparing image…',
  'تعذر تحميل الصورة': 'Failed to load image',
  'جارٍ تبديل الصورة…': 'Replacing image…',
  'جارٍ تجهيز صورة الخلفية…': 'Preparing background image…',
  'جارٍ تجهيز ملف الوسائط…': 'Preparing media file…',
  'تعذر قراءة ملف الوسائط': 'Failed to read media file',
  'لم يتم العثور على النص المطلوب': 'Search text not found',
  'جارٍ تجهيز الملف للتحميل…': 'Preparing file for download…',
  'تم تحميل الملف': 'File downloaded',
  'حدث خطأ أثناء تصدير الملف': 'An error occurred while exporting the file',
  'قائمة': 'List', 'ترقيم': 'Numbering', 'الأنماط': 'Styles', 'تصفير': 'Reset',
  'السابق': 'Previous', 'التالي': 'Next', 'تأكيد': 'Confirm', 'تحميل': 'Loading',
  'إضافة رابط': 'Add link', 'تصغير': 'Zoom out',
  'جارٍ التحميل…': 'Loading…', 'جارٍ الحفظ…': 'Saving…', 'تم الحفظ': 'Saved',
  'تم التحديث': 'Updated', 'الرابط': 'URL',
  'صنع بواسطة': 'Made by', 'موقعي': 'my website', 'مفتوح المصدر على GitHub': 'Open source on GitHub',
  'مستند Word (.docx) أو عرض PowerPoint (.pptx) أو PDF (.pdf)': 'Word (.docx), PowerPoint (.pptx), or PDF (.pdf)',
  'السمة': 'Theme', 'قالب Word': 'Word Template', 'بسيط، سريع، ويحافظ على الخصوصية': 'Basic, Private, Fast',
  '🛡️ تنظيف البيانات': '🛡️ Metadata',
  'تنظيف البيانات': 'Metadata',
  '💾 حفظ مباشر': '💾 Save',
  'حفظ مباشر': 'Save',
  '⬛ تعتيم حقيقي': '⬛ Redact',
  'تعتيم حقيقي': 'Redact',
  '💧 علامة مائية': '💧 Watermark',
  'علامة مائية': 'Watermark',
  '🔢 ترقيم الصفحات': '🔢 Page #',
  'أوامر سريعة': 'Commands',
  '⚖️ مقارنة مستندين': 'Compare Docs',
  'مقارنة مستندين': 'Compare Docs',
  '🖼️ استخراج الصور': 'Extract Images',
  'استخراج الصور': 'Extract Images',
  '📐 مقاس A4 / قص': '📐 A4 / Crop',
  '📑 دمج PDF': 'Merge PDFs',
  'دمج PDF': 'Merge PDFs',
  '🛡️ تنظيف البيانات المخفية': 'Clean Metadata',
  'تنظيف البيانات المخفية': 'Clean Metadata'
};

const reverseTranslations = new Map();
const reverseCounts = new Map();
for (const en of Object.values(translations)) reverseCounts.set(en, (reverseCounts.get(en) || 0) + 1);
for (const [ar, en] of Object.entries(translations)) {
  if (reverseCounts.get(en) === 1) reverseTranslations.set(en, ar);
}

let currentLanguage = 'en';
const originalText = new WeakMap(), originalAttrs = new WeakMap();

function tr(value) {
  const text = String(value == null ? '' : value);
  if (currentLanguage === 'ar') {
    return reverseTranslations.get(text) || text;
  }
  if (translations[text] != null) return translations[text];
  const m = /^شريحة\s+(\d+)(\s*\(مخفية\))?$/.exec(text);
  if (m) return `Slide ${m[1]}${m[2] ? ' (Hidden)' : ''}`;
  const n = /^هل أنت متأكد من حذف الشريحة (\d+)\؟/.exec(text);
  if (n) return `Are you sure you want to delete slide ${n[1]}?`;
  const repl = /^تم استبدال (\d+) موضع$/.exec(text);
  if (repl) return `Replaced ${repl[1]} occurrence(s)`;
  const expProg = /^جارٍ تصدير الشريحة (\d+) من (\d+)…$/.exec(text);
  if (expProg) return `Exporting slide ${expProg[1]} of ${expProg[2]}…`;
  const batchProg = /^جارٍ تحويل (\d+) صور إلى شرائح…$/.exec(text);
  if (batchProg) return `Converting ${batchProg[1]} images to slides…`;
  const batchDone = /^تم تحويل (\d+) صور إلى شرائح بنجاح$/.exec(text);
  if (batchDone) return `Converted ${batchDone[1]} image(s) to slides`;
  return text;
}

function applyLanguage(lang) {
  const nextLanguage = lang === 'ar' ? 'ar' : 'en';
  const changed = nextLanguage !== currentLanguage;
  const paperBefore = $('.paper');
  let selectionOffset = null;
  if (changed && paperBefore) {
    const sel = window.getSelection();
    if (sel && sel.rangeCount && paperBefore.contains(sel.anchorNode)) {
      const pre = sel.getRangeAt(0).cloneRange(); pre.selectNodeContents(paperBefore); pre.setEnd(sel.anchorNode, sel.anchorOffset);
      selectionOffset = pre.toString().length;
      try { syncPaperToDocxXml(paperBefore); } catch (_) {}
    }
  }
  currentLanguage = nextLanguage;
  document.documentElement.lang = currentLanguage;
  document.documentElement.dir = currentLanguage === 'en' ? 'ltr' : 'rtl';
  document.body.classList.toggle('lang-en', currentLanguage === 'en');
  document.title = currentLanguage === 'en' ? 'brivatedocs — Document Editor' : 'brivatedocs — محرر المستندات';
  try { localStorage.setItem('doc-studio-language', currentLanguage); } catch (_) {}
  for (const s of [$('#languageSelect'), $('#languageSelectEditor')]) if (s) s.value = currentLanguage;
  translateTree(document.body);
  if (changed && $('#editor') && !$('#editor').hidden && kind) {
    try {
      if (kind === 'pdf') renderPdfEditor();
      else renderEditor();
      if (selectionOffset != null) requestAnimationFrame(() => {
        const paper = $('.paper'); if (!paper) return;
        const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT); let node, remaining = selectionOffset;
        while ((node = walker.nextNode())) { if (remaining <= node.nodeValue.length) { const r = document.createRange(); r.setStart(node, remaining); r.collapse(true); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r); break; } remaining -= node.nodeValue.length; }
      });
    } catch (err) { console.warn('Language redraw skipped:', err); }
  }
}

function translateTree(root) {
  if (!root) return;
  if (root.nodeType === 3) {
    if (root.parentElement && root.parentElement.closest('.slide, .paper, .pdf-page-box, textarea, input, #fname, #draftMeta')) return;
    const rawTrim = root.nodeValue?.trim();
    if (rawTrim === 'undefined' || rawTrim === 'null') {
      root.remove();
      return;
    }
    if (!originalText.has(root)) originalText.set(root, root.nodeValue);
    const source = originalText.get(root);
    const leading = (source.match(/^\s*/) || [''])[0];
    const trailing = (source.match(/\s*$/) || [''])[0];
    const core = source.slice(leading.length, source.length - trailing.length || source.length);
    const nextValue = leading + tr(core) + trailing;
    if (root.nodeValue !== nextValue) root.nodeValue = nextValue;
    return;
  }

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    if (node.parentElement && node.parentElement.closest('.slide, .paper, .pdf-page-box, textarea, input, #fname, #draftMeta')) continue;
    const rawTrim = node.nodeValue?.trim();
    if (rawTrim === 'undefined' || rawTrim === 'null') {
      node.remove();
      continue;
    }
    if (!originalText.has(node)) originalText.set(node, node.nodeValue);
    const source = originalText.get(node);
    const leading = (source.match(/^\s*/) || [''])[0];
    const trailing = (source.match(/\s*$/) || [''])[0];
    const core = source.slice(leading.length, source.length - trailing.length || source.length);
    const nextValue = leading + tr(core) + trailing;
    if (node.nodeValue !== nextValue) node.nodeValue = nextValue;
  }

  if (root.nodeType === 1) {
    const els = [root, ...$$('*', root)];
    for (const el of els) {
      for (const attr of ['title', 'placeholder', 'aria-label']) {
        if (!el.hasAttribute(attr)) continue;
        let map = originalAttrs.get(el);
        if (!map) { map = {}; originalAttrs.set(el, map); }
        if (map[attr] === undefined) map[attr] = el.getAttribute(attr);
        el.setAttribute(attr, tr(map[attr]));
      }
    }
  }
}

function setupLanguage() {
  let saved = 'en';
  try { saved = localStorage.getItem('doc-studio-language') || 'en'; } catch (_) {}
  $('#languageSelect')?.addEventListener('change', e => applyLanguage(e.target.value));
  $('#languageSelectEditor')?.addEventListener('change', e => applyLanguage(e.target.value));
  const observer = new MutationObserver(records => records.forEach(r => {
    if (r.type === 'childList') r.addedNodes.forEach(n => {
      if (n.nodeType === 1 && n.classList?.contains('smart-guide')) return;
      if (n.nodeType === 1 || n.nodeType === 3) translateTree(n);
    });
  }));
  observer.observe(document.body, { childList: true, subtree: true });
  applyLanguage(saved);
}

class UserError extends Error {}

function h(tag, attrs, ...children) {
  const e = document.createElement(tag);
  for (const k in (attrs || {})) {
    const v = attrs[k];
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else if (k.startsWith('on')) e[k] = v;
    else if (k === 'value') e.value = v;
    else if (k === 'draggable') e.setAttribute('draggable', v === true || v === 'true' ? 'true' : 'false');
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) if (c != null && c !== false) e.append(c);
  return e;
}

function svgEl(tag, attrs) {
  const e = document.createElementNS(NS.svg, tag);
  for (const k in (attrs || {})) e.setAttribute(k, attrs[k]);
  return e;
}

function btn(label, fn, cls, opt) {
  opt = opt || {};
  return h('button', {
    class: 'btn ' + (cls || ''),
    onclick: fn,
    title: opt.title ? tr(opt.title) : null,
    disabled: opt.disabled,
    onpointerdown: opt.keep ? (e => e.preventDefault()) : null
  }, tr(label));
}

let toastTimer;
function toast(m) {
  const t = $('#toast');
  if (!t) return;
  t.textContent = tr(m);
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3200);
}

function busy(on, msg) {
  const b = $('#busy');
  if (b) b.hidden = !on;
  if (msg && $('#busyMsg')) $('#busyMsg').textContent = tr(msg);
}

/* ── الحفظ التلقائي وإدارة آخر 5 مشاريع عبر IndexedDB ── */
let activeDraftId = null;
const DraftDB = {
  dbName: 'DocStudioAutoSave',
  storeName: 'drafts',
  listKey: 'recent_drafts_v2',
  open() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) return reject(new Error('IndexedDB unsupported'));
      const req = indexedDB.open(this.dbName, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },
  async loadAll() {
    let db;
    try {
      db = await this.open();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(this.storeName, 'readonly');
        const req = tx.objectStore(this.storeName).get(this.listKey);
        let result = [];
        req.onsuccess = () => { result = Array.isArray(req.result) ? req.result : []; };
        req.onerror = () => reject(req.error || new Error('Could not read saved drafts'));
        tx.oncomplete = () => resolve(result);
        tx.onerror = () => reject(tx.error || new Error('Could not read saved drafts'));
        tx.onabort = () => reject(tx.error || new Error('Draft read was aborted'));
      });
    } catch (err) {
      console.warn('Could not load saved drafts:', err);
      return [];
    } finally {
      if (db) db.close();
    }
  },
  // Keep list reads and writes in one transaction to prevent stale autosave updates.
  async updateList(mutator) {
    const db = await this.open();
    try {
      await new Promise((resolve, reject) => {
        const tx = db.transaction(this.storeName, 'readwrite');
        const store = tx.objectStore(this.storeName);
        const req = store.get(this.listKey);
        req.onsuccess = () => {
          try {
            const current = Array.isArray(req.result) ? req.result : [];
            const next = mutator(current);
            if (next === null) store.delete(this.listKey);
            else store.put(next, this.listKey);
          } catch (err) {
            try { tx.abort(); } catch (_) {}
            reject(err);
          }
        };
        req.onerror = () => reject(req.error || new Error('Could not read saved drafts'));
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error || new Error('Could not update saved drafts'));
        tx.onabort = () => reject(tx.error || new Error('Saved draft update was aborted'));
      });
    } finally {
      db.close();
    }
  },
  async save(payload) {
    try {
      if (!activeDraftId) activeDraftId = 'draft_' + Date.now();
      const draftId = activeDraftId;
      const entry = { ...payload, id: draftId };
      await this.updateList((list) => {
        const filtered = list.filter(item => item.id !== draftId);
        filtered.unshift(entry);
        return filtered.slice(0, 5);
      });
    } catch (err) {
      console.warn('Draft autosave failed:', err);
    }
  },
  async removeById(id) {
    try {
      await this.updateList((list) => list.filter(item => item.id !== id));
    } catch (err) {
      console.warn('Could not remove saved draft:', err);
    }
  },
  async clear() {
    try {
      await this.updateList(() => null);
    } catch (err) {
      console.warn('Could not clear saved drafts:', err);
    }
  }
};

/* ── أدوات DOM XML ── */
function kids(el, ns, ln) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling)
    if ((!ns || c.namespaceURI === ns) && (!ln || c.localName === ln)) out.push(c);
  return out;
}
function kid(el, ns, ln) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling)
    if ((!ns || c.namespaceURI === ns) && (!ln || c.localName === ln)) return c;
  return null;
}
function desc(el, ns, ln) { return el ? Array.from(el.getElementsByTagNameNS(ns, ln)) : []; }
function mk(doc, ns, qname, attrs) {
  const e = doc.createElementNS(ns, qname);
  for (const k in (attrs || {})) {
    const v = attrs[k];
    if (v == null) continue;
    if (k.startsWith('w:')) e.setAttributeNS(W, k, v);
    else if (k.startsWith('r:')) e.setAttributeNS(R, k, v);
    else if (k === 'xml:space') e.setAttributeNS('http://www.w3.org/XML/1998/namespace', k, v);
    else e.setAttribute(k, v);
  }
  return e;
}
const XMLNS_ALL = `xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${R}" xmlns:w="${W}" xmlns:wp="${NS.wp}" xmlns:pic="${NS.pic}" xmlns:p14="${NS.p14}" xmlns:c="${C}"`;
function frag(doc, xml) {
  const d = new DOMParser().parseFromString(`<root ${XMLNS_ALL}>${xml}</root>`, 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw new Error('قالب XML داخلي غير صالح');
  return doc.importNode(d.documentElement.firstElementChild, true);
}
function removeNode(n) { if (n && n.parentNode) n.parentNode.removeChild(n); }
function insertAfter(ref, n) { ref.parentNode.insertBefore(n, ref.nextSibling); }

/* ── الحزمة (ملف ZIP في الذاكرة) ── */
const TEXT_EXT = /\.(xml|rels)$/i;
function serialize(d) {
  let s = new XMLSerializer().serializeToString(d);
  if (!/^<\?xml/.test(s)) s = XMLDECL + s;
  return s;
}
class Pkg {
  constructor() { this.files = new Map(); this.docs = new Map(); this.mainPart = ''; }
  static async load(buf) {
    const MAX_ZIP_ENTRIES = 5000;
    const MAX_ZIP_ENTRY_BYTES = 100 * 1024 * 1024;
    const MAX_ZIP_TOTAL_BYTES = 250 * 1024 * 1024;
    if (!buf || !Number.isFinite(buf.byteLength) || buf.byteLength > MAX_FILE_BYTES) {
      throw new UserError('الملف كبير جداً أو غير صالح (الحد 100MB).');
    }

    let zip;
    try { zip = await JSZip.loadAsync(buf); }
    catch (_) { throw new UserError('الملف تالف أو ليس ملف Office صحيحاً (تعذر فك الضغط).'); }

    const entries = Object.entries(zip.files).filter(([, file]) => !file.dir);
    if (entries.length > MAX_ZIP_ENTRIES) {
      throw new UserError('ملف Office يحتوي على عدد كبير جداً من الملفات الداخلية.');
    }

    // Check ZIP central-directory sizes before decompressing content to limit ZIP bombs.
    let declaredTotal = 0;
    for (const [name, file] of entries) {
      const declaredSize = Number(file?._data?.uncompressedSize);
      if (Number.isFinite(declaredSize) && declaredSize >= 0) {
        if (declaredSize > MAX_ZIP_ENTRY_BYTES) {
          throw new UserError('أحد الملفات الداخلية كبير جداً: ' + name);
        }
        declaredTotal += declaredSize;
        if (declaredTotal > MAX_ZIP_TOTAL_BYTES) {
          throw new UserError('حجم محتويات ملف Office بعد فك الضغط يتجاوز الحد المسموح (250MB).');
        }
      }
    }

    const pk = new Pkg();
    let actualTotal = 0;
    for (const [name, f] of entries) {
      if (TEXT_EXT.test(name)) {
        let s = await f.async('string');
        const size = new Blob([s]).size;
        actualTotal += size;
        if (size > MAX_ZIP_ENTRY_BYTES || actualTotal > MAX_ZIP_TOTAL_BYTES) {
          throw new UserError('حجم محتويات ملف Office بعد فك الضغط يتجاوز الحدود المسموحة.');
        }
        if (s.charCodeAt(0) === 0xFEFF) s = s.slice(1);
        pk.files.set(name, s);
      } else {
        const data = await f.async('uint8array');
        actualTotal += data.byteLength;
        if (data.byteLength > MAX_ZIP_ENTRY_BYTES || actualTotal > MAX_ZIP_TOTAL_BYTES) {
          throw new UserError('حجم محتويات ملف Office بعد فك الضغط يتجاوز الحدود المسموحة.');
        }
        pk.files.set(name, data);
      }
    }
    return pk;
  }
  has(p) { return this.files.has(p); }
  doc(p) {
    if (this.docs.has(p)) return this.docs.get(p);
    const s = this.files.get(p);
    if (typeof s !== 'string') return null;
    const d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new UserError('ملف داخلي تالف: ' + p);
    this.docs.set(p, d);
    return d;
  }
  setText(p, s) { this.files.set(p, s); this.docs.delete(p); }
  setBin(p, u8) { this.files.set(p, u8); }
  remove(p) { this.files.delete(p); this.docs.delete(p); }
  commit() { for (const [p, d] of this.docs) if (this.files.has(p)) this.files.set(p, serialize(d)); }
  snapshot() { this.commit(); return new Map(this.files); }
  restore(m) { this.files = new Map(m); this.docs.clear(); }
  async toBlob(mime) {
    this.commit();
    const zip = new JSZip();
    const names = [...this.files.keys()];
    names.sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
    for (const n of names) zip.file(n, this.files.get(n), { createFolders: false });
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: mime });
  }
}
