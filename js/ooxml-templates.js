'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 2 من 6: إنشاء الملفات الجديدة، العلاقات، الصور، والألوان
   ═════════════════════════════════════════════════════════════ */

/* ── إنشاء عرض تقديمي PPTX جديد من الصفر ── */
const DECK_THEMES = {
  light:      { bg: 'FFFFFF', ink: '1F1C17', acc: 'B4531A', sub: '6E655A' },
  dark:       { bg: '14171F', ink: 'F5F7FA', acc: '38BDF8', sub: '94A3B8' },
  navy:       { bg: '0F2547', ink: 'FFFFFF', acc: 'F59E0B', sub: 'CBD5E1' },
  warm:       { bg: 'FAF6EE', ink: '2C241B', acc: '0D9488', sub: '786A58' },
  emerald:    { bg: '0D2B22', ink: 'F0FDF4', acc: '34D399', sub: 'A7F3D0' },
  royal:      { bg: '1E1B3A', ink: 'F8FAFC', acc: 'A78BFA', sub: 'C4B5FD' },
  terracotta: { bg: 'FBF7F2', ink: '1C1917', acc: 'C2410C', sub: '78716C' },
  slate:      { bg: 'F8FAFC', ink: '0F172A', acc: '2563EB', sub: '64748B' }
};

function createNewBlankDeck(themeKey) {
  const th = DECK_THEMES[themeKey] || DECK_THEMES.light;
  const newPk = new Pkg();
  newPk.mainPart = 'ppt/presentation.xml';

  newPk.setText('[Content_Types].xml', XMLDECL + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="${CT_SLIDE}"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  </Types>`);

  newPk.setText('_rels/.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}officeDocument" Target="ppt/presentation.xml"/>
  </Relationships>`);

  newPk.setText('ppt/presentation.xml', XMLDECL + `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
    <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
    <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
    <p:sldSz cx="12192000" cy="6858000"/>
    <p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>`);

  newPk.setText('ppt/_rels/presentation.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}slideMaster" Target="slideMasters/slideMaster1.xml"/>
    <Relationship Id="rId2" Type="${REL}slide" Target="slides/slide1.xml"/>
  </Relationships>`);

  newPk.setText('ppt/theme/theme1.xml', XMLDECL + `<a:theme xmlns:a="${A}" name="DocStudioTheme">
    <a:themeElements><a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="${th.ink}"/></a:dk1><a:lt1><a:srgbClr val="${th.bg}"/></a:lt1>
      <a:dk2><a:srgbClr val="${th.sub}"/></a:dk2><a:lt2><a:srgbClr val="${th.bg}"/></a:lt2>
      <a:accent1><a:srgbClr val="${th.acc}"/></a:accent1><a:accent2><a:srgbClr val="2563EB"/></a:accent2>
      <a:accent3><a:srgbClr val="10B981"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4>
      <a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="EC4899"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme><a:fontScheme name="Standard"><a:majorFont><a:latin typeface="Cairo"/></a:majorFont><a:minorFont><a:latin typeface="Cairo"/></a:minorFont></a:fontScheme><a:fmtScheme name="Standard"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>
  </a:theme>`);

  newPk.setText('ppt/slideMasters/slideMaster1.xml', XMLDECL + `<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
    <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${th.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>
  </p:sldMaster>`);

  newPk.setText('ppt/slideMasters/_rels/slideMaster1.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="${REL}slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
    <Relationship Id="rId3" Type="${REL}theme" Target="../theme/theme1.xml"/>
  </Relationships>`);

  newPk.setText('ppt/slideLayouts/slideLayout1.xml', XMLDECL + `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="title"><p:cSld name="${currentLanguage === 'en' ? 'Title Slide' : 'شريحة عنوان'}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld></p:sldLayout>`);
  newPk.setText('ppt/slideLayouts/_rels/slideLayout1.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${REL}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);

  newPk.setText('ppt/slideLayouts/slideLayout2.xml', XMLDECL + `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="blank"><p:cSld name="${currentLanguage === 'en' ? 'Blank Slide' : 'شريحة فارغة'}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld></p:sldLayout>`);
  newPk.setText('ppt/slideLayouts/_rels/slideLayout2.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${REL}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);

  newPk.setText('ppt/slides/slide1.xml', XMLDECL + `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1219200" y="2100000"/><a:ext cx="9753600" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"${currentLanguage === 'en' ? '' : ' rtl="1"'}/><a:r><a:rPr lang="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}" sz="4400" b="1"><a:solidFill><a:srgbClr val="${th.ink}"/></a:solidFill></a:rPr><a:t>${currentLanguage === 'en' ? 'Presentation Title' : 'عنوان العرض التقديمي'}</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="4" name="AccentBar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="5181600" y="3320000"/><a:ext cx="1828800" cy="54000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${th.acc}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
      </p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="2133600" y="3580000"/><a:ext cx="7924800" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"${currentLanguage === 'en' ? '' : ' rtl="1"'}/><a:r><a:rPr lang="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}" sz="2200"><a:solidFill><a:srgbClr val="${th.acc}"/></a:solidFill></a:rPr><a:t>${currentLanguage === 'en' ? 'Double-click any text to edit' : 'انقر مرتين على أي نص للتعديل المباشر'}</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree></p:cSld>
  </p:sld>`);

  newPk.setText('ppt/slides/_rels/slide1.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  </Relationships>`);

  return newPk;
}

/* ── إنشاء مستند Word (DOCX) جديد ── */
function createNewBlankDocx(templateKey = 'blank') {
  const newPk = new Pkg();
  newPk.mainPart = 'word/document.xml';

  newPk.setText('[Content_Types].xml', XMLDECL + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  </Types>`);

  newPk.setText('_rels/.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/>
  </Relationships>`);

  newPk.setText('word/_rels/document.xml.rels', XMLDECL + `<Relationships xmlns="${NS.rel}">
    <Relationship Id="rId1" Type="${REL}numbering" Target="numbering.xml"/>
    <Relationship Id="rId2" Type="${REL}styles" Target="styles.xml"/>
  </Relationships>`);
  newPk.setText('word/styles.xml', XMLDECL + `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:lang w:val="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}" w:bidi="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`);

  newPk.setText('word/numbering.xml', XMLDECL + `<w:numbering xmlns:w="${W}">
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    </w:abstractNum>
    <w:abstractNum w:abstractNumId="1">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    </w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
  </w:numbering>`);

  const templateEnglish = {
    'الاسم الكامل': 'Full Name', 'المسمى الوظيفي | البريد الإلكتروني | رقم الهاتف | المدينة': 'Job Title | Email | Phone | City',
    'الملخص المهني': 'Professional Summary', 'اكتب نبذة موجزة عن خبراتك ومهاراتك وأبرز إنجازاتك المهنية هنا…': 'Write a brief summary of your experience, skills, and key achievements here…',
    'الخبرات العملية': 'Work Experience', 'المسمى الوظيفي — اسم الجهة (2024 - الآن)': 'Job Title — Organization (2024 - Present)',
    'المؤهلات العلمية': 'Education', 'درجة البكالوريوس — اسم الجامعة': 'Bachelor’s Degree — University Name',
    'المهارات الأساسية': 'Core Skills', 'إدارة المشاريع، تحليل البيانات، التواصل الفعّال': 'Project Management, Data Analysis, Effective Communication',
    'التاريخ:': 'Date:', 'سعادة مدير الإدارة المحترم': 'Dear Department Manager,', 'السلام عليكم ورحمة الله وبركاته،،،': 'Greetings,',
    'الموضوع: خطاب رسمي بخصوص ...': 'Subject: Official letter regarding …',
    'نفيد سعادتكم بأنه تم اعتماد الخطة التشغيلية للربع القادم، ونأمل منكم الاطلاع والتوجيه بما ترونه مناسباً.': 'Please be informed that the operating plan for the next quarter has been approved. Kindly review and provide your guidance.',
    'وتفضلوا بقبول فائق الاحترام والتقدير،،،': 'Sincerely,', 'مقدمه: الاسم والمنصب': 'Submitted by: Name and Title',
    'عنوان التقرير أو البحث العلمي': 'Report or Research Title', 'إعداد: اسم الطالب / الباحث — العام الجامعي 1448هـ': 'Prepared by: Student / Researcher — Academic Year 2026/2027',
    '1. المقدمة': '1. Introduction', 'تتناول هذه الدراسة استعراض أبرز المحاور والأهداف الرئيسية للموضوع...': 'This study reviews the main themes and objectives of the topic…',
    '2. المنهجية والنتائج': '2. Methodology and Results', 'تم تحليل البيانات والوصول إلى النتائج التالية:': 'The data was analyzed and the following results were obtained:',
    '3. الخاتمة والتوصيات': '3. Conclusion and Recommendations', 'بناءً على ما سبق نوصي بتطوير آليات العمل والمتابعة الدورية.': 'Based on the above, we recommend improving work processes and regular follow-up.',
    'فاتورة ضريبية مبسطة': 'Simplified Tax Invoice', 'رقم الفاتورة: #INV-1001  |  التاريخ:': 'Invoice No.: #INV-1001 | Date:',
    'العميل: اسم العميل أو الشركة': 'Customer: Customer or Company Name', 'البند / الخدمة': 'Item / Service', 'الكمية': 'Quantity', 'السعر': 'Price',
    'تصميم وتطوير موقع إلكتروني': 'Website Design and Development', 'الإجمالي شامل الضريبة': 'Total Including Tax',
    'المجموع الفرعي': 'Subtotal', 'ضريبة القيمة المضافة (15%)': 'VAT (15%)',
    'محضر اجتماع': 'Meeting Minutes', 'التاريخ والوقت:': 'Date & Time:', 'المكان / المنصة: قاعة الاجتماعات الرئيسية': 'Location / Platform: Main Conference Room',
    'الحضور: الاسم 1، الاسم 2، الاسم 3': 'Attendees: Name 1, Name 2, Name 3',
    '1. جدول الأعمال': '1. Agenda', 'مراجعة تقدم المشروع الحالي وتحديد أولويات المرحلة القادمة.': 'Review current project progress and prioritize upcoming milestones.',
    '2. القرارات المعتمدة': '2. Key Decisions', 'اعتماد الجدول الزمني المحدّث وتوزيع المهام على الفريق.': 'Approved the updated schedule and assigned team deliverables.',
    '3. المهام والمتابعة': '3. Action Items', 'المهمة': 'Task', 'المسؤول': 'Owner', 'تاريخ التسليم': 'Due Date',
    'إعداد التقرير النهائي للمرحلة الأولى': 'Prepare Phase 1 final report', 'قائد الفريق': 'Team Lead',
    'مقترح مشروع': 'Project Proposal', 'مقدّم إلى: اسم الجهة / العميل': 'Prepared for: Client / Organization Name',
    '1. الملخص التنفيذي': '1. Executive Summary', 'يهدف هذا المقترح إلى تقديم حل متكامل يرفع كفاءة العمل ويحقق الأهداف التشغيلية بأعلى جودة.': 'This proposal outlines a comprehensive solution designed to improve efficiency and deliver operational goals.',
    '2. نطاق العمل والمخرجات': '2. Scope & Deliverables', 'تحليل المتطلبات وتصميم الحل التقني الكامل.': 'Requirements analysis and complete solution design.',
    'تنفيذ الاختبارات والتسليم النهائي مع التدريب.': 'Quality assurance testing, final delivery, and handover.',
    '3. الجدول الزمني والتكلفة': '3. Timeline & Budget', 'المرحلة': 'Phase', 'المدة': 'Duration', 'التكلفة التقديرية': 'Est. Cost',
    'التخطيط والتصميم': 'Planning & Design', 'أسبوعان': '2 Weeks', 'التطوير والإطلاق': 'Development & Launch', '4 أسابيع': '4 Weeks',
    'اتفاقية تقديم خدمات': 'Service Agreement', 'الطرف الأول (العميل): الاسم / الشركة': 'First Party (Client): Name / Company',
    'الطرف الثاني (مقدّم الخدمة): الاسم / الشركة': 'Second Party (Provider): Name / Company',
    'البند الأول: نطاق الخدمات': 'Article 1: Scope of Services', 'يلتزم الطرف الثاني بتنفيذ الخدمات المتفق عليها وفق أعلى معايير الجودة وفي المواعيد المحددة.': 'The Service Provider agrees to deliver the agreed services with high quality and within the scheduled timeline.',
    'البند الثاني: المقابل المالي والدفعات': 'Article 2: Fees & Payment Terms', 'تُسدد الدفعة الأولى (50%) عند التوقيع، والدفعة النهائية (50%) عند التسليم.': '50% upfront upon signing, and 50% final payment upon completion.',
    'البند الثالث: السرية وحماية البيانات': 'Article 3: Confidentiality', 'يلتزم الطرفان بالحفاظ على سرية كافة المعلومات والوثائق المتبادلة.': 'Both parties agree to keep all shared information and documents strictly confidential.',
    'توقيع الطرف الأول: ________________   |   توقيع الطرف الثاني: ________________': 'First Party Signature: ________________   |   Second Party Signature: ________________',
    'ملخص دراسي — عنوان المادة': 'Study Notes — Course Title', 'الفصل / الوحدة: الأولى  |  التاريخ:': 'Chapter / Unit: 01  |  Date:',
    'أهم المفاهيم والمصطلحات': 'Key Concepts & Definitions', 'المفهوم الأول: شرح مختصر وواضح للفكرة الأساسية.': 'Concept 1: Clear and concise explanation of the core idea.',
    'المفهوم الثاني: مثال تطبيقي يسهل تذكره في الاختبار.': 'Concept 2: Practical example for quick exam recall.',
    'نقاط المراجعة السريعة': 'Quick Review Questions', 'س1: ما هو الفرق الرئيسي بين المفهوم الأول والثاني؟': 'Q1: What is the main difference between Concept 1 and Concept 2?'
  };
  const pXml = (txt, opts = {}) => {
    const displayText = currentLanguage === 'en' ? (templateEnglish[txt] || txt) : txt;
    const styleTag = opts.h ? `<w:pStyle w:val="Heading${opts.h}"/>` : '';
    const numTag = opts.list ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${opts.list}"/></w:numPr>` : '';
    const jc = opts.jc || (currentLanguage === 'en' ? 'left' : 'right');
    const sz = (opts.sz || 12) * 2;
    const bTag = opts.b ? '<w:b/>' : '';
    const colTag = opts.col ? `<w:color w:val="${opts.col}"/>` : '';
    const safeTxt = String(displayText).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const runTag = safeTxt
      ? `<w:r><w:rPr><w:rFonts w:ascii="Cairo" w:hAnsi="Cairo" w:cs="Cairo"/><w:lang w:val="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}" w:bidi="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}"/>${bTag}${colTag}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t>${safeTxt}</w:t></w:r>`
      : '';
    return `<w:p><w:pPr>${styleTag}${numTag}${currentLanguage === 'en' ? '<w:bidi w:val="0"/>' : '<w:bidi/>'}<w:jc w:val="${jc}"/></w:pPr>${runTag}</w:p>`;
  };

  let bodyInner = '';
  if (templateKey === 'cv') {
    bodyInner = [
      pXml('الاسم الكامل', { h: 1, sz: 22, b: true, col: 'B4531A', jc: 'center' }),
      pXml('المسمى الوظيفي | البريد الإلكتروني | رقم الهاتف | المدينة', { sz: 11, col: '666666', jc: 'center' }),
      pXml('الملخص المهني', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('اكتب نبذة موجزة عن خبراتك ومهاراتك وأبرز إنجازاتك المهنية هنا…'),
      pXml('الخبرات العملية', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('المسمى الوظيفي — اسم الجهة (2024 - الآن)', { list: 1, b: true }),
      pXml('المؤهلات العلمية', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('درجة البكالوريوس — اسم الجامعة', { list: 1 }),
      pXml('المهارات الأساسية', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('إدارة المشاريع، تحليل البيانات، التواصل الفعّال', { list: 1 })
    ].join('');
  } else if (templateKey === 'letter') {
    bodyInner = [
      pXml(`${currentLanguage === 'en' ? 'Date:' : 'التاريخ:'} ${new Date().toLocaleDateString(currentLanguage === 'en' ? 'en-US' : 'ar-SA')}`,  { jc: currentLanguage === 'en' ? 'right' : 'left', sz: 11, col: '666666' }),
      pXml('سعادة مدير الإدارة المحترم', { h: 2, sz: 15, b: true }),
      pXml('السلام عليكم ورحمة الله وبركاته،،،', { b: true }),
      pXml('الموضوع: خطاب رسمي بخصوص ...', { h: 3, sz: 13, b: true, col: 'B4531A' }),
      pXml('نفيد سعادتكم بأنه تم اعتماد الخطة التشغيلية للربع القادم، ونأمل منكم الاطلاع والتوجيه بما ترونه مناسباً.'),
      pXml('وتفضلوا بقبول فائق الاحترام والتقدير،،،', { jc: 'center' }),
      pXml('مقدمه: الاسم والمنصب', { jc: 'left', b: true })
    ].join('');
  } else if (templateKey === 'report') {
    bodyInner = [
      pXml('عنوان التقرير أو البحث العلمي', { h: 1, sz: 22, b: true, jc: 'center', col: '0F2547' }),
      pXml('إعداد: اسم الطالب / الباحث — العام الجامعي 1448هـ', { jc: 'center', sz: 11, col: '666666' }),
      pXml('1. المقدمة', { h: 2, sz: 16, b: true }),
      pXml('تتناول هذه الدراسة استعراض أبرز المحاور والأهداف الرئيسية للموضوع...'),
      pXml('2. المنهجية والنتائج', { h: 2, sz: 16, b: true }),
      pXml('تم تحليل البيانات والوصول إلى النتائج التالية:', { list: 2 }),
      pXml('3. الخاتمة والتوصيات', { h: 2, sz: 16, b: true }),
      pXml('بناءً على ما سبق نوصي بتطوير آليات العمل والمتابعة الدورية.')
    ].join('');
  } else if (templateKey === 'invoice') {
    bodyInner = [
      pXml('فاتورة ضريبية مبسطة', { h: 1, sz: 20, b: true, jc: 'center', col: 'B4531A' }),
      pXml(`${currentLanguage === 'en' ? 'Invoice No.: #INV-1001 | Date:' : 'رقم الفاتورة: #INV-1001  |  التاريخ:'} ${new Date().toLocaleDateString(currentLanguage === 'en' ? 'en-US' : 'ar-SA')}`,  { jc: 'center', sz: 11 }),
      pXml('العميل: اسم العميل أو الشركة', { b: true }),
      `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="666666"/><w:left w:val="single" w:sz="4" w:color="666666"/><w:bottom w:val="single" w:sz="4" w:color="666666"/><w:right w:val="single" w:sz="4" w:color="666666"/><w:insideH w:val="single" w:sz="4" w:color="666666"/><w:insideV w:val="single" w:sz="4" w:color="666666"/></w:tblBorders></w:tblPr>
        <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('البند / الخدمة', { b: true })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('الكمية', { b: true, jc: 'center' })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('السعر', { b: true, jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('تصميم وتطوير موقع إلكتروني')}</w:tc><w:tc>${pXml('1', { jc: 'center' })}</w:tc><w:tc>${pXml('2,500', { jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('المجموع الفرعي')}</w:tc><w:tc>${pXml('-', { jc: 'center' })}</w:tc><w:tc>${pXml('2,500', { jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('ضريبة القيمة المضافة (15%)')}</w:tc><w:tc>${pXml('-', { jc: 'center' })}</w:tc><w:tc>${pXml('375', { jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FEF3C7"/></w:tcPr>${pXml('الإجمالي شامل الضريبة', { b: true })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FEF3C7"/></w:tcPr>${pXml('-', { jc: 'center' })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FEF3C7"/></w:tcPr>${pXml('2,875', { b: true, jc: 'center', col: 'B4531A' })}</w:tc></w:tr>
      </w:tbl>`
    ].join('');
  } else if (templateKey === 'meeting') {
    const today = new Date().toLocaleDateString(currentLanguage === 'en' ? 'en-US' : 'ar-SA');
    bodyInner = [
      pXml('محضر اجتماع', { h: 1, sz: 22, b: true, jc: 'center', col: '0F2547' }),
      pXml(`${currentLanguage === 'en' ? 'Date & Time:' : 'التاريخ والوقت:'} ${today}`, { jc: 'center', sz: 11, col: '666666' }),
      pXml('المكان / المنصة: قاعة الاجتماعات الرئيسية', { sz: 11 }),
      pXml('الحضور: الاسم 1، الاسم 2، الاسم 3', { sz: 11, b: true }),
      pXml('1. جدول الأعمال', { h: 2, sz: 15, b: true, col: 'B4531A' }),
      pXml('مراجعة تقدم المشروع الحالي وتحديد أولويات المرحلة القادمة.', { list: 1 }),
      pXml('2. القرارات المعتمدة', { h: 2, sz: 15, b: true, col: 'B4531A' }),
      pXml('اعتماد الجدول الزمني المحدّث وتوزيع المهام على الفريق.', { list: 2 }),
      pXml('3. المهام والمتابعة', { h: 2, sz: 15, b: true, col: 'B4531A' }),
      `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="666666"/><w:left w:val="single" w:sz="4" w:color="666666"/><w:bottom w:val="single" w:sz="4" w:color="666666"/><w:right w:val="single" w:sz="4" w:color="666666"/><w:insideH w:val="single" w:sz="4" w:color="666666"/><w:insideV w:val="single" w:sz="4" w:color="666666"/></w:tblBorders></w:tblPr>
        <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('المهمة', { b: true })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('المسؤول', { b: true, jc: 'center' })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('تاريخ التسليم', { b: true, jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('إعداد التقرير النهائي للمرحلة الأولى')}</w:tc><w:tc>${pXml('قائد الفريق', { jc: 'center' })}</w:tc><w:tc>${pXml(today, { jc: 'center' })}</w:tc></w:tr>
      </w:tbl>`
    ].join('');
  } else if (templateKey === 'proposal') {
    bodyInner = [
      pXml('مقترح مشروع', { h: 1, sz: 24, b: true, jc: 'center', col: 'B4531A' }),
      pXml('مقدّم إلى: اسم الجهة / العميل', { jc: 'center', sz: 12, col: '666666' }),
      pXml('1. الملخص التنفيذي', { h: 2, sz: 16, b: true, col: '0F2547' }),
      pXml('يهدف هذا المقترح إلى تقديم حل متكامل يرفع كفاءة العمل ويحقق الأهداف التشغيلية بأعلى جودة.'),
      pXml('2. نطاق العمل والمخرجات', { h: 2, sz: 16, b: true, col: '0F2547' }),
      pXml('تحليل المتطلبات وتصميم الحل التقني الكامل.', { list: 1 }),
      pXml('تنفيذ الاختبارات والتسليم النهائي مع التدريب.', { list: 1 }),
      pXml('3. الجدول الزمني والتكلفة', { h: 2, sz: 16, b: true, col: '0F2547' }),
      `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="666666"/><w:left w:val="single" w:sz="4" w:color="666666"/><w:bottom w:val="single" w:sz="4" w:color="666666"/><w:right w:val="single" w:sz="4" w:color="666666"/><w:insideH w:val="single" w:sz="4" w:color="666666"/><w:insideV w:val="single" w:sz="4" w:color="666666"/></w:tblBorders></w:tblPr>
        <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('المرحلة', { b: true })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('المدة', { b: true, jc: 'center' })}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="F1F5F9"/></w:tcPr>${pXml('التكلفة التقديرية', { b: true, jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('التخطيط والتصميم')}</w:tc><w:tc>${pXml('أسبوعان', { jc: 'center' })}</w:tc><w:tc>${pXml('3,000', { jc: 'center' })}</w:tc></w:tr>
        <w:tr><w:tc>${pXml('التطوير والإطلاق')}</w:tc><w:tc>${pXml('4 أسابيع', { jc: 'center' })}</w:tc><w:tc>${pXml('7,000', { jc: 'center' })}</w:tc></w:tr>
      </w:tbl>`
    ].join('');
  } else if (templateKey === 'contract') {
    bodyInner = [
      pXml('اتفاقية تقديم خدمات', { h: 1, sz: 21, b: true, jc: 'center', col: '1F1C17' }),
      pXml('الطرف الأول (العميل): الاسم / الشركة', { b: true }),
      pXml('الطرف الثاني (مقدّم الخدمة): الاسم / الشركة', { b: true }),
      pXml('البند الأول: نطاق الخدمات', { h: 2, sz: 14, b: true, col: 'B4531A' }),
      pXml('يلتزم الطرف الثاني بتنفيذ الخدمات المتفق عليها وفق أعلى معايير الجودة وفي المواعيد المحددة.'),
      pXml('البند الثاني: المقابل المالي والدفعات', { h: 2, sz: 14, b: true, col: 'B4531A' }),
      pXml('تُسدد الدفعة الأولى (50%) عند التوقيع، والدفعة النهائية (50%) عند التسليم.'),
      pXml('البند الثالث: السرية وحماية البيانات', { h: 2, sz: 14, b: true, col: 'B4531A' }),
      pXml('يلتزم الطرفان بالحفاظ على سرية كافة المعلومات والوثائق المتبادلة.'),
      pXml(''),
      pXml('توقيع الطرف الأول: ________________   |   توقيع الطرف الثاني: ________________', { jc: 'center', b: true })
    ].join('');
  } else if (templateKey === 'study') {
    const today = new Date().toLocaleDateString(currentLanguage === 'en' ? 'en-US' : 'ar-SA');
    bodyInner = [
      pXml('ملخص دراسي — عنوان المادة', { h: 1, sz: 22, b: true, jc: 'center', col: '0D9488' }),
      pXml(`${currentLanguage === 'en' ? 'Chapter / Unit: 01  |  Date:' : 'الفصل / الوحدة: الأولى  |  التاريخ:'} ${today}`, { jc: 'center', sz: 11, col: '666666' }),
      pXml('أهم المفاهيم والمصطلحات', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('المفهوم الأول: شرح مختصر وواضح للفكرة الأساسية.', { list: 1 }),
      pXml('المفهوم الثاني: مثال تطبيقي يسهل تذكره في الاختبار.', { list: 1 }),
      pXml('نقاط المراجعة السريعة', { h: 2, sz: 16, b: true, col: '1F1C17' }),
      pXml('س1: ما هو الفرق الرئيسي بين المفهوم الأول والثاني؟', { list: 2 })
    ].join('');
  } else {
    bodyInner = pXml('');
  }

  newPk.setText('word/document.xml', XMLDECL + `<w:document xmlns:w="${W}" xmlns:r="${R}">
    <w:body>
      ${bodyInner}
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr>
    </w:body>
  </w:document>`);

  return newPk;
}

/* ── المسارات والعلاقات (rels) ── */
function relsPathOf(part) { const i = part.lastIndexOf('/'); return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels'; }
function resolveTarget(part, target) {
  try { target = decodeURIComponent(target); } catch (_) {}
  if (target[0] === '/') return target.slice(1);
  const base = part.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop(); else if (seg !== '.' && seg !== '') base.push(seg);
  }
  return base.join('/');
}
function relPath(from, to) {
  const f = from.split('/').slice(0, -1), t = to.split('/');
  let i = 0; while (i < f.length && i < t.length - 1 && f[i] === t[i]) i++;
  return '../'.repeat(f.length - i) + t.slice(i).join('/');
}
function readRels(pk, part) {
  const m = new Map();
  const rp = relsPathOf(part);
  if (!pk.has(rp)) return m;
  const d = pk.doc(rp);
  for (const r of kids(d.documentElement, null, 'Relationship')) {
    const mode = r.getAttribute('TargetMode');
    const t = r.getAttribute('Target') || '';
    m.set(r.getAttribute('Id'), {
      id: r.getAttribute('Id'), type: r.getAttribute('Type') || '', target: t, mode,
      path: mode === 'External' ? null : resolveTarget(part, t)
    });
  }
  return m;
}
function relByType(rels, suffix) { for (const r of rels.values()) if (r.type.endsWith('/' + suffix)) return r; return null; }
function addRel(pk, part, type, toPath, isExternal) {
  const rp = relsPathOf(part);
  if (!pk.has(rp)) pk.setText(rp, XMLDECL + `<Relationships xmlns="${NS.rel}"/>`);
  const d = pk.doc(rp), root = d.documentElement;
  if (isExternal) {
    const existing = kids(root).find(r => r.getAttribute('Type') === type && r.getAttribute('TargetMode') === 'External' && r.getAttribute('Target') === toPath);
    if (existing) return existing.getAttribute('Id');
  }
  let n = 0;
  for (const r of kids(root)) { const m = /^rId(\d+)$/.exec(r.getAttribute('Id') || ''); if (m) n = Math.max(n, +m[1]); }
  const id = 'rId' + (n + 1);
  const e = d.createElementNS(NS.rel, 'Relationship');
  e.setAttribute('Id', id);
  e.setAttribute('Type', type);
  e.setAttribute('Target', isExternal ? toPath : relPath(part, toPath));
  if (isExternal) e.setAttribute('TargetMode', 'External');
  root.appendChild(e);
  return id;
}
function removeRelIfUnused(pk, part, id, partDoc) {
  if (!id) return;
  const doc = partDoc || (pk.has(part) ? pk.doc(part) : null);
  if (doc) {
    const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_ELEMENT);
    let el;
    while ((el = walker.nextNode())) for (const attr of ['id', 'embed', 'link']) if (el.getAttributeNS(R, attr) === id) return;
  }
  removeRel(pk, part, id);
}
function removeRel(pk, part, id) {
  const rp = relsPathOf(part);
  if (!pk.has(rp)) return;
  const d = pk.doc(rp);
  for (const r of kids(d.documentElement)) if (r.getAttribute('Id') === id) removeNode(r);
}

/* ── أنواع المحتوى ── */
function ctRoot(pk) { return pk.doc('[Content_Types].xml').documentElement; }
function ensureDefault(pk, ext, type) {
  const r = ctRoot(pk); ext = ext.toLowerCase();
  if (kids(r, null, 'Default').some(e => (e.getAttribute('Extension') || '').toLowerCase() === ext)) return;
  const e = r.ownerDocument.createElementNS(r.namespaceURI, 'Default');
  e.setAttribute('Extension', ext); e.setAttribute('ContentType', type);
  r.insertBefore(e, r.firstChild);
}
function removeOverride(pk, name) { for (const e of kids(ctRoot(pk), null, 'Override')) if (e.getAttribute('PartName') === name) removeNode(e); }
function setOverride(pk, name, type) {
  removeOverride(pk, name);
  const r = ctRoot(pk), e = r.ownerDocument.createElementNS(r.namespaceURI, 'Override');
  e.setAttribute('PartName', name); e.setAttribute('ContentType', type); r.appendChild(e);
}
function uniquePath(pk, dir, base, ext) {
  for (let i = 1; ; i++) { const p = `${dir}/${base}${i}.${ext}`; if (!pk.has(p)) return p; }
}

/* ── الصور والوسائط ── */
const MIME = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml',
  bmp: 'image/bmp', webp: 'image/webp', mp4: 'video/mp4', webm: 'video/webm', mp3: 'audio/mpeg', wav: 'audio/wav'
};
const urlCache = new WeakMap();
function mediaURL(u8, path) {
  if (!(u8 instanceof Uint8Array)) return null;
  const ext = (path.split('.').pop() || '').toLowerCase();
  if (!MIME[ext]) return null;
  let u = urlCache.get(u8);
  if (!u) { u = URL.createObjectURL(new Blob([u8], { type: MIME[ext] })); urlCache.set(u8, u); }
  return u;
}
async function prepImage(file) {
  if (!file) throw new UserError('لم يتم استلام أي صورة');
  const bytes = new Uint8Array(await file.arrayBuffer());
  const url = URL.createObjectURL(new Blob([bytes], { type: file.type || 'application/octet-stream' }));
  try {
    const img = new Image();
    img.src = url;
    try { await img.decode(); } catch (_) { throw new UserError('الملف ليس صورة صالحة'); }
    const w = img.naturalWidth || 300, hh = img.naturalHeight || 150;
    const isPng = bytes[0] === 0x89 && bytes[1] === 0x50, isJpg = bytes[0] === 0xFF && bytes[1] === 0xD8, isGif = bytes[0] === 0x47 && bytes[1] === 0x49;
    if (isPng || isJpg || isGif) return { bytes, ext: isPng ? 'png' : isJpg ? 'jpeg' : 'gif', w, h: hh };
    const c = document.createElement('canvas'); c.width = w; c.height = hh;
    c.getContext('2d').drawImage(img, 0, 0, w, hh);
    const blob = await new Promise(res => c.toBlob(res, 'image/png'));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: 'png', w, h: hh };
  } finally { URL.revokeObjectURL(url); }
}
const CT_IMG = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif' };

/* ── الألوان والتدرجات اللونية ── */
const PRST_CLR = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', lime: '00FF00', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080', silver: 'C0C0C0', orange: 'FFA500', purple: '800080', brown: 'A52A2A', navy: '000080', maroon: '800000', olive: '808000', teal: '008080' };
function hex2rgb(x) { return [parseInt(x.slice(0, 2), 16), parseInt(x.slice(2, 4), 16), parseInt(x.slice(4, 6), 16)]; }
function rgb2hex(r, g, b) { return [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase(); }
function rgb2hsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let hh = 0, s = 0; const l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn);
    hh = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; hh /= 6;
  }
  return [hh, s, l];
}
function hsl2rgb(hh, s, l) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  return [f(p, q, hh + 1 / 3) * 255, f(p, q, hh) * 255, f(p, q, hh - 1 / 3) * 255];
}
function cssColor(c) {
  if (!c) return '';
  if (c.alpha != null && c.alpha < 1) { const [r, g, b] = hex2rgb(c.hex); return `rgba(${r},${g},${b},${c.alpha})`; }
  return '#' + c.hex;
}
function schemeHex(name, ctx) {
  if (name === 'phClr') return ctx.phClr || null;
  const m = (ctx.clrMap && ctx.clrMap[name]) || name;
  return ctx.theme[m] || ctx.theme[name] || null;
}
function parseColorEl(el, ctx) {
  let hex = null;
  switch (el.localName) {
    case 'srgbClr': hex = el.getAttribute('val'); break;
    case 'sysClr': hex = el.getAttribute('lastClr') || ({ windowText: '000000', window: 'FFFFFF' })[el.getAttribute('val')]; break;
    case 'schemeClr': hex = schemeHex(el.getAttribute('val'), ctx); break;
    case 'prstClr': hex = PRST_CLR[el.getAttribute('val')] || '000000'; break;
    default: return null;
  }
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  hex = hex.toUpperCase();
  let alpha = 1;
  for (const m of kids(el)) {
    const v = +m.getAttribute('val') / 100000;
    const [r, g, b] = hex2rgb(hex);
    switch (m.localName) {
      case 'lumMod': { const [hh, s, l] = rgb2hsl(r, g, b); hex = rgb2hex(...hsl2rgb(hh, s, Math.min(1, l * v))); break; }
      case 'lumOff': { const [hh, s, l] = rgb2hsl(r, g, b); hex = rgb2hex(...hsl2rgb(hh, s, Math.max(0, Math.min(1, l + v)))); break; }
      case 'satMod': { const [hh, s, l] = rgb2hsl(r, g, b); hex = rgb2hex(...hsl2rgb(hh, Math.min(1, s * v), l)); break; }
      case 'tint': hex = rgb2hex(r * v + 255 * (1 - v), g * v + 255 * (1 - v), b * v + 255 * (1 - v)); break;
      case 'shade': hex = rgb2hex(r * v, g * v, b * v); break;
      case 'alpha': alpha = v; break;
      default: break;
    }
  }
  return { hex, alpha };
}
function colorChild(el) { return kids(el).find(c => /Clr$/.test(c.localName)) || null; }
function fillFrom(parent, ctx) {
  if (!parent) return null;
  for (const c of kids(parent, A)) {
    switch (c.localName) {
      case 'noFill': return { type: 'none' };
      case 'solidFill': { const ce = kids(c)[0]; const col = ce && parseColorEl(ce, ctx); return col ? { type: 'solid', hex: col.hex, alpha: col.alpha } : null; }
      case 'gradFill': {
        const stops = [];
        const rawStops = [];
        for (const gs of kids(kid(c, A, 'gsLst'), A, 'gs')) {
          const ce = kids(gs)[0]; const col = ce && parseColorEl(ce, ctx);
          if (col) {
            const posPct = (+gs.getAttribute('pos') || 0) / 1000;
            stops.push(`#${col.hex} ${posPct}%`);
            rawStops.push({ hex: col.hex, alpha: col.alpha, pos: posPct });
          }
        }
        if (!stops.length) return null;
        const lin = kid(c, A, 'lin');
        const ooxmlDeg = lin ? Math.round((+(lin.getAttribute('ang')) || 0) / 60000) : 0;
        const cssDeg = ((ooxmlDeg + 90) % 360 + 360) % 360;
        const css = kid(c, A, 'path') ? `radial-gradient(circle, ${stops.join(',')})` : `linear-gradient(${cssDeg}deg, ${stops.join(',')})`;
        return { type: 'grad', css, rawStops, angle: cssDeg };
      }
      case 'blipFill': { const b = kid(c, A, 'blip'); return { type: 'blip', rid: b && b.getAttributeNS(R, 'embed') }; }
      case 'pattFill': { const fg = kid(c, A, 'fgClr'); const ce = fg && kids(fg)[0]; const col = ce && parseColorEl(ce, ctx); return col ? { type: 'solid', hex: col.hex, alpha: col.alpha } : null; }
      default: break;
    }
  }
  return null;
}
function applyFill(el, fill, L, pk) {
  if (!fill || fill.type === 'none') return;
  if (fill.type === 'solid') el.style.background = cssColor(fill);
  else if (fill.type === 'grad') el.style.background = fill.css;
  else if (fill.type === 'blip' && L) {
    const rel = L.rels.get(fill.rid), u8 = rel && rel.path && pk.files.get(rel.path);
    const u = u8 && mediaURL(u8, rel.path);
    if (u) { el.style.backgroundImage = `url("${u}")`; el.style.backgroundSize = '100% 100%'; }
  }
}
