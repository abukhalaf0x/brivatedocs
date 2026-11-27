'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 6 (ب): محرر DOCX المتكامل، التحميل، والاختصارات العامة
   ═════════════════════════════════════════════════════════════ */

/* ── محرر DOCX المتكامل ── */
function ensureDocxNumberingPart() {
  if (!pk.has('word/numbering.xml')) {
    pk.setText('word/numbering.xml', XMLDECL + `<w:numbering xmlns:w="${W}">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
      </w:abstractNum>
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
    </w:numbering>`);
  }
  setOverride(pk, '/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  const rels = readRels(pk, pk.mainPart);
  if (!relByType(rels, 'numbering')) {
    addRel(pk, pk.mainPart, REL + 'numbering', 'word/numbering.xml');
  }
}

function getDocxHeaderFooterText() {
  const rels = readRels(pk, pk.mainPart);
  const hRel = relByType(rels, 'header');
  const fRel = relByType(rels, 'footer');
  let headerTxt = '', footerTxt = '', hasPageNum = false;
  if (hRel && hRel.path && pk.has(hRel.path)) {
    headerTxt = desc(pk.doc(hRel.path).documentElement, W, 't').map(t => t.textContent).join('');
  }
  if (fRel && fRel.path && pk.has(fRel.path)) {
    const fDoc = pk.doc(fRel.path);
    footerTxt = desc(fDoc.documentElement, W, 't').map(t => t.textContent).join('');
    hasPageNum = desc(fDoc.documentElement, W, 'instrText').some(i => /PAGE/i.test(i.textContent || ''));
  }
  return { headerTxt, footerTxt, hasPageNum };
}

function setDocxHeaderFooter(headerTxt, footerTxt, includePageNum) {
  const doc = pk.doc(pk.mainPart);
  const body = kid(doc.documentElement, W, 'body');
  if (!body) return;
  let sectPr = kid(body, W, 'sectPr');
  if (!sectPr) {
    sectPr = mk(doc, W, 'w:sectPr');
    body.appendChild(sectPr);
  }

  for (const oldRef of [...kids(sectPr, W, 'headerReference'), ...kids(sectPr, W, 'footerReference')]) {
    const rid = oldRef.getAttributeNS(R, 'id');
    if (rid) removeRel(pk, pk.mainPart, rid);
    removeNode(oldRef);
  }

  if (headerTxt && headerTxt.trim()) {
    const hPath = 'word/header1.xml';
    const cleanH = escXml(headerTxt.trim());
    pk.setText(hPath, XMLDECL + `<w:hdr xmlns:w="${W}"><w:p><w:pPr>${currentLanguage === 'ar' ? '<w:bidi/>' : '<w:bidi w:val="0"/>'}<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/><w:color w:val="666666"/></w:rPr><w:t>${cleanH}</w:t></w:r></w:p></w:hdr>`);
    setOverride(pk, '/' + hPath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    const hRid = addRel(pk, pk.mainPart, REL + 'header', hPath);
    const hRef = mk(doc, W, 'w:headerReference', { 'w:type': 'default' });
    hRef.setAttributeNS(R, 'r:id', hRid);
    sectPr.insertBefore(hRef, sectPr.firstChild);
  }

  if ((footerTxt && footerTxt.trim()) || includePageNum) {
    const fPath = 'word/footer1.xml';
    const cleanF = escXml((footerTxt || '').trim());
    const pageFieldXml = includePageNum
      ? `<w:r><w:rPr><w:sz w:val="20"/><w:color w:val="666666"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
      : '';
    const txtXml = cleanF
      ? `<w:r><w:rPr><w:sz w:val="20"/><w:color w:val="666666"/></w:rPr><w:t xml:space="preserve">${cleanF}${includePageNum ? '  |  ' : ''}</w:t></w:r>`
      : '';
    pk.setText(fPath, XMLDECL + `<w:ftr xmlns:w="${W}"><w:p><w:pPr>${currentLanguage === 'ar' ? '<w:bidi/>' : '<w:bidi w:val="0"/>'}<w:jc w:val="center"/></w:pPr>${txtXml}${pageFieldXml}</w:p></w:ftr>`);
    setOverride(pk, '/' + fPath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
    const fRid = addRel(pk, pk.mainPart, REL + 'footer', fPath);
    const fRef = mk(doc, W, 'w:footerReference', { 'w:type': 'default' });
    fRef.setAttributeNS(R, 'r:id', fRid);
    sectPr.insertBefore(fRef, sectPr.firstChild);
  }
}

function readDocxPageSetup() {
  const doc = pk.doc(pk.mainPart);
  const body = kid(doc.documentElement, W, 'body');
  const sectPr = body && kid(body, W, 'sectPr');
  const pgSz = sectPr && kid(sectPr, W, 'pgSz');
  const pgMar = sectPr && kid(sectPr, W, 'pgMar');
  const pgBorders = sectPr && kid(sectPr, W, 'pgBorders');
  const bgEl = kid(doc.documentElement, W, 'background');

  if (bgEl) {
    const col = bgEl.getAttributeNS(W, 'color') || bgEl.getAttribute('color');
    if (col && /^[0-9a-fA-F]{6}$/.test(col)) curDocx.pageColor = '#' + col.toUpperCase();
  }
  curDocx.hasBorder = !!(pgBorders && kid(pgBorders, W, 'top'));

  if (pgSz) {
    const w = +(pgSz.getAttributeNS(W, 'w') || pgSz.getAttribute('w') || 11906);
    const orient = pgSz.getAttributeNS(W, 'orient') || pgSz.getAttribute('orient');
    curDocx.orientation = orient === 'landscape' ? 'landscape' : 'portrait';
    const minDim = Math.min(w, +(pgSz.getAttributeNS(W, 'h') || pgSz.getAttribute('h') || 16838));
    curDocx.pageSize = Math.abs(minDim - 12240) < 200 ? 'letter' : 'a4';
  }
  if (pgMar) {
    const left = +(pgMar.getAttributeNS(W, 'left') || pgMar.getAttribute('left') || 1440);
    if (left <= 800) curDocx.margins = 'narrow';
    else if (left >= 2000) curDocx.margins = 'wide';
    else curDocx.margins = 'normal';
  }
}

function applyDocxPageSetupToXml() {
  const doc = pk.doc(pk.mainPart);
  const root = doc.documentElement;
  const body = kid(root, W, 'body');
  if (!body) return;

  const oldBg = kid(root, W, 'background');
  if (oldBg) removeNode(oldBg);
  if (curDocx.pageColor && curDocx.pageColor.toUpperCase() !== '#FFFFFF') {
    const bg = mk(doc, W, 'w:background', { 'w:color': curDocx.pageColor.slice(1).toUpperCase() });
    root.insertBefore(bg, body);
  }

  let sectPr = kid(body, W, 'sectPr');
  if (!sectPr) {
    sectPr = mk(doc, W, 'w:sectPr');
    body.appendChild(sectPr);
  }

  let baseW = curDocx.pageSize === 'letter' ? 12240 : 11906;
  let baseH = curDocx.pageSize === 'letter' ? 15840 : 16838;
  if (curDocx.orientation === 'landscape') [baseW, baseH] = [baseH, baseW];

  let pgSz = kid(sectPr, W, 'pgSz');
  if (!pgSz) { pgSz = mk(doc, W, 'w:pgSz'); sectPr.insertBefore(pgSz, sectPr.firstChild); }
  pgSz.setAttributeNS(W, 'w:w', String(baseW));
  pgSz.setAttributeNS(W, 'w:h', String(baseH));
  if (curDocx.orientation === 'landscape') pgSz.setAttributeNS(W, 'w:orient', 'landscape');
  else pgSz.removeAttributeNS(W, 'orient');

  const marMap = {
    narrow: { top: '720', bottom: '720', left: '720', right: '720' },
    normal: { top: '1440', bottom: '1440', left: '1440', right: '1440' },
    wide:   { top: '1440', bottom: '1440', left: '2160', right: '2160' }
  };
  const m = marMap[curDocx.margins] || marMap.normal;
  let pgMar = kid(sectPr, W, 'pgMar');
  if (!pgMar) { pgMar = mk(doc, W, 'w:pgMar'); sectPr.appendChild(pgMar); }
  pgMar.setAttributeNS(W, 'w:top', m.top);
  pgMar.setAttributeNS(W, 'w:bottom', m.bottom);
  pgMar.setAttributeNS(W, 'w:left', m.left);
  pgMar.setAttributeNS(W, 'w:right', m.right);

  const oldBorders = kid(sectPr, W, 'pgBorders');
  if (oldBorders) removeNode(oldBorders);
  if (curDocx.hasBorder) {
    sectPr.appendChild(frag(doc, `<w:pgBorders xmlns:w="${W}" w:offsetFrom="page">
      <w:top w:val="single" w:sz="12" w:space="24" w:color="1F1C17"/>
      <w:left w:val="single" w:sz="12" w:space="24" w:color="1F1C17"/>
      <w:bottom w:val="single" w:sz="12" w:space="24" w:color="1F1C17"/>
      <w:right w:val="single" w:sz="12" w:space="24" w:color="1F1C17"/>
    </w:pgBorders>`));
  }
}

function openDocxCommentsDrawer(paper) {
  $$('.doc-comments-drawer').forEach(d => d.remove());
  const marks = $$('.doc-comment-mark', paper);
  if (!marks.length) return;
  const drawer = h('div', { class: 'doc-comments-drawer' },
    h('div', { class: 'row', style: { justifyContent: 'space-between' } },
      h('strong', null, tr('💬 التعليقات')),
      h('button', { class: 'btn', style: { padding: '2px 6px' }, onclick: () => drawer.remove() }, '✕')
    )
  );
  marks.forEach((m, i) => {
    drawer.append(h('div', { class: 'doc-comment-card' },
      h('div', { style: { fontWeight: '700', color: 'var(--acc)' } }, `#${i + 1}: "${m.innerText.slice(0, 28)}…"`),
      h('div', { style: { margin: '4px 0' } }, m.getAttribute('data-comment') || ''),
      btn('✔ حل وإزالة', () => {
        const parent = m.parentNode;
        while (m.firstChild) parent.insertBefore(m.firstChild, m);
        m.remove();
        openDocxCommentsDrawer(paper);
        syncPaperToDocxXml(paper);
        pushHistory('resolveComment');
      }, 'danger')
    ));
  });
  document.body.append(drawer);
}

function updateDocxPagination(paper) {
  if (!paper) return 1;
  const manualBreaks = $$('.doc-page-break', paper).length;
  const pageH = paper.classList.contains('landscape') ? 794 : 1123;
  const estimated = Math.max(1, Math.ceil((paper.scrollHeight || pageH) / pageH));
  return Math.max(estimated, manualBreaks + 1);
}

function insertDocxPageBreak(paper) {
  if (!paper) return;
  const br = h('div', { class: 'doc-page-break', contenteditable: 'false' }, h('span', null, tr('فاصل صفحة')));
  const nextP = h('div', { class: 'dp' }, '');
  const blocks = getDocxTargetBlocks(paper);
  if (blocks.length && blocks[0].parentNode === paper) {
    insertAfter(blocks[0], br);
    insertAfter(br, nextP);
  } else {
    paper.append(br, nextP);
  }
  syncPaperToDocxXml(paper);
  pushHistory('docPageBreak');
}

function renderDocxEditor() {
  syncEditorTopBarForKind();
  const app = $('#app');
  app.innerHTML = '';
  readDocxPageSetup();

  const container = h('div', { class: 'doc' });
  const tb = h('div', { class: 'tb2' });
  const ctxBar = h('div', { class: 'doc-ctx-bar', hidden: true });
  const hfPanel = h('div', { class: 'doc-hf-panel', hidden: !curDocx.showHf });
  const stickyBars = h('div', { class: 'doc-sticky-bars' }, tb, hfPanel, ctxBar);
  const paperWrap = h('div', { class: 'paper-wrap' });
  const ruler = h('div', { class: 'doc-ruler', contenteditable: 'false' });
  for (let cm = 1; cm <= 19; cm++) ruler.append(h('span', null, String(cm)));
  const paper = h('div', { class: 'paper', contenteditable: 'true', spellcheck: 'true' });
  paper.addEventListener('click', (event) => {
    const chk = event.target.closest?.('.doc-check-box');
    if (chk && paper.contains(chk)) {
      event.preventDefault();
      event.stopPropagation();
      const isDone = (chk.textContent || '').includes('☑');
      chk.textContent = isDone ? '☐ ' : '☑ ';
      chk.closest('.dp')?.classList.toggle('doc-check-done', !isDone);
      syncPaperToDocxXml(paper);
      pushHistory('docChecklistToggle');
      return;
    }
    const link = event.target.closest?.('a[href]');
    if (!link || !(event.ctrlKey || event.metaKey)) return;
    const href = link.getAttribute('href') || '';
    if (href.startsWith('#')) return;
    const safe = safeUrl(href); if (safe === '#') return;
    event.preventDefault(); window.open(safe, '_blank', 'noopener,noreferrer');
  });
  const statusBar = h('div', { class: 'doc-status-bar' });

  const renderWatermarkOverlay = () => {
    $$('.paper-watermark', paper).forEach(w => w.remove());
    if (curDocx.watermark && curDocx.watermark.trim()) {
      paper.prepend(h('div', { class: 'paper-watermark', contenteditable: 'false' }, curDocx.watermark.trim()));
    }
  };

  const renderHeaderFooterOverlay = () => {
    $$('.paper-hf-preview', paper).forEach(x => x.remove());
    const hf = getDocxHeaderFooterText();
    if (hf.headerTxt && hf.headerTxt.trim()) {
      paper.prepend(h('div', { class: 'paper-hf-preview paper-hdr-preview', contenteditable: 'false' }, hf.headerTxt.trim()));
    }
    if ((hf.footerTxt && hf.footerTxt.trim()) || hf.hasPageNum) {
      const fLabel = [hf.footerTxt?.trim(), hf.hasPageNum ? (currentLanguage === 'en' ? 'Page 1' : 'صفحة 1') : ''].filter(Boolean).join('  |  ');
      paper.append(h('div', { class: 'paper-hf-preview paper-ftr-preview', contenteditable: 'false' }, fLabel));
    }
  };

  const updatePaperClasses = () => {
    paper.classList.toggle('landscape', curDocx.orientation === 'landscape');
    paper.classList.toggle('letter', curDocx.pageSize === 'letter');
    paper.classList.toggle('has-border', !!curDocx.hasBorder);
    paper.classList.remove('margins-narrow', 'margins-wide');
    if (curDocx.margins === 'narrow') paper.classList.add('margins-narrow');
    else if (curDocx.margins === 'wide') paper.classList.add('margins-wide');
    paper.style.backgroundColor = curDocx.pageColor || '#FFFFFF';
    const scale = (+curDocx.zoom || 100) / 100;
    paper.style.transform = '';
    paper.style.zoom = scale === 1 ? '' : String(scale);
    renderWatermarkOverlay();
    renderHeaderFooterOverlay();
    updateDocxPagination(paper);
  };
  updatePaperClasses();

  let refreshDocxContextBar = () => {};
  let updateDocxContextBarVisibility = () => {};

  const capturePaperSel = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount && paper.contains(sel.anchorNode) && paper.contains(sel.focusNode)) {
      savedDocxSelectionRange = sel.getRangeAt(0).cloneRange();
    }
  };
  paper.addEventListener('keyup', capturePaperSel);
  paper.addEventListener('mouseup', capturePaperSel);
  const onDocxSelectionChange = () => {
    if (!paper.isConnected) {
      document.removeEventListener('selectionchange', onDocxSelectionChange);
      return;
    }
    capturePaperSel();
    if (curDocx.activeImg && paper.contains(curDocx.activeImg)) {
      updateDocxContextBarVisibility();
      return;
    }
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const anchor = sel.anchorNode;
    if (!anchor || !paper.contains(anchor)) return;
    const el = anchor.nodeType === 1 ? anchor : anchor.parentElement;
    const curBlk = el?.closest?.('.dp');
    if (curBlk && paper.contains(curBlk) && typeof headingSel !== 'undefined') {
      headingSel.value = curBlk.dataset?.heading || 'normal';
    }
    const td = el?.closest?.('td');
    const nextTd = (td && paper.contains(td)) ? td : null;
    if (nextTd !== curDocx.activeTd) {
      $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
      curDocx.activeTd = nextTd;
      if (nextTd) nextTd.classList.add('doc-td-active');
      refreshDocxContextBar();
    } else {
      updateDocxContextBarVisibility();
    }
  };
  document.addEventListener('selectionchange', onDocxSelectionChange);

  const headingSel = h('select', { title: tr('الأنماط') },
    h('option', { value: 'normal' }, tr('نص عادي')),
    h('option', { value: 'h1' }, tr('عنوان 1')),
    h('option', { value: 'h2' }, tr('عنوان 2')),
    h('option', { value: 'h3' }, tr('عنوان 3'))
  );
  headingSel.onchange = () => {
    applyDocxHeadingStyle(paper, headingSel.value);
    updateDocxPagination(paper);
  };

  const fontSel = h('select', { title: tr('نوع الخط') },
    ...[
      'Cairo', 'Tajawal', 'Noto Sans Arabic', 'IBM Plex Sans Arabic', 'Amiri', 'Traditional Arabic', 'Sakkal Majalla',
      'Tahoma', 'Segoe UI', 'Inter', 'Calibri', 'Arial', 'Verdana', 'Trebuchet MS',
      'Times New Roman', 'Georgia', 'Cambria', 'Garamond', 'Courier New', 'Consolas'
    ].map(f => h('option', { value: f }, f))
  );
  fontSel.onchange = () => execDocCmd('fontName', fontSel.value);

  const sizeSel = h('select', { title: tr('حجم الخط') },
    ...[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 26, 32, 36, 48].map(pt => h('option', { value: String(pt), selected: pt === 12 }, pt + ' pt'))
  );
  sizeSel.onchange = () => { applyDocxSelectionFontSize(paper, +sizeSel.value); updateDocxPagination(paper); };

  const bBold = btn('B', () => execDocCmd('bold'), '', { keep: true, title: 'عريض' }); bBold.style.fontWeight = '700';
  const bItalic = btn('I', () => execDocCmd('italic'), '', { keep: true, title: 'مائل' }); bItalic.style.fontStyle = 'italic';
  const bUnder = btn('U', () => execDocCmd('underline'), '', { keep: true, title: 'تحته خط' }); bUnder.style.textDecoration = 'underline';
  const bStrike = btn('S', () => execDocCmd('strikeThrough'), '', { keep: true, title: 'يتوسطه خط' }); bStrike.style.textDecoration = 'line-through';
  const bSup = btn('X²', () => execDocCmd('superscript'), '', { keep: true, title: 'نص مرتفع' });
  const bSub = btn('X₂', () => execDocCmd('subscript'), '', { keep: true, title: 'نص منخفض' });

  const colIn = h('input', {
    type: 'color',
    value: '#1F1C17',
    title: tr('لون النص'),
    onpointerdown: () => capturePaperSel(),
    oninput: () => applyDocxSelectionStyle(paper, 'color', colIn.value, false),
    onchange: () => applyDocxSelectionStyle(paper, 'color', colIn.value, true)
  });
  const hiIn = h('input', {
    type: 'color',
    value: '#FEF08A',
    title: tr('تظليل النص'),
    onpointerdown: () => capturePaperSel(),
    oninput: () => applyDocxSelectionStyle(paper, 'backgroundColor', hiIn.value, false),
    onchange: () => applyDocxSelectionStyle(paper, 'backgroundColor', hiIn.value, true)
  });
  const bPainter = btn('🖌 نسخ التنسيق', () => toggleDocxFormatPainter(bPainter), '', { keep: true });
  const bClear = btn('🧹', () => clearDocxSelectionFormatting(paper), '', { keep: true, title: 'مسح التنسيق' });

  const bAlignR = btn('يمين', () => execDocCmd('justifyRight'), '', { keep: true });
  const bAlignC = btn('وسط', () => execDocCmd('justifyCenter'), '', { keep: true });
  const bAlignL = btn('يسار', () => execDocCmd('justifyLeft'), '', { keep: true });
  const bAlignJ = btn('ضبط', () => execDocCmd('justifyFull'), '', { keep: true });

  const bRtl = btn('RTL', () => setDocxBlockDirection(paper, 'rtl'), '', { keep: true, title: 'اتجاه RTL' });
  const bLtr = btn('LTR', () => setDocxBlockDirection(paper, 'ltr'), '', { keep: true, title: 'اتجاه LTR' });

  const bUl = btn('• قائمة', () => { execDocCmd('insertUnorderedList'); updateDocxPagination(paper); }, '', { keep: true, title: 'قائمة نقطية' });
  const bOl = btn('1. ترقيم', () => { execDocCmd('insertOrderedList'); updateDocxPagination(paper); }, '', { keep: true, title: 'قائمة رقمية' });
  const bIndentPlus = btn('⇥', () => adjustDocxBlockIndent(paper, 24), '', { keep: true, title: 'زيادة المسافة البادئة' });
  const bIndentMinus = btn('⇤', () => adjustDocxBlockIndent(paper, -24), '', { keep: true, title: 'إنقاص المسافة البادئة' });

  const lnSpcSel = h('select', { title: tr('تباعد الأسطر') },
    h('option', { value: '1.15' }, '1.15'),
    h('option', { value: '1.5' }, '1.5'),
    h('option', { value: '2.0' }, '2.0'),
    h('option', { value: '1.0' }, '1.0')
  );
  lnSpcSel.onchange = () => { setDocxBlockLineHeight(paper, lnSpcSel.value); updateDocxPagination(paper); };

  const bImg = btn('🖼 صورة', () => { saveDocxSelection(paper); triggerImageInsert((info) => { insertDocxImage(paper, info); updateDocxPagination(paper); }); }, '', { keep: true });
  const bTbl = btn('▦ جدول', () => {
    saveDocxSelection(paper);
    const rStr = prompt(currentLanguage === 'en' ? 'Number of rows (1-20):' : 'عدد الصفوف (1-20):', '3');
    if (rStr == null) return;
    const cStr = prompt(currentLanguage === 'en' ? 'Number of columns (1-12):' : 'عدد الأعمدة (1-12):', '3');
    if (cStr == null) return;
    const rows = Math.max(1, Math.min(20, parseInt(rStr, 10) || 3));
    const cols = Math.max(1, Math.min(12, parseInt(cStr, 10) || 3));
    insertDocxTable(paper, rows, cols);
    refreshDocxContextBar();
    updateDocxPagination(paper);
  }, '', { keep: true });
  const bLink = btn('🔗 رابط', () => insertDocxHyperlink(paper), '', { keep: true });
  const bPageBrk = btn('📄 فاصل صفحة', () => { insertDocxPageBreak(paper); updateDocxStats(); }, '', { keep: true });
  const bToc = btn('📑 فهرس تلقائي', () => { insertDocxAutoToc(paper); updateDocxPagination(paper); });
  const bCallout = btn('💬 صندوق تنبيه', () => { insertDocxCalloutBox(paper); updateDocxPagination(paper); });
  const bChecklist = btn('☑️ مهام', () => { insertDocxChecklistItem(paper); updateDocxPagination(paper); }, '', { keep: true });
  const bDate = btn('📅 التاريخ', () => insertDocxDateStamp(paper), '', { keep: true });
  const bHr = btn('➖ خط فاصل', () => { insertDocxHorizontalLine(paper); updateDocxPagination(paper); });
  const bMd = btn('📝 استيراد Markdown', () => openMarkdownImportModal(paper));

  const bCols = btn('📰 عمودين', () => {
    const on = paper.classList.toggle('cols-2');
    bCols.classList.toggle('on', on);
    updateDocxPagination(paper);
  }, '', { keep: true });

  const bFootnote = btn('¹ حاشية سفلية', () => {
    const note = prompt(currentLanguage === 'en' ? 'Footnote text:' : 'اكتب نص الحاشية السفلية:');
    if (!note || !note.trim()) return;
    let fnBox = $('.doc-footnotes-box', paper);
    if (!fnBox) {
      fnBox = h('div', { class: 'doc-footnotes-box' });
      paper.append(fnBox);
    }
    const num = $$('.doc-fn-item', fnBox).length + 1;
    const sup = h('sup', { class: 'doc-fn-ref', contenteditable: 'false' }, `[${num}]`);
    const sel = window.getSelection();
    if (sel && sel.rangeCount && paper.contains(sel.anchorNode)) {
      const r = sel.getRangeAt(0);
      r.collapse(false);
      r.insertNode(sup);
    } else {
      paper.append(sup);
    }
    fnBox.append(h('div', { class: 'doc-fn-item' }, `${num}. ${note.trim()}`));
    syncPaperToDocxXml(paper);
    pushHistory('docFootnote');
    updateDocxPagination(paper);
  }, '', { keep: true });

  const bComment = btn('💬 تعليق', () => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed || !paper.contains(sel.anchorNode)) {
      toast(currentLanguage === 'en' ? 'Select text first' : 'حدد نصاً أولاً لإضافة تعليق');
      return;
    }
    const txt = prompt(currentLanguage === 'en' ? 'Comment:' : 'اكتب تعليقك على النص المحدد:');
    if (!txt || !txt.trim()) return;
    const range = sel.getRangeAt(0);
    const mark = h('span', { class: 'doc-comment-mark', 'data-comment': txt.trim() });
    try { range.surroundContents(mark); } catch (_) { const c = range.extractContents(); mark.append(c); range.insertNode(mark); }
    openDocxCommentsDrawer(paper);
    syncPaperToDocxXml(paper);
    pushHistory('docComment');
  }, '', { keep: true });

  const pageColIn = h('input', {
    type: 'color',
    value: curDocx.pageColor || '#FFFFFF',
    title: tr('لون الورقة'),
    onchange: () => {
      curDocx.pageColor = pageColIn.value;
      updatePaperClasses();
      applyDocxPageSetupToXml();
      pushHistory('docPageColor');
    }
  });
  const bBorder = btn('إطار الصفحة', () => {
    curDocx.hasBorder = !curDocx.hasBorder;
    bBorder.classList.toggle('on', curDocx.hasBorder);
    updatePaperClasses();
    applyDocxPageSetupToXml();
    pushHistory('docBorder');
  }, curDocx.hasBorder ? 'on' : '');

  const bWatermark = btn('علامة مائية', () => {
    const txt = prompt(tr('اكتب نص العلامة المائية (اتركه فارغاً للإزالة):'), curDocx.watermark || (currentLanguage === 'en' ? 'DRAFT' : 'مسودة'));
    if (txt == null) return;
    curDocx.watermark = txt.trim();
    updatePaperClasses();
    toast('تم تحديث العلامة المائية');
  });

  const bHfToggle = btn('رأس وتذييل', () => {
    curDocx.showHf = !curDocx.showHf;
    hfPanel.hidden = !curDocx.showHf;
  }, curDocx.showHf ? 'on' : '');

  const pgSizeSel = h('select', { title: tr('حجم الورقة:') },
    h('option', { value: 'a4' }, 'A4'),
    h('option', { value: 'letter' }, 'Letter')
  );
  pgSizeSel.value = curDocx.pageSize;
  pgSizeSel.onchange = () => {
    curDocx.pageSize = pgSizeSel.value;
    updatePaperClasses();
    applyDocxPageSetupToXml();
    pushHistory('docPageSize');
  };

  const orientSel = h('select', null,
    h('option', { value: 'portrait' }, tr('عمودي')),
    h('option', { value: 'landscape' }, tr('أفقي'))
  );
  orientSel.value = curDocx.orientation;
  orientSel.onchange = () => {
    curDocx.orientation = orientSel.value;
    updatePaperClasses();
    applyDocxPageSetupToXml();
    pushHistory('docOrient');
  };

  const marginSel = h('select', null,
    h('option', { value: 'normal' }, tr('هوامش عادية')),
    h('option', { value: 'narrow' }, tr('هوامش ضيقة')),
    h('option', { value: 'wide' }, tr('هوامش واسعة'))
  );
  marginSel.value = curDocx.margins;
  marginSel.onchange = () => {
    curDocx.margins = marginSel.value;
    updatePaperClasses();
    applyDocxPageSetupToXml();
    pushHistory('docMargin');
  };

  const docZoomSel = h('select', { title: tr('تكبير') },
    ...['75', '90', '100', '115', '130'].map(z => h('option', { value: z, selected: z === curDocx.zoom }, z + '%'))
  );
  docZoomSel.onchange = () => {
    curDocx.zoom = docZoomSel.value;
    updatePaperClasses();
  };

  const bToPptx = btn('🔄 تحويل إلى عرض PPTX', () => exportDocxToPptx(paper));
  const bToPdfEd = btn('🔄 تحويل إلى PDF', () => convertDocxToPdfEditor(paper));
  const bPrint = btn('🖨 طباعة', () => window.print());
  const bPdf = btn('📄 تصدير PDF', () => exportDocxToPdf(paper), 'pri');

  // Grouped cleanly by topic (Typography → Paragraph/Lists → Insert → Page Setup → Export) without visual separators
  tb.append(
    headingSel, fontSel, sizeSel, bBold, bItalic, bUnder, bStrike, bSup, bSub, colIn, hiIn, bPainter, bClear,
    bAlignR, bAlignC, bAlignL, bAlignJ, bRtl, bLtr, bUl, bOl, bIndentPlus, bIndentMinus, lnSpcSel,
    bImg, bTbl, bLink, bChecklist, bDate, bCallout, bFootnote, bComment, bToc, bHr, bPageBrk, bMd,
    pgSizeSel, orientSel, marginSel, bCols, bBorder, pageColIn, bWatermark, bHfToggle, docZoomSel,
    bPrint, bToPdfEd, bToPptx, bPdf
  );

  const hfInit = getDocxHeaderFooterText();
  const hdrIn = h('input', { type: 'text', placeholder: tr('نص رأس الصفحة…'), value: hfInit.headerTxt });
  const ftrIn = h('input', { type: 'text', placeholder: tr('نص تذييل الصفحة…'), value: hfInit.footerTxt });
  const pgNumChk = h('input', { type: 'checkbox', checked: hfInit.hasPageNum });
  hfPanel.append(
    hdrIn, ftrIn,
    h('label', { style: { display: 'inline-flex', alignItems: 'center', gap: '5px', fontSize: '.85rem' } }, pgNumChk, tr('ترقيم الصفحات')),
    btn('حفظ الرأس والتذييل', () => {
      setDocxHeaderFooter(hdrIn.value, ftrIn.value, pgNumChk.checked);
      renderHeaderFooterOverlay();
      pushHistory('docHeaderFooter');
      toast('تم حفظ الرأس والتذييل');
    }, 'pri')
  );

  const doc = pk.doc(pk.mainPart);
  const body = kid(doc.documentElement, W, 'body');
  const rels = readRels(pk, pk.mainPart);

  const renderDocxRunsInto = (parentEl, containerNode) => {
    let hasContent = false;
    for (const child of Array.from(containerNode.childNodes)) {
      if (child.nodeType !== 1) continue;
      if (child.namespaceURI === W && child.localName === 'bookmarkStart') {
        parentEl.append(h('span', { 'data-docx-bookmark-start': child.getAttributeNS(W, 'id') || child.getAttribute('w:id') || '0', 'data-docx-bookmark-name': child.getAttributeNS(W, 'name') || child.getAttribute('w:name') || '', contenteditable: 'false' }));
        hasContent = true; continue;
      }
      if (child.namespaceURI === W && child.localName === 'bookmarkEnd') {
        parentEl.append(h('span', { 'data-docx-bookmark-end': child.getAttributeNS(W, 'id') || child.getAttribute('w:id') || '0', contenteditable: 'false' }));
        hasContent = true; continue;
      }
      if (child.namespaceURI === W && child.localName === 'hyperlink') {
        const rid = child.getAttributeNS(R, 'id');
        const rel = rid && rels.get(rid);
        const bookmark = child.getAttributeNS(W, 'anchor') || child.getAttribute('w:anchor');
        const aEl = h('a', { href: bookmark ? `#${bookmark}` : safeUrl(rel?.target), target: bookmark ? null : '_blank', rel: bookmark ? null : 'noopener noreferrer', style: { color: '#0563C1', textDecoration: 'underline' } });
        if (renderDocxRunsInto(aEl, child)) {
          parentEl.append(aEl);
          hasContent = true;
        }
        continue;
      }
      if (child.namespaceURI !== W || child.localName !== 'r') continue;

      const rPr = kid(child, W, 'rPr');
      const isB = !!(rPr && kid(rPr, W, 'b') && kid(rPr, W, 'b').getAttributeNS(W, 'val') !== '0');
      const isI = !!(rPr && kid(rPr, W, 'i') && kid(rPr, W, 'i').getAttributeNS(W, 'val') !== '0');
      const isU = !!(rPr && kid(rPr, W, 'u') && kid(rPr, W, 'u').getAttributeNS(W, 'val') !== 'none');
      const isStrike = !!(rPr && kid(rPr, W, 'strike') && kid(rPr, W, 'strike').getAttributeNS(W, 'val') !== '0');
      const vAlign = rPr && kid(rPr, W, 'vertAlign')?.getAttributeNS(W, 'val');
      const colVal = rPr && kid(rPr, W, 'color')?.getAttributeNS(W, 'val');
      const shdFill = rPr && kid(rPr, W, 'shd')?.getAttributeNS(W, 'fill');
      const hiVal = rPr && kid(rPr, W, 'highlight')?.getAttributeNS(W, 'val');
      const rFonts = rPr && kid(rPr, W, 'rFonts');
      const fontName = rFonts && (rFonts.getAttributeNS(W, 'cs') || rFonts.getAttributeNS(W, 'ascii') || rFonts.getAttributeNS(W, 'hAnsi'));
      const szEl = rPr && kid(rPr, W, 'sz');
      const szHalfPt = szEl && +(szEl.getAttributeNS(W, 'val') || szEl.getAttribute('val'));

      const br = kid(child, W, 'br');
      if (br && (br.getAttributeNS(W, 'type') === 'page' || br.getAttribute('type') === 'page')) {
        parentEl.append(h('div', { class: 'doc-page-break', contenteditable: 'false' }, h('span', null, tr('فاصل صفحة'))));
        hasContent = true;
      } else if (br) {
        parentEl.append(h('br'));
        hasContent = true;
      }

      const blips = desc(child, A, 'blip');
      for (const b of blips) {
        const rid = b.getAttributeNS(R, 'embed');
        const rel = rid && rels.get(rid);
        const u8 = rel && rel.path && pk.files.get(rel.path);
        const url = u8 && mediaURL(u8, rel.path);
        if (url && rel.path) {
          const extEl = desc(child, NS.wp, 'extent')[0];
          const cx = extEl ? +extEl.getAttribute('cx') : 0;
          const img = h('img', {
            src: url,
            alt: '',
            'data-media-path': rel.path,
            style: cx > 0 ? { width: Math.round(cx / 9525) + 'px' } : null
          });
          parentEl.append(img);
          hasContent = true;
        }
      }

      const txt = desc(child, W, 't').map(t => t.textContent).join('');
      if (txt) {
        const sp = h('span', null, txt);
        if (isB) sp.style.fontWeight = '700';
        if (isI) sp.style.fontStyle = 'italic';
        const dec = [];
        if (isU) dec.push('underline');
        if (isStrike) dec.push('line-through');
        if (dec.length) sp.style.textDecoration = dec.join(' ');
        if (vAlign === 'superscript') sp.style.verticalAlign = 'super';
        else if (vAlign === 'subscript') sp.style.verticalAlign = 'sub';
        if (colVal && /^[0-9a-fA-F]{6}$/.test(colVal)) sp.style.color = '#' + colVal;
        if (shdFill && /^[0-9a-fA-F]{6}$/.test(shdFill)) sp.style.backgroundColor = '#' + shdFill;
        else if (hiVal === 'yellow') sp.style.backgroundColor = '#FEF08A';
        if (fontName) sp.style.fontFamily = `"${fontName}", sans-serif`;
        if (szHalfPt && szHalfPt > 0) sp.style.fontSize = (szHalfPt / 2) + 'pt';
        parentEl.append(sp);
        hasContent = true;
      }
    }
    return hasContent;
  };

  const renderDocxParagraph = (pNode) => {
    const pPr = kid(pNode, W, 'pPr');
    const numPr = pPr && kid(pPr, W, 'numPr');
    const pStyle = pPr && kid(pPr, W, 'pStyle')?.getAttributeNS(W, 'val');
    const pBdr = pPr && kid(pPr, W, 'pBdr');
    const pShd = pPr && kid(pPr, W, 'shd');

    if (pBdr && kid(pBdr, W, 'bottom') && !kid(pBdr, W, 'left') && !kid(pBdr, W, 'right') && !desc(pNode, W, 't').length) {
      return { el: h('hr', { class: 'doc-hr', contenteditable: 'false' }), listType: null };
    }

    const isCallout = !!(pShd && (kid(pBdr, W, 'left') || kid(pBdr, W, 'right')));
    const pEl = h('div', { class: isCallout ? 'dp doc-callout' : 'dp' });

    if (pStyle) {
      if (/Heading1/i.test(pStyle)) { pEl.dataset.heading = 'h1'; pEl.style.fontSize = '22pt'; pEl.style.fontWeight = '700'; }
      else if (/Heading2/i.test(pStyle)) { pEl.dataset.heading = 'h2'; pEl.style.fontSize = '17pt'; pEl.style.fontWeight = '700'; }
      else if (/Heading3/i.test(pStyle)) { pEl.dataset.heading = 'h3'; pEl.style.fontSize = '14pt'; pEl.style.fontWeight = '700'; }
    }

    const jc = pPr && kid(pPr, W, 'jc')?.getAttributeNS(W, 'val');
    if (jc) pEl.style.textAlign = ({ left: 'left', center: 'center', right: 'right', both: 'justify' })[jc] || '';
    const bidiEl = pPr && kid(pPr, W, 'bidi');
    const bidiVal = bidiEl && (bidiEl.getAttributeNS(W, 'val') || bidiEl.getAttribute('val'));
    if (bidiEl && bidiVal !== '0' && bidiVal !== 'false') pEl.style.direction = 'rtl';
    else if (bidiVal === '0' || bidiVal === 'false' || jc === 'left') pEl.style.direction = 'ltr';

    const spc = pPr && kid(pPr, W, 'spacing');
    const lineVal = spc && +(spc.getAttributeNS(W, 'line') || spc.getAttribute('line') || 0);
    if (lineVal > 0) pEl.style.lineHeight = String(+(lineVal / 240).toFixed(2));

    const ind = pPr && kid(pPr, W, 'ind');
    const startTwips = ind && +(ind.getAttributeNS(W, 'start') || ind.getAttributeNS(W, 'left') || 0);
    if (startTwips > 0 && !numPr) pEl.style.paddingInlineStart = Math.round(startTwips / 20) + 'pt';

    const hasContent = renderDocxRunsInto(pEl, pNode);
    if (!hasContent) pEl.append('');
    else {
      const firstSp = pEl.firstElementChild;
      const rawTxt = firstSp && firstSp.tagName === 'SPAN' && !firstSp.classList.length ? (firstSp.textContent || '') : '';
      const chkM = /^([☐☑])\s?/.exec(rawTxt);
      if (chkM) {
        const rest = rawTxt.slice(chkM[0].length);
        const chkSpan = h('span', { class: 'doc-check-box', contenteditable: 'false' }, chkM[1] + ' ');
        pEl.classList.add('doc-check-item');
        if (chkM[1] === '☑') pEl.classList.add('doc-check-done');
        if (rest) {
          firstSp.textContent = rest;
          pEl.insertBefore(chkSpan, firstSp);
        } else {
          firstSp.replaceWith(chkSpan);
        }
      }
    }

    if (numPr) {
      const numId = kid(numPr, W, 'numId')?.getAttributeNS(W, 'val') || '1';
      return { el: pEl, listType: numId === '2' ? 'ol' : 'ul' };
    }
    return { el: pEl, listType: null };
  };

  if (body) {
    let activeList = null, activeListTag = null;
    const flushList = () => {
      if (activeList) { paper.append(activeList); activeList = null; activeListTag = null; }
    };

    for (const node of kids(body, W)) {
      if (node.localName === 'p') {
        const res = renderDocxParagraph(node);
        if (res.listType) {
          if (!activeList || activeListTag !== res.listType) {
            flushList();
            activeListTag = res.listType;
            activeList = h(res.listType);
          }
          const li = h('li');
          while (res.el.firstChild) li.append(res.el.firstChild);
          activeList.append(li);
        } else {
          flushList();
          paper.append(res.el);
        }
      } else if (node.localName === 'tbl') {
        flushList();
        const tw = h('div', { class: 'tw' });
        const tblPr = kid(node, W, 'tblPr');
        const tblBorders = tblPr && kid(tblPr, W, 'tblBorders');
        const noBorder = tblBorders && kid(tblBorders, W, 'top')?.getAttributeNS(W, 'val') === 'none';
        const tbl = h('table', { class: 'dt' + (noBorder ? '' : ' grid') });

        for (const trRow of kids(node, W, 'tr')) {
          const row = h('tr');
          for (const tc of kids(trRow, W, 'tc')) {
            const td = h('td');
            const tcPr = kid(tc, W, 'tcPr');
            const gridSpan = tcPr && +(kid(tcPr, W, 'gridSpan')?.getAttributeNS(W, 'val') || 1);
            if (gridSpan > 1) td.colSpan = gridSpan;
            const vAlign = tcPr && kid(tcPr, W, 'vAlign')?.getAttributeNS(W, 'val');
            if (vAlign === 'center') td.style.verticalAlign = 'middle';
            const shd = tcPr && kid(tcPr, W, 'shd')?.getAttributeNS(W, 'fill');
            if (shd && /^[0-9a-fA-F]{6}$/.test(shd)) td.style.backgroundColor = '#' + shd;

            const ps = kids(tc, W, 'p');
            if (ps.length) {
              ps.forEach(p => td.append(renderDocxParagraph(p).el));
            } else {
              td.append(h('div', { class: 'dp' }, desc(tc, W, 't').map(t => t.textContent).join(' ')));
            }
            row.append(td);
          }
          tbl.append(row);
        }
        tw.append(tbl);
        paper.append(tw);
      }
    }
    flushList();
  }
  if (!paper.childNodes.length) paper.append(h('div', { class: 'dp' }, ''));
  renderWatermarkOverlay();
  renderHeaderFooterOverlay();

  const getActiveDocxContextTarget = () => {
    if (curDocx.activeImg && paper.contains(curDocx.activeImg)) return curDocx.activeImg;
    if (curDocx.activeTd && paper.contains(curDocx.activeTd)) {
      return curDocx.activeTd.closest('table.dt') || curDocx.activeTd;
    }
    return null;
  };

  updateDocxContextBarVisibility = () => {
    if (!paper.isConnected) return;
    const targetEl = getActiveDocxContextTarget();
    if (!targetEl || !ctxBar.childElementCount) {
      ctxBar.hidden = true;
      return;
    }
    const rect = targetEl.getBoundingClientRect();
    if (rect.width <= 0 && rect.height <= 0) {
      ctxBar.hidden = true;
      return;
    }
    const tbRect = tb.getBoundingClientRect();
    const topEl = $('#editor .top');
    const headerBottom = topEl ? topEl.getBoundingClientRect().bottom : 0;
    const barPos = window.getComputedStyle(stickyBars).position;
    const isSticky = barPos === 'sticky' || barPos === '-webkit-sticky';
    const topCutoff = isSticky ? Math.max(headerBottom, tbRect.bottom + 8) : Math.max(0, headerBottom);
    const viewH = window.innerHeight || document.documentElement.clientHeight || 800;
    const inWindow = rect.bottom > topCutoff + 6 && rect.top < viewH - 8;
    const wrapRect = paperWrap.getBoundingClientRect();
    const inWrap = rect.bottom > wrapRect.top + 4 && rect.top < wrapRect.bottom - 4;
    ctxBar.hidden = !(inWindow && inWrap);
  };

  refreshDocxContextBar = () => {
    ctxBar.innerHTML = '';
    if (curDocx.activeImg && paper.contains(curDocx.activeImg)) {
      const img = curDocx.activeImg;
      ctxBar.append(
        h('strong', null, tr('أدوات الصورة:')),
        btn('صغير 25%', () => { img.style.width = '25%'; syncPaperToDocxXml(paper); pushHistory('docImgSize'); updateDocxPagination(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        btn('متوسط 50%', () => { img.style.width = '50%'; syncPaperToDocxXml(paper); pushHistory('docImgSize'); updateDocxPagination(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        btn('كبير 75%', () => { img.style.width = '75%'; syncPaperToDocxXml(paper); pushHistory('docImgSize'); updateDocxPagination(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        btn('كامل 100%', () => { img.style.width = '100%'; syncPaperToDocxXml(paper); pushHistory('docImgSize'); updateDocxPagination(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        h('span', { class: 'sep' }),
        btn('التفاف يمين', () => { img.classList.remove('wrap-left'); img.classList.toggle('wrap-right'); syncPaperToDocxXml(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        btn('التفاف يسار', () => { img.classList.remove('wrap-right'); img.classList.toggle('wrap-left'); syncPaperToDocxXml(paper); updateDocxContextBarVisibility(); }, '', { keep: true }),
        h('span', { class: 'sep' }),
        btn('يمين', () => { if (img.parentElement) img.parentElement.style.textAlign = 'right'; syncPaperToDocxXml(paper); pushHistory('docImgAlign'); }, '', { keep: true }),
        btn('وسط', () => { if (img.parentElement) img.parentElement.style.textAlign = 'center'; syncPaperToDocxXml(paper); pushHistory('docImgAlign'); }, '', { keep: true }),
        btn('يسار', () => { if (img.parentElement) img.parentElement.style.textAlign = 'left'; syncPaperToDocxXml(paper); pushHistory('docImgAlign'); }, '', { keep: true }),
        h('span', { class: 'sep' }),
        btn('حذف الصورة', () => { img.remove(); curDocx.activeImg = null; refreshDocxContextBar(); syncPaperToDocxXml(paper); pushHistory('docImgDel'); updateDocxPagination(paper); }, 'danger', { keep: true })
      );
      observeActiveDocxTarget();
      updateDocxContextBarVisibility();
      return;
    }

    if (curDocx.activeTd && paper.contains(curDocx.activeTd)) {
      const td = curDocx.activeTd;
      const tbl = td.closest('table.dt');
      const cellBgIn = h('input', { type: 'color', value: '#E2E8F0', title: tr('لون الخلية المحددة') });
      ctxBar.append(
        h('strong', null, tr('أدوات الجدول:')),
        btn('➕ صف', () => { modifyDocxTable(paper, tbl, td, 'addRow'); refreshDocxContextBar(); updateDocxPagination(paper); }, '', { keep: true }),
        btn('➖ صف', () => { modifyDocxTable(paper, tbl, td, 'delRow'); refreshDocxContextBar(); updateDocxPagination(paper); }, 'danger', { keep: true }),
        btn('➕ عمود', () => { modifyDocxTable(paper, tbl, td, 'addCol'); refreshDocxContextBar(); }, '', { keep: true }),
        btn('➖ عمود', () => { modifyDocxTable(paper, tbl, td, 'delCol'); refreshDocxContextBar(); }, 'danger', { keep: true }),
        btn('🎨 تنسيق رأس الجدول', () => {
          if (!tbl) return;
          $$('tr', tbl).forEach((trRow, ri) => {
            $$('td', trRow).forEach(cell => {
              if (ri === 0) {
                cell.style.backgroundColor = '#B4531A';
                cell.style.color = '#FFFFFF';
                cell.style.fontWeight = '700';
              } else if (ri % 2 === 1) {
                cell.style.backgroundColor = '#FAF7F1';
              }
            });
          });
          syncPaperToDocxXml(paper);
          pushHistory('docTblStyle');
        }, '', { keep: true }),
        btn('دمج مع الخلية التالية', () => {
          const nextTd = td.nextElementSibling;
          if (!nextTd) return;
          td.colSpan = (td.colSpan || 1) + (nextTd.colSpan || 1);
          while (nextTd.firstChild) td.append(nextTd.firstChild);
          nextTd.remove();
          syncPaperToDocxXml(paper);
          pushHistory('docMergeCell');
        }, '', { keep: true }),
        btn('توسيط عمودي', () => {
          td.style.verticalAlign = td.style.verticalAlign === 'middle' ? 'top' : 'middle';
          syncPaperToDocxXml(paper);
          pushHistory('docCellVAlign');
        }, '', { keep: true }),
        h('span', { class: 'sep' }),
        cellBgIn,
        btn('لون الخلية المحددة', () => { td.style.backgroundColor = cellBgIn.value; syncPaperToDocxXml(paper); pushHistory('docCellBg'); }, '', { keep: true }),
        btn('إخفاء/إظهار الحدود', () => { tbl.classList.toggle('grid'); syncPaperToDocxXml(paper); pushHistory('docTblBorder'); }, '', { keep: true }),
        btn('🗑 حذف الجدول', () => {
          const twWrap = tbl?.closest('.tw') || tbl;
          if (twWrap) twWrap.remove();
          curDocx.activeTd = null;
          refreshDocxContextBar();
          syncPaperToDocxXml(paper);
          pushHistory('docTblDel');
          updateDocxPagination(paper);
          toast('تم حذف الجدول');
        }, 'danger', { keep: true })
      );
      observeActiveDocxTarget();
      updateDocxContextBarVisibility();
      return;
    }

    if (activeTargetObserver) { activeTargetObserver.disconnect(); activeTargetObserver = null; }
    ctxBar.hidden = true;
  };

  paper.addEventListener('pointerdown', (e) => {
    $$('.docx-img-sel', paper).forEach(x => x.classList.remove('docx-img-sel'));
    $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
    curDocx.activeImg = null;
    curDocx.activeTd = null;

    const img = e.target.closest('img');
    if (img && paper.contains(img)) {
      img.classList.add('docx-img-sel');
      curDocx.activeImg = img;
    } else {
      const td = e.target.closest('td');
      if (td && paper.contains(td)) {
        td.classList.add('doc-td-active');
        curDocx.activeTd = td;
      }
    }
    refreshDocxContextBar();
  });

  paper.addEventListener('pointerup', () => {
    if (curDocx.copiedStyle) applyDocxFormatPainter(paper, bPainter);
    updateDocxStats();
  });

  const updateDocxStats = () => {
    const raw = (paper.innerText || '').trim();
    const words = raw ? raw.split(/\s+/).filter(Boolean).length : 0;
    const chars = raw.replace(/\s+/g, '').length;
    const paras = Math.max(1, $$('.dp, li', paper).length);
    const readMin = Math.max(1, Math.ceil(words / 200));
    const pageCount = updateDocxPagination(paper);

    const selTxt = (window.getSelection()?.toString() || '').trim();
    const selWords = selTxt ? selTxt.split(/\s+/).filter(Boolean).length : 0;

    statusBar.innerHTML = '';
    statusBar.append(
      h('span', { id: 'docPageCountStat' }, tr('الصفحات:'), ' ', h('strong', null, String(pageCount))),
      h('span', null, tr('الكلمات:'), ' ', h('strong', null, String(words))),
      selWords > 0 ? h('span', { style: { color: 'var(--acc)' } }, tr('المحدد:'), ' ', h('strong', null, String(selWords))) : null,
      h('span', null, tr('الأحرف:'), ' ', h('strong', null, String(chars))),
      h('span', null, tr('الفقرات:'), ' ', h('strong', null, String(paras))),
      h('span', null, tr('وقت القراءة:'), ' ', h('strong', null, `${readMin} ${tr('دقيقة')}`))
    );
  };

  paper.addEventListener('input', () => {
    updateDocxStats();
    syncPaperToDocxXml(paper);
    scheduleHistory('editDoc', 450);
  });

  const onDocxOutsidePointerDown = (ev) => {
    if (!paper.isConnected) {
      document.removeEventListener('pointerdown', onDocxOutsidePointerDown);
      return;
    }
    if (paper.contains(ev.target) || stickyBars.contains(ev.target)) return;
    if (curDocx.activeImg || curDocx.activeTd) {
      $$('.docx-img-sel', paper).forEach(x => x.classList.remove('docx-img-sel'));
      $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
      curDocx.activeImg = null;
      curDocx.activeTd = null;
      refreshDocxContextBar();
    }
  };
  document.addEventListener('pointerdown', onDocxOutsidePointerDown);

  let activeTargetObserver = null;
  const observeActiveDocxTarget = () => {
    if (activeTargetObserver) {
      activeTargetObserver.disconnect();
      activeTargetObserver = null;
    }
    const targetEl = getActiveDocxContextTarget();
    if (!targetEl || typeof IntersectionObserver === 'undefined') return;
    activeTargetObserver = new IntersectionObserver(() => {
      updateDocxContextBarVisibility();
    }, { threshold: [0, 0.01, 0.1, 0.25, 0.5, 1] });
    activeTargetObserver.observe(targetEl);
  };

  let ctxVisRaf = 0;
  let ctxVisSettleTimer = 0;
  const scheduleCtxVisCheck = () => {
    if (!paper.isConnected) {
      if (activeTargetObserver) { activeTargetObserver.disconnect(); activeTargetObserver = null; }
      window.removeEventListener('scroll', scheduleCtxVisCheck, true);
      window.removeEventListener('scrollend', scheduleCtxVisCheck, true);
      window.removeEventListener('resize', scheduleCtxVisCheck);
      return;
    }
    if (!ctxVisRaf) {
      ctxVisRaf = requestAnimationFrame(() => {
        ctxVisRaf = 0;
        updateDocxContextBarVisibility();
      });
    }
    clearTimeout(ctxVisSettleTimer);
    ctxVisSettleTimer = setTimeout(updateDocxContextBarVisibility, 140);
  };
  window.addEventListener('scroll', scheduleCtxVisCheck, { passive: true, capture: true });
  window.addEventListener('scrollend', scheduleCtxVisCheck, { passive: true, capture: true });
  window.addEventListener('resize', scheduleCtxVisCheck, { passive: true });

  paperWrap.append(ruler, paper, statusBar);
  container.append(stickyBars, paperWrap);
  app.append(container);
  setTimeout(() => { updateDocxStats(); refreshDocxContextBar(); }, 30);
}

function toggleDocxFormatPainter(btnEl) {
  if (curDocx.copiedStyle) {
    curDocx.copiedStyle = null;
    btnEl?.classList.remove('on');
    return;
  }
  const sel = window.getSelection();
  const node = sel?.anchorNode;
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  if (!el) return;
  const cs = window.getComputedStyle(el);
  curDocx.copiedStyle = {
    color: cs.color,
    backgroundColor: cs.backgroundColor,
    fontWeight: cs.fontWeight,
    fontStyle: cs.fontStyle,
    textDecoration: cs.textDecoration,
    fontFamily: cs.fontFamily
  };
  btnEl?.classList.add('on');
  toast('انقر الآن على العنصر الهدف لتطبيق التنسيق عليه');
}

function applyDocxFormatPainter(paper, btnEl) {
  const st = curDocx.copiedStyle;
  curDocx.copiedStyle = null;
  btnEl?.classList.remove('on');
  if (!st) return;
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount || sel.isCollapsed) return;
  const range = sel.getRangeAt(0);
  const span = document.createElement('span');
  span.style.color = st.color;
  if (st.backgroundColor && st.backgroundColor !== 'rgba(0, 0, 0, 0)' && st.backgroundColor !== 'rgb(255, 255, 255)') {
    span.style.backgroundColor = st.backgroundColor;
  }
  span.style.fontWeight = st.fontWeight;
  span.style.fontStyle = st.fontStyle;
  span.style.textDecoration = st.textDecoration;
  span.style.fontFamily = st.fontFamily;
  try { range.surroundContents(span); } catch (_) { const c = range.extractContents(); span.appendChild(c); range.insertNode(span); }
  syncPaperToDocxXml(paper);
  pushHistory('docFormatPainter');
  toast('تم نسخ التنسيق بنجاح');
}

function insertDocxAutoToc(paper) {
  const headings = $$('[data-heading]', paper);
  if (!headings.length) {
    toast('لا توجد عناوين (Heading 1-3) لبناء الفهرس');
    return;
  }
  $$('.doc-toc', paper).forEach(t => t.remove());
  const firstHeadingDir = headings[0]?.style?.direction || (currentLanguage === 'ar' ? 'rtl' : 'ltr');
  const tocBox = h('div', {
    class: 'doc-toc',
    dir: firstHeadingDir,
    contenteditable: 'false',
    style: { direction: firstHeadingDir, textAlign: 'start' }
  },
    h('div', { class: 'doc-toc-title' }, '📑 ' + tr('فهرس المحتويات'))
  );
  headings.forEach((hd, i) => {
    const lvl = hd.dataset.heading === 'h1' ? 0 : hd.dataset.heading === 'h2' ? 16 : 32;
    const item = h('div', { class: 'doc-toc-item', style: { paddingInlineStart: lvl + 'px' } }, `${i + 1}. ${hd.innerText.trim()}`);
    tocBox.append(item);
  });
  paper.insertBefore(tocBox, paper.firstChild);
  syncPaperToDocxXml(paper);
  pushHistory('docToc');
  toast('تم إدراج فهرس المحتويات');
}

function insertDocxCalloutBox(paper) {
  const box = h('div', { class: 'dp doc-callout' }, h('b', null, '💡 ملاحظة: '), 'اكتب نص التنبيه أو الاقتباس هنا…');
  const blocks = getDocxTargetBlocks(paper);
  if (blocks.length && blocks[0].parentNode === paper) insertAfter(blocks[0], box);
  else paper.append(box);
  syncPaperToDocxXml(paper);
  pushHistory('docCallout');
}

function insertDocxChecklistItem(paper) {
  const item = h('div', { class: 'dp doc-check-item' },
    h('span', { class: 'doc-check-box', contenteditable: 'false' }, '☐ '),
    tr('مهمة جديدة')
  );
  const blocks = getDocxTargetBlocks(paper);
  if (blocks.length && blocks[0].parentNode === paper) insertAfter(blocks[0], item);
  else paper.append(item);
  syncPaperToDocxXml(paper);
  pushHistory('docChecklist');
}

function insertDocxDateStamp(paper) {
  ensureActiveDocxSelection(paper);
  const dateStr = new Date().toLocaleDateString(currentLanguage === 'ar' ? 'ar-SA' : 'en-CA', {
    year: 'numeric', month: 'long', day: 'numeric'
  });
  const sel = window.getSelection();
  if (sel && sel.rangeCount && paper.contains(sel.anchorNode)) {
    document.execCommand('insertText', false, dateStr);
  } else {
    paper.append(h('div', { class: 'dp' }, dateStr));
  }
  syncPaperToDocxXml(paper);
  pushHistory('docDate');
}

function insertDocxHorizontalLine(paper) {
  const hr = h('hr', { class: 'doc-hr', contenteditable: 'false' });
  const blocks = getDocxTargetBlocks(paper);
  if (blocks.length && blocks[0].parentNode === paper) {
    insertAfter(blocks[0], hr);
    insertAfter(hr, h('div', { class: 'dp' }, ''));
  } else {
    paper.append(hr, h('div', { class: 'dp' }, ''));
  }
  syncPaperToDocxXml(paper);
  pushHistory('docHr');
}

function openMarkdownImportModal(paper) {
  const backdrop = h('div', { class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) backdrop.remove(); } });
  const ta = h('textarea', { placeholder: '# عنوان رئيسي\n## عنوان فرعي\n- نقطة أولى\n- نقطة ثانية\nنص عادي **عريض**' });
  const modal = h('div', { class: 'md-modal' },
    h('strong', null, tr('الصق نص Markdown هنا لتحويله تلقائياً إلى محتوى منسق:')),
    ta,
    h('div', { class: 'row', style: { justifyContent: 'flex-end' } },
      btn('إلغاء', () => backdrop.remove()),
      btn('تحويل وإدراج', () => {
        const md = ta.value;
        backdrop.remove();
        if (!md.trim()) return;
        importMarkdownIntoDocx(paper, md);
      }, 'pri')
    )
  );
  backdrop.append(modal);
  document.body.append(backdrop);
  ta.focus();
}

function importMarkdownIntoDocx(paper, mdText) {
  const lines = mdText.split(/\r?\n/);
  let curUl = null, curOl = null;
  const flush = () => {
    if (curUl) { paper.append(curUl); curUl = null; }
    if (curOl) { paper.append(curOl); curOl = null; }
  };

  const appendFormattedText = (el, text) => {
    const parts = text.split(/(\*\*[^*]+\*\*)/g);
    parts.forEach(part => {
      if (part.startsWith('**') && part.endsWith('**')) {
        el.append(h('b', null, part.slice(2, -2)));
      } else if (part) {
        el.append(part);
      }
    });
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) { flush(); continue; }
    if (line.startsWith('### ')) {
      flush();
      const d = h('div', { class: 'dp', 'data-heading': 'h3', style: { fontSize: '14pt', fontWeight: '700' } });
      appendFormattedText(d, line.slice(4));
      paper.append(d);
    } else if (line.startsWith('## ')) {
      flush();
      const d = h('div', { class: 'dp', 'data-heading': 'h2', style: { fontSize: '17pt', fontWeight: '700' } });
      appendFormattedText(d, line.slice(3));
      paper.append(d);
    } else if (line.startsWith('# ')) {
      flush();
      const d = h('div', { class: 'dp', 'data-heading': 'h1', style: { fontSize: '22pt', fontWeight: '700' } });
      appendFormattedText(d, line.slice(2));
      paper.append(d);
    } else if (/^[-*]\s+/.test(line)) {
      if (curOl) { paper.append(curOl); curOl = null; }
      if (!curUl) curUl = h('ul');
      const li = h('li');
      appendFormattedText(li, line.replace(/^[-*]\s+/, ''));
      curUl.append(li);
    } else if (/^\d+\.\s+/.test(line)) {
      if (curUl) { paper.append(curUl); curUl = null; }
      if (!curOl) curOl = h('ol');
      const li = h('li');
      appendFormattedText(li, line.replace(/^\d+\.\s+/, ''));
      curOl.append(li);
    } else {
      flush();
      const d = h('div', { class: 'dp' });
      appendFormattedText(d, line);
      paper.append(d);
    }
  }
  flush();
  syncPaperToDocxXml(paper);
  pushHistory('importMarkdown');
  updateDocxPagination(paper);
  toast('تم استيراد Markdown بنجاح');
}

function exportDocxToPptx(paper) {
  const sections = [];
  let curSec = { title: fileName.replace(/\.docx$/i, '') || 'العرض التقديمي', bullets: [] };

  for (const block of kids(paper)) {
    if (block.classList?.contains('doc-toc') || block.classList?.contains('paper-watermark') || block.classList?.contains('doc-page-break')) continue;
    const txt = (block.innerText || '').trim();
    if (!txt) continue;
    if (block.dataset?.heading) {
      if (curSec.bullets.length || sections.length > 0) sections.push(curSec);
      curSec = { title: txt, bullets: [] };
    } else {
      txt.split('\n').forEach(l => { if (l.trim()) curSec.bullets.push(l.trim()); });
    }
  }
  if (curSec.title || curSec.bullets.length) sections.push(curSec);
  if (!sections.length) return;

  const newPk = createNewBlankDeck('light');
  const deck = (() => {
    const prevPk = pk;
    pk = newPk;
    const d = readDeck();
    pk = prevPk;
    return d;
  })();

  pk = newPk;
  const W0 = deck.W, H0 = deck.H;

  sections.forEach((sec, idx) => {
    let slPath;
    if (idx === 0) {
      slPath = 'ppt/slides/slide1.xml';
    } else {
      slPath = uniquePath(pk, 'ppt/slides', 'slide', 'xml');
      pk.setText(slPath, XMLDECL + `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld></p:sld>`);
      addRel(pk, slPath, REL + 'slideLayout', 'ppt/slideLayouts/slideLayout2.xml');
      registerSlidePart(deck, slPath);
    }

    const sDoc = pk.doc(slPath);
    const spTree = spTreeOf(sDoc);
    for (const oldSh of shapesOf(sDoc)) removeNode(oldSh);

    const bodyLines = sec.bullets.slice(0, 8).map(b =>
      `<a:p><a:pPr algn="r" rtl="1" marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="ar-SA" sz="2000"><a:solidFill><a:srgbClr val="1F1C17"/></a:solidFill></a:rPr><a:t>${escXml(b)}</a:t></a:r></a:p>`
    ).join('');

    const titleXml = `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="${Math.round(W0 * 0.08)}" y="${Math.round(H0 * 0.08)}"/><a:ext cx="${Math.round(W0 * 0.84)}" cy="${Math.round(H0 * 0.14)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr lang="ar-SA" sz="3000" b="1"><a:solidFill><a:srgbClr val="B4531A"/></a:solidFill></a:rPr><a:t>${escXml(sec.title)}</a:t></a:r></a:p></p:txBody></p:sp>`;
    spTree.appendChild(frag(sDoc, titleXml));

    if (bodyLines) {
      const bodyXml = `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="${Math.round(W0 * 0.08)}" y="${Math.round(H0 * 0.25)}"/><a:ext cx="${Math.round(W0 * 0.84)}" cy="${Math.round(H0 * 0.65)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr anchor="t"/><a:lstStyle/>${bodyLines}</p:txBody></p:sp>`;
      spTree.appendChild(frag(sDoc, bodyXml));
    }
  });

  activeDraftId = null;
  kind = 'pptx';
  fileName = fileName.replace(/\.(docx|docm)$/i, '') + '.pptx';
  $('#fname').textContent = fileName;
  curPptx = { deck: null, slideIdx: 0, selShapeIdx: null, multiSel: new Set(), selCell: null, activeTab: 'home', clipboardXml: null, clipboardRels: [], clipboardSourcePath: null, copiedStyle: null, viewMode: 'normal', zoom: '100', showGrid: false };
  resetHistory();
  pushHistory('docxToPptx');
  renderEditor();
  toast('تم تحويل المستند إلى عرض تقديمي');
}

function getDocxTargetBlocks(paper) {
  const sel = window.getSelection();
  const blocks = new Set();

  if (sel && sel.rangeCount) {
    const range = sel.getRangeAt(0);
    const findBlock = (n) => {
      const el = n?.nodeType === 1 ? n : n?.parentElement;
      return el?.closest('.dp, li');
    };

    const startB = findBlock(range.startContainer);
    const endB = findBlock(range.endContainer);
    if (startB && paper.contains(startB)) blocks.add(startB);
    if (endB && paper.contains(endB)) blocks.add(endB);

    $$('.dp, li', paper).forEach(el => {
      try { if (range.intersectsNode(el)) blocks.add(el); } catch (_) {}
    });

    if (!blocks.size) {
      let cur = range.startContainer;
      if (cur === paper) {
        const firstReal = kids(paper).find(el => !el.classList?.contains('paper-watermark') && !el.classList?.contains('paper-hf-preview'));
        if (firstReal) {
          if (firstReal.matches?.('.dp, li')) blocks.add(firstReal);
          else {
            const dp = h('div', { class: 'dp' }, '');
            paper.append(dp);
            blocks.add(dp);
          }
        }
      } else {
        while (cur && cur.parentNode !== paper) cur = cur.parentNode;
        if (cur && cur !== paper && !cur.classList?.contains('paper-watermark') && !cur.classList?.contains('paper-hf-preview')) {
          if (cur.nodeType === 3) {
            const dp = h('div', { class: 'dp' }, cur.nodeValue);
            paper.replaceChild(dp, cur);
            blocks.add(dp);
          } else if (cur.tagName === 'DIV' || cur.tagName === 'P') {
            cur.classList.add('dp');
            blocks.add(cur);
          }
        }
      }
    }
  }

  if (!blocks.size) {
    const fallback = $$('.dp, li', paper)[0];
    if (fallback) blocks.add(fallback);
    else {
      const dp = h('div', { class: 'dp' }, '');
      paper.append(dp);
      blocks.add(dp);
    }
  }
  return Array.from(blocks);
}

function applyDocxHeadingStyle(paper, level) {
  const blocks = getDocxTargetBlocks(paper);
  if (!blocks.length) return;
  const map = {
    normal: { size: '', weight: '' },
    h1: { size: '22pt', weight: '700' },
    h2: { size: '17pt', weight: '700' },
    h3: { size: '14pt', weight: '700' }
  };
  const cfg = map[level] || map.normal;
  for (const b of blocks) {
    if (level === 'normal') delete b.dataset.heading;
    else b.dataset.heading = level;
    b.style.fontSize = cfg.size;
    b.style.fontWeight = cfg.weight;
  }
  syncPaperToDocxXml(paper);
  pushHistory('docHeading');
}

function ensureActiveDocxSelection(paper) {
  let sel = window.getSelection();
  if (!sel || !sel.rangeCount || !paper.contains(sel.anchorNode)) {
    restoreDocxSelection(paper);
    sel = window.getSelection();
  }
  return sel;
}

function applyDocxInlineWrapperStyle(paper, styleProp, styleVal) {
  const sel = ensureActiveDocxSelection(paper);
  if (!sel || !sel.rangeCount || sel.isCollapsed || !paper.contains(sel.anchorNode)) return false;
  const range = sel.getRangeAt(0);
  const blocks = $$('.dp, li, td', paper).filter(el => {
    if (el.tagName === 'TD' && el.querySelector('.dp')) return false;
    try { return range.intersectsNode(el); } catch (_) { return false; }
  });

  if (blocks.length > 1 || (blocks.length === 1 && !blocks[0].contains(range.commonAncestorContainer))) {
    for (const blk of blocks) {
      const subRange = document.createRange();
      subRange.selectNodeContents(blk);
      if (blk.contains(range.startContainer)) subRange.setStart(range.startContainer, range.startOffset);
      if (blk.contains(range.endContainer)) subRange.setEnd(range.endContainer, range.endOffset);
      if (subRange.collapsed) continue;
      const span = document.createElement('span');
      span.style[styleProp] = styleVal;
      const content = subRange.extractContents();
      span.appendChild(content);
      subRange.insertNode(span);
    }
    if (blocks.length) {
      const newRange = document.createRange();
      newRange.setStartBefore(blocks[0].firstChild || blocks[0]);
      const lastBlk = blocks[blocks.length - 1];
      newRange.setEndAfter(lastBlk.lastChild || lastBlk);
      sel.removeAllRanges();
      sel.addRange(newRange);
      savedDocxSelectionRange = newRange.cloneRange();
    }
    return true;
  }

  const span = document.createElement('span');
  span.style[styleProp] = styleVal;
  try {
    range.surroundContents(span);
  } catch (_) {
    const fragContent = range.extractContents();
    span.appendChild(fragContent);
    range.insertNode(span);
  }
  const nextRange = document.createRange();
  nextRange.selectNodeContents(span);
  sel.removeAllRanges();
  sel.addRange(nextRange);
  savedDocxSelectionRange = nextRange.cloneRange();
  return true;
}

function applyDocxSelectionStyle(paper, styleProp, styleVal, commitHistory) {
  if (!applyDocxInlineWrapperStyle(paper, styleProp, styleVal)) return;
  syncPaperToDocxXml(paper);
  if (commitHistory) pushHistory('docStyle:' + styleProp);
  else scheduleHistory('docStyle:' + styleProp, 300);
}

function applyDocxSelectionFontSize(paper, pt) {
  if (!applyDocxInlineWrapperStyle(paper, 'fontSize', pt + 'pt')) return;
  syncPaperToDocxXml(paper);
  pushHistory('docFontSize');
}

function clearDocxSelectionFormatting(paper) {
  const sel = ensureActiveDocxSelection(paper);
  if (!sel || !sel.rangeCount || !paper.contains(sel.anchorNode)) return;
  document.execCommand('removeFormat', false, null);
  const activeSel = window.getSelection();
  if (activeSel && activeSel.rangeCount) {
    const r = activeSel.getRangeAt(0);
    const inlineTags = new Set(['SUB', 'SUP', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'FONT']);
    $$('span, sub, sup, b, strong, i, em, u, s, strike, font', paper).forEach(el => {
      if (!el.isConnected || el.closest('.paper-watermark, .paper-hf-preview, .doc-toc, .doc-page-break')) return;
      if (el.classList.contains('doc-comment-mark') || el.classList.contains('doc-fn-ref')) return;
      let hit = false;
      try { hit = r.intersectsNode(el); } catch (_) {}
      if (!hit) return;
      if (inlineTags.has(el.tagName)) {
        el.replaceWith(...Array.from(el.childNodes));
      } else {
        el.style.fontWeight = '';
        el.style.fontStyle = '';
        el.style.textDecoration = '';
        el.style.verticalAlign = '';
        el.style.color = '';
        el.style.backgroundColor = '';
        el.style.fontSize = '';
        el.style.fontFamily = '';
        if (!el.getAttribute('style')?.trim() && !el.attributes.length) {
          el.replaceWith(...Array.from(el.childNodes));
        }
      }
    });
  }
  $$('span', paper).forEach(sp => {
    if (!sp.textContent && !sp.childElementCount && !sp.hasAttribute('data-docx-bookmark-start') && !sp.hasAttribute('data-docx-bookmark-end')) {
      sp.remove();
    }
  });
  normalizePaperBlocks(paper);
  syncPaperToDocxXml(paper);
  pushHistory('fmt:removeFormat');
}

function adjustDocxBlockIndent(paper, deltaPt) {
  ensureActiveDocxSelection(paper);
  const blocks = getDocxTargetBlocks(paper);
  for (const b of blocks) {
    const cur = parseFloat(b.style.paddingInlineStart || '0') || 0;
    const next = Math.max(0, Math.min(240, cur + deltaPt));
    b.style.paddingInlineStart = next > 0 ? next + 'pt' : '';
  }
  syncPaperToDocxXml(paper);
  pushHistory(deltaPt > 0 ? 'docIndent' : 'docOutdent');
  updateDocxPagination(paper);
}

function setDocxBlockDirection(paper, dir) {
  paper.focus();
  const blocks = getDocxTargetBlocks(paper);
  for (const b of blocks) {
    b.style.direction = dir;
    b.style.unicodeBidi = 'isolate';
    b.style.textAlign = dir === 'rtl' ? 'right' : 'left';
    const parentList = b.closest('ul, ol');
    if (parentList) {
      parentList.style.direction = dir;
      parentList.style.textAlign = dir === 'rtl' ? 'right' : 'left';
    }
  }
  syncPaperToDocxXml(paper);
  pushHistory('docDir:' + dir);
}

function setDocxBlockLineHeight(paper, lh) {
  const blocks = getDocxTargetBlocks(paper);
  for (const b of blocks) b.style.lineHeight = lh;
  syncPaperToDocxXml(paper);
  pushHistory('docLineSpc');
}

function insertDocxHyperlink(paper) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) { toast(currentLanguage === 'en' ? 'Place the cursor in text first' : 'ضع المؤشر داخل النص أولاً'); return; }
  const range = sel.getRangeAt(0).cloneRange();
  const anchorNode = sel.anchorNode?.nodeType === 1 ? sel.anchorNode : sel.anchorNode?.parentElement;
  const existing = anchorNode?.closest?.('a');
  const currentUrl = existing?.getAttribute('href') || '';
  const entered = prompt(currentLanguage === 'en' ? 'URL (leave blank to remove the link):' : 'الرابط (اتركه فارغاً لإزالة الرابط):', currentUrl || 'https://');
  if (entered == null) return;
  if (!entered.trim() && existing) {
    existing.replaceWith(...existing.childNodes);
    syncPaperToDocxXml(paper); pushHistory('removeDocHyperlink'); toast(currentLanguage === 'en' ? 'Link removed' : 'تمت إزالة الرابط'); return;
  }
  const url = safeUrl(entered);
  if (!entered.trim() || url === '#') { toast(currentLanguage === 'en' ? 'Enter a valid web or email address' : 'أدخل رابطاً أو بريداً إلكترونياً صحيحاً'); return; }
  if (existing) {
    existing.href = url;
    existing.target = '_blank';
    existing.rel = 'noopener noreferrer';
  } else {
    const a = document.createElement('a');
    a.href = url; a.target = '_blank'; a.rel = 'noopener noreferrer';
    if (!range.collapsed) {
      try { range.surroundContents(a); }
      catch (_) { const content = range.extractContents(); a.appendChild(content); range.insertNode(a); }
    } else {
      const label = prompt(currentLanguage === 'en' ? 'Link text:' : 'النص الظاهر للرابط:');
      if (label == null || !label.trim()) return;
      a.textContent = label.trim(); range.insertNode(a);
    }
  }
  syncPaperToDocxXml(paper); pushHistory('docHyperlink'); toast(currentLanguage === 'en' ? 'Link updated' : 'تم إدراج الرابط بنجاح');
}

function modifyDocxTable(paper, tbl, activeTd, action) {
  if (!tbl || !activeTd) return;
  const tr = activeTd.closest('tr');
  const rows = $$('tr', tbl);
  const colIdx = Array.from(tr.children).indexOf(activeTd);

  if (action === 'addRow') {
    const newTr = h('tr');
    const colCount = tr.children.length || 1;
    for (let i = 0; i < colCount; i++) newTr.append(h('td', null, h('div', { class: 'dp' }, '')));
    insertAfter(tr, newTr);
  } else if (action === 'delRow' && rows.length > 1) {
    const fallbackRow = tr.nextElementSibling || tr.previousElementSibling;
    tr.remove();
    const fallbackTd = fallbackRow ? (fallbackRow.children[Math.min(colIdx, fallbackRow.children.length - 1)] || fallbackRow.firstElementChild) : null;
    $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
    curDocx.activeTd = fallbackTd || null;
    if (curDocx.activeTd) curDocx.activeTd.classList.add('doc-td-active');
  } else if (action === 'addCol') {
    for (const r of rows) {
      const refTd = r.children[colIdx] || r.lastElementChild;
      const newTd = h('td', null, h('div', { class: 'dp' }, ''));
      if (refTd) insertAfter(refTd, newTd);
      else r.append(newTd);
    }
  } else if (action === 'delCol' && tr.children.length > 1) {
    for (const r of rows) {
      if (r.children[colIdx]) r.children[colIdx].remove();
    }
    const fallbackTd = tr.children[Math.min(colIdx, tr.children.length - 1)] || tr.firstElementChild || null;
    $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
    curDocx.activeTd = fallbackTd;
    if (curDocx.activeTd) curDocx.activeTd.classList.add('doc-td-active');
  }
  syncPaperToDocxXml(paper);
  pushHistory('docTableEdit');
}

function normalizePaperBlocks(paper) {
  if (!paper) return;
  for (const ch of Array.from(paper.childNodes)) {
    if (ch.nodeType === 3 && ch.nodeValue && ch.nodeValue.trim()) {
      const dp = h('div', { class: 'dp' }, ch.nodeValue);
      paper.replaceChild(dp, ch);
    } else if (ch.nodeType === 1 && ch.tagName === 'SPAN' && ch.querySelector('.dp')) {
      ch.replaceWith(...Array.from(ch.childNodes));
    }
  }
}

function execDocCmd(cmd, val) {
  const paper = $('#app .paper');
  if (paper) ensureActiveDocxSelection(paper);
  document.execCommand('styleWithCSS', false, true);
  document.execCommand(cmd, false, val || null);
  if (paper) {
    const sel = window.getSelection();
    if (sel && sel.rangeCount && paper.contains(sel.anchorNode)) {
      savedDocxSelectionRange = sel.getRangeAt(0).cloneRange();
    }
    normalizePaperBlocks(paper);
    syncPaperToDocxXml(paper);
    pushHistory('fmt:' + cmd);
  }
}

function buildDocxDrawingRun(doc, rId, idNum, cxEmu, cyEmu) {
  return frag(doc, `<w:r xmlns:w="${W}" xmlns:wp="${NS.wp}" xmlns:a="${A}" xmlns:pic="${NS.pic}" xmlns:r="${R}">
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${cxEmu}" cy="${cyEmu}"/>
        <wp:docPr id="${idNum}" name="Picture ${idNum}"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr><pic:cNvPr id="${idNum}" name="Picture ${idNum}"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>`);
}

function syncPaperToDocxXml(paper) {
  normalizePaperBlocks(paper);
  const doc = pk.doc(pk.mainPart);
  const body = kid(doc.documentElement, W, 'body');
  if (!body) return;

  const existingRels = readRels(pk, pk.mainPart);
  const pathToRid = new Map();
  for (const r of existingRels.values()) {
    if (r.path && r.type.endsWith('/image')) pathToRid.set(r.path, r.id);
  }
  let picCounter = 1;
  let hasLists = false;

  const buildParagraphXmlFromBlock = (domEl, listNumId) => {
    const wP = mk(doc, W, 'w:p');
    const pPr = mk(doc, W, 'w:pPr');
    let hasPPr = false;

    if (domEl.dataset?.heading) {
      const hMap = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3' };
      if (hMap[domEl.dataset.heading]) {
        pPr.appendChild(mk(doc, W, 'w:pStyle', { 'w:val': hMap[domEl.dataset.heading] }));
        hasPPr = true;
      }
    }

    if (domEl.classList?.contains('doc-callout')) {
      pPr.appendChild(frag(doc, `<w:pBdr xmlns:w="${W}"><w:left w:val="single" w:sz="24" w:space="8" w:color="B4531A"/><w:right w:val="single" w:sz="24" w:space="8" w:color="B4531A"/></w:pBdr>`));
      pPr.appendChild(mk(doc, W, 'w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'FFF9F1' }));
      hasPPr = true;
    }

    if (listNumId) {
      hasLists = true;
      const numPr = mk(doc, W, 'w:numPr');
      numPr.appendChild(mk(doc, W, 'w:ilvl', { 'w:val': '0' }));
      numPr.appendChild(mk(doc, W, 'w:numId', { 'w:val': listNumId }));
      pPr.appendChild(numPr);
      hasPPr = true;
    }

    const stBlock = domEl.style || {};
    if (stBlock.direction === 'rtl') {
      pPr.appendChild(mk(doc, W, 'w:bidi'));
      hasPPr = true;
    } else if (stBlock.direction === 'ltr') {
      pPr.appendChild(mk(doc, W, 'w:bidi', { 'w:val': '0' }));
      hasPPr = true;
    }
    if (stBlock.textAlign) {
      const jcMap = { left: 'left', center: 'center', right: 'right', justify: 'both' };
      if (jcMap[stBlock.textAlign]) {
        pPr.appendChild(mk(doc, W, 'w:jc', { 'w:val': jcMap[stBlock.textAlign] }));
        hasPPr = true;
      }
    }
    if (stBlock.lineHeight) {
      const lh = parseFloat(stBlock.lineHeight);
      if (!isNaN(lh) && lh > 0) {
        pPr.appendChild(mk(doc, W, 'w:spacing', { 'w:line': String(Math.round(lh * 240)), 'w:lineRule': 'auto' }));
        hasPPr = true;
      }
    }
    if (stBlock.paddingInlineStart) {
      const pt = parseFloat(stBlock.paddingInlineStart);
      if (!isNaN(pt) && pt > 0) {
        pPr.appendChild(mk(doc, W, 'w:ind', { 'w:start': String(Math.round(pt * 20)) }));
        hasPPr = true;
      }
    }
    if (hasPPr) wP.appendChild(pPr);

    const walk = (node, fmt, targetContainer) => {
      if (node.nodeType === 3) {
        const txt = node.nodeValue;
        if (!txt) return;
        const wR = mk(doc, W, 'w:r');
        if (fmt.link || fmt.b || fmt.i || fmt.u || fmt.strike || fmt.vAlign || fmt.hex || fmt.bgHex || fmt.font || fmt.szPt) {
          const rPr = mk(doc, W, 'w:rPr');
          if (fmt.link) rPr.appendChild(mk(doc, W, 'w:rStyle', { 'w:val': 'Hyperlink' }));
          if (fmt.font) rPr.appendChild(mk(doc, W, 'w:rFonts', { 'w:ascii': fmt.font, 'w:hAnsi': fmt.font, 'w:cs': fmt.font }));
          if (fmt.b) rPr.appendChild(mk(doc, W, 'w:b'));
          if (fmt.i) rPr.appendChild(mk(doc, W, 'w:i'));
          if (fmt.u) rPr.appendChild(mk(doc, W, 'w:u', { 'w:val': 'single' }));
          if (fmt.strike) rPr.appendChild(mk(doc, W, 'w:strike'));
          if (fmt.vAlign) rPr.appendChild(mk(doc, W, 'w:vertAlign', { 'w:val': fmt.vAlign }));
          if (fmt.hex) rPr.appendChild(mk(doc, W, 'w:color', { 'w:val': fmt.hex }));
          if (fmt.bgHex) rPr.appendChild(mk(doc, W, 'w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': fmt.bgHex }));
          if (fmt.szPt) {
            const halfPt = String(Math.round(fmt.szPt * 2));
            rPr.appendChild(mk(doc, W, 'w:sz', { 'w:val': halfPt }));
            rPr.appendChild(mk(doc, W, 'w:szCs', { 'w:val': halfPt }));
          }
          wR.appendChild(rPr);
        }
        const wT = mk(doc, W, 'w:t');
        if (/^\s|\s$/.test(txt)) wT.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        wT.textContent = txt;
        wR.appendChild(wT);
        targetContainer.appendChild(wR);
      } else if (node.nodeType === 1) {
        if (node.hasAttribute('data-docx-bookmark-start')) {
          targetContainer.appendChild(mk(doc, W, 'w:bookmarkStart', { 'w:id': node.getAttribute('data-docx-bookmark-start'), 'w:name': node.getAttribute('data-docx-bookmark-name') || '' }));
          return;
        }
        if (node.hasAttribute('data-docx-bookmark-end')) {
          targetContainer.appendChild(mk(doc, W, 'w:bookmarkEnd', { 'w:id': node.getAttribute('data-docx-bookmark-end') }));
          return;
        }
        if (node.classList?.contains('paper-watermark') || node.classList?.contains('paper-hf-preview') || node.classList?.contains('doc-toc') || node.classList?.contains('docx-img-resizer')) return;
        if (node.classList?.contains('doc-page-break')) {
          const wR = mk(doc, W, 'w:r');
          wR.appendChild(mk(doc, W, 'w:br', { 'w:type': 'page' }));
          targetContainer.appendChild(wR);
          return;
        }
        if (node.tagName === 'IMG') {
          const mPath = node.getAttribute('data-media-path');
          if (mPath && pk.has(mPath)) {
            let rId = pathToRid.get(mPath);
            if (!rId) {
              rId = addRel(pk, pk.mainPart, REL + 'image', mPath);
              pathToRid.set(mPath, rId);
            }
            const rect = node.getBoundingClientRect();
            const nw = rect.width || node.naturalWidth || node.width || 300;
            const nh = rect.height || node.naturalHeight || node.height || 180;
            const cx = Math.min(5486400, Math.round(nw * 9525));
            const cy = Math.round(cx * (nh / (nw || 1)));
            targetContainer.appendChild(buildDocxDrawingRun(doc, rId, picCounter++, cx, cy));
          }
          return;
        }
        if (node.tagName === 'BR') {
          const wR = mk(doc, W, 'w:r');
          wR.appendChild(mk(doc, W, 'w:br'));
          targetContainer.appendChild(wR);
          return;
        }
        if (node.tagName === 'A' && node.getAttribute('href')) {
          const rawHref = node.getAttribute('href').trim();
          const hl = mk(doc, W, 'w:hyperlink');
          if (rawHref.startsWith('#')) hl.setAttributeNS(W, 'w:anchor', rawHref.slice(1));
          else {
            const href = safeUrl(rawHref);
            if (href === '#') { for (const ch of Array.from(node.childNodes)) walk(ch, fmt, targetContainer); return; }
            const rId = addRel(pk, pk.mainPart, REL + 'hyperlink', href, true);
            hl.setAttributeNS(R, 'r:id', rId);
          }
          for (const ch of Array.from(node.childNodes)) walk(ch, { ...fmt, link: true, u: true, hex: fmt.hex || '0563C1' }, hl);
          targetContainer.appendChild(hl);
          return;
        }

        const next = { ...fmt };
        const st = node.style || {};
        if (node.tagName === 'B' || node.tagName === 'STRONG' || st.fontWeight === '700' || st.fontWeight === 'bold') next.b = true;
        if (node.tagName === 'I' || node.tagName === 'EM' || st.fontStyle === 'italic') next.i = true;
        if (node.tagName === 'U' || (st.textDecoration && st.textDecoration.includes('underline'))) next.u = true;
        if (node.tagName === 'S' || node.tagName === 'STRIKE' || (st.textDecoration && st.textDecoration.includes('line-through'))) next.strike = true;
        if (node.tagName === 'SUP' || st.verticalAlign === 'super') next.vAlign = 'superscript';
        if (node.tagName === 'SUB' || st.verticalAlign === 'sub') next.vAlign = 'subscript';

        const hex = parseCssColorToHex(st.color) || (node.tagName === 'FONT' ? parseCssColorToHex(node.getAttribute('color')) : null);
        if (hex) next.hex = hex;
        const bgHex = parseCssColorToHex(st.backgroundColor);
        if (bgHex && bgHex !== 'FFFFFF') next.bgHex = bgHex;

        if (st.fontFamily) next.font = st.fontFamily.replace(/["']/g, '').split(',')[0].trim();
        else if (node.tagName === 'FONT' && node.getAttribute('face')) next.font = node.getAttribute('face');

        if (st.fontSize && st.fontSize.endsWith('pt')) {
          const pt = parseFloat(st.fontSize);
          if (!isNaN(pt)) next.szPt = pt;
        }

        for (const ch of Array.from(node.childNodes)) walk(ch, next, targetContainer);
      }
    };

    walk(domEl, {}, wP);
    return wP;
  };

  const sectPr = kid(body, W, 'sectPr');
  for (const child of Array.from(kids(body, W))) if (child.localName !== 'sectPr') removeNode(child);

  for (const block of kids(paper)) {
    if (block.classList?.contains('paper-watermark') || block.classList?.contains('paper-hf-preview') || block.classList?.contains('docx-img-resizer')) continue;
    if (block.classList?.contains('doc-page-break')) {
      const brP = mk(doc, W, 'w:p');
      const brR = mk(doc, W, 'w:r');
      brR.appendChild(mk(doc, W, 'w:br', { 'w:type': 'page' }));
      brP.appendChild(brR);
      body.insertBefore(brP, sectPr || null);
      continue;
    }
    if (block.tagName === 'HR') {
      const hrP = frag(doc, `<w:p xmlns:w="${W}"><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="CBD5E1"/></w:pBdr></w:pPr></w:p>`);
      body.insertBefore(hrP, sectPr || null);
      continue;
    }
    if (block.classList?.contains('doc-toc')) {
      for (const item of kids(block)) {
        body.insertBefore(buildParagraphXmlFromBlock(item, null), sectPr || null);
      }
      continue;
    }
    if (block.classList && block.classList.contains('tw')) {
      const htmlTbl = $('table.dt', block);
      const hasGrid = htmlTbl && htmlTbl.classList.contains('grid');
      const tbl = mk(doc, W, 'w:tbl');
      const tblPr = mk(doc, W, 'w:tblPr');
      const bVal = hasGrid ? 'single' : 'none';
      tblPr.appendChild(frag(doc, `<w:tblBorders xmlns:w="${W}">
        <w:top w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
        <w:left w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
        <w:bottom w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
        <w:right w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
        <w:insideH w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
        <w:insideV w:val="${bVal}" w:sz="4" w:space="0" w:color="666666"/>
      </w:tblBorders>`));
      tbl.appendChild(tblPr);

      for (const trRow of $$('tr', block)) {
        const wTr = mk(doc, W, 'w:tr');
        for (const td of $$('td', trRow)) {
          const wTc = mk(doc, W, 'w:tc');
          const tcPr = mk(doc, W, 'w:tcPr');
          let hasTcPr = false;
          if (td.colSpan > 1) {
            tcPr.appendChild(mk(doc, W, 'w:gridSpan', { 'w:val': String(td.colSpan) }));
            hasTcPr = true;
          }
          if (td.style?.verticalAlign === 'middle') {
            tcPr.appendChild(mk(doc, W, 'w:vAlign', { 'w:val': 'center' }));
            hasTcPr = true;
          }
          const cellBg = parseCssColorToHex(td.style?.backgroundColor);
          if (cellBg) {
            tcPr.appendChild(mk(doc, W, 'w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': cellBg }));
            hasTcPr = true;
          }
          if (hasTcPr) wTc.appendChild(tcPr);

          const subPs = kids(td);
          if (subPs.length) {
            for (const sp of subPs) wTc.appendChild(buildParagraphXmlFromBlock(sp, null));
          } else {
            wTc.appendChild(buildParagraphXmlFromBlock(td, null));
          }
          wTr.appendChild(wTc);
        }
        tbl.appendChild(wTr);
      }
      body.insertBefore(tbl, sectPr || null);
    } else if (block.tagName === 'UL' || block.tagName === 'OL') {
      const numId = block.tagName === 'OL' ? '2' : '1';
      for (const li of kids(block)) {
        body.insertBefore(buildParagraphXmlFromBlock(li, numId), sectPr || null);
      }
    } else {
      body.insertBefore(buildParagraphXmlFromBlock(block, null), sectPr || null);
    }
  }

  if (hasLists) ensureDocxNumberingPart();
}

let savedDocxSelectionRange = null;
function saveDocxSelection(paper) {
  const sel = window.getSelection();
  savedDocxSelectionRange = sel && sel.rangeCount && paper.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : null;
}
function restoreDocxSelection(paper) {
  if (!savedDocxSelectionRange) return null;
  const sel = window.getSelection();
  try { paper.focus(); sel.removeAllRanges(); sel.addRange(savedDocxSelectionRange); return savedDocxSelectionRange; }
  catch (_) { return null; }
}
function insertAtDocxSelection(paper, node) {
  const range = restoreDocxSelection(paper);
  if (range && paper.contains(range.commonAncestorContainer)) {
    range.deleteContents(); range.insertNode(node);
    range.setStartAfter(node); range.collapse(true);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
  } else paper.append(node);
  savedDocxSelectionRange = null;
}

function insertDocxImage(paper, imgInfo) {
  const mediaPath = uniquePath(pk, 'word/media', 'image', imgInfo.ext);
  pk.setBin(mediaPath, imgInfo.bytes);
  ensureDefault(pk, imgInfo.ext, CT_IMG[imgInfo.ext] || 'image/png');
  addRel(pk, pk.mainPart, REL + 'image', mediaPath);
  const url = mediaURL(imgInfo.bytes, mediaPath);
  const imageBlock = h('div', { class: 'dp' }, h('img', {
    src: url, alt: '', 'data-media-path': mediaPath, style: { width: '340px' }
  }));
  insertAtDocxSelection(paper, imageBlock);
  syncPaperToDocxXml(paper);
  pushHistory('addDocImg');
  toast('تمت إضافة الصورة');
}

function insertDocxTable(paper, rows, cols) {
  const tw = h('div', { class: 'tw' });
  const tbl = h('table', { class: 'dt grid' });
  for (let r = 0; r < rows; r++) {
    const trRow = h('tr');
    for (let c = 0; c < cols; c++) trRow.append(h('td', null, h('div', { class: 'dp' }, '')));
    tbl.append(trRow);
  }
  tw.append(tbl);
  insertAtDocxSelection(paper, tw);
  const firstTd = $('td', tbl);
  if (firstTd) {
    $$('.docx-img-sel', paper).forEach(x => x.classList.remove('docx-img-sel'));
    $$('.doc-td-active', paper).forEach(x => x.classList.remove('doc-td-active'));
    curDocx.activeImg = null;
    curDocx.activeTd = firstTd;
    firstTd.classList.add('doc-td-active');
    const firstDp = $('.dp', firstTd) || firstTd;
    const rng = document.createRange();
    rng.selectNodeContents(firstDp);
    rng.collapse(true);
    const sel = window.getSelection();
    if (sel) { sel.removeAllRanges(); sel.addRange(rng); }
  }
  syncPaperToDocxXml(paper);
  pushHistory('addDocTbl');
  toast('تمت إضافة جدول');
}

async function buildDocxPdfBytes(paper) {
  let holder = null;
  try {
    const isLand = curDocx.orientation === 'landscape';
    const isLetter = curDocx.pageSize === 'letter';
    const baseWPt = isLetter ? 612 : 595.28;
    const baseHPt = isLetter ? 792 : 841.89;
    const pageWPt = isLand ? baseHPt : baseWPt;
    const pageHPt = isLand ? baseWPt : baseHPt;
    const pxPerPt = 96 / 72;
    const pageWPx = Math.round(pageWPt * pxPerPt);
    const pageHPx = Math.round(pageHPt * pxPerPt);
    const marginTopPx = curDocx.margins === 'narrow' ? 48 : curDocx.margins === 'wide' ? 88 : 72;
    const marginBotPx = marginTopPx;
    const marginSidePx = curDocx.margins === 'narrow' ? 48 : curDocx.margins === 'wide' ? 136 : 75;
    const usableHPx = Math.max(200, pageHPx - marginTopPx - marginBotPx);

    const clone = paper.cloneNode(true);
    clone.removeAttribute('contenteditable');
    clone.style.transform = '';
    clone.style.zoom = '';
    clone.style.margin = '0';
    clone.style.width = pageWPx + 'px';
    clone.style.maxWidth = 'none';
    clone.style.minHeight = '0';
    clone.style.padding = `0 ${marginSidePx}px`;
    $$('.docx-img-sel, .doc-td-active, .docx-img-resizer, .paper-hf-preview, .paper-watermark', clone).forEach(x => {
      if (x.classList.contains('docx-img-resizer') || x.classList.contains('paper-hf-preview') || x.classList.contains('paper-watermark')) x.remove();
      else x.classList.remove('docx-img-sel', 'doc-td-active');
    });
    $$('.doc-page-break span', clone).forEach(s => s.remove());
    await inlineBlobAssets(clone);

    const hf = getDocxHeaderFooterText();
    const sheetCss = `
      *{box-sizing:border-box}
      .pdf-doc-root{background:${curDocx.pageColor || '#FFFFFF'};color:#000;font-family:"Cairo","Segoe UI",Tahoma,Arial,sans-serif;width:${pageWPx}px}
      .paper{position:relative;background:${curDocx.pageColor || '#FFFFFF'};width:${pageWPx}px;padding:0 ${marginSidePx}px;direction:${clone.style.direction || 'rtl'};font-size:12pt;line-height:1.65;color:#000;overflow-wrap:anywhere;border:none;box-shadow:none}
      .paper.cols-2{column-count:2;column-gap:24px;column-rule:1px solid #e0d9cb}
      .dp{position:relative;white-space:pre-wrap;min-height:1.35em;margin:0 0 8px}
      .dp img{max-width:100%;height:auto}
      .dp a{color:#2563eb;text-decoration:underline}
      .tw{margin:14px 0}
      table.dt{border-collapse:collapse;width:100%;table-layout:fixed}
      table.dt td{vertical-align:top;padding:5px 8px;overflow-wrap:anywhere}
      table.dt.grid td{border:1px solid #666}
      ul,ol{margin:6px 0 10px;padding-inline-start:26px}
      li{margin-bottom:4px}
      .doc-callout{background:#fff9f1;border-inline-start:4px solid #b4531a;padding:10px 14px;border-radius:6px;margin:10px 0;color:#1f1c17}
      .doc-hr{border:0;border-top:2px solid #cbd5e1;margin:16px 0}
      .doc-toc{background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:12px 16px;margin:12px 0 18px;color:#0f172a}
      .doc-toc-title{font-weight:800;font-size:13pt;margin-bottom:6px;border-bottom:1px solid #e2e8f0;padding-bottom:4px}
      .doc-toc-item{font-size:10.5pt;margin:3px 0;color:#334155}
      .doc-fn-ref{color:#b4531a;font-weight:800;padding:0 2px}
      .doc-footnotes-box{margin-top:34px;padding-top:12px;border-top:1.5px solid #94a3b8;font-size:9.5pt;color:#334155;clear:both}
      .doc-comment-mark{background:#fef08a;border-bottom:2px solid #eab308;padding:0 2px}
      .doc-page-break{height:1px;border:0;margin:0;padding:0;opacity:0}
    `;

    const wrap = h('div', { class: 'pdf-doc-root' });
    wrap.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
    const styleNode = document.createElement('style');
    styleNode.textContent = sheetCss;
    wrap.append(styleNode, clone);

    holder = h('div', { style: { position: 'fixed', left: '-9999px', top: '0', width: pageWPx + 'px' } }, wrap);
    document.body.appendChild(holder);

    // Block-aware pagination: split cleanly between paragraphs/tables & honor .doc-page-break
    const totalH = Math.max(1, Math.ceil(wrap.scrollHeight));
    const slices = []; // [{ yStart, yEnd }]
    if (!curDocx.orientation && clone.classList.contains('cols-2')) {
      for (let y = 0; y < totalH; y += usableHPx) slices.push({ yStart: y, yEnd: Math.min(totalH, y + usableHPx) });
    } else {
      const wrapRect = wrap.getBoundingClientRect();
      const children = kids(clone);
      let curStart = 0;
      for (const ch of children) {
        const r = ch.getBoundingClientRect();
        const topY = Math.max(0, Math.floor(r.top - wrapRect.top));
        const botY = Math.max(topY, Math.ceil(r.bottom - wrapRect.top));
        if (ch.classList?.contains('doc-page-break')) {
          slices.push({ yStart: curStart, yEnd: Math.max(curStart + 1, topY) });
          curStart = botY;
          continue;
        }
        if (botY - curStart > usableHPx) {
          if (topY > curStart + 24) {
            slices.push({ yStart: curStart, yEnd: topY });
            curStart = topY;
          }
          while (botY - curStart > usableHPx) {
            slices.push({ yStart: curStart, yEnd: curStart + usableHPx });
            curStart += usableHPx;
          }
        }
      }
      if (curStart < totalH || !slices.length) {
        slices.push({ yStart: curStart, yEnd: Math.max(curStart + 1, totalH) });
      }
    }

    const xml = new XMLSerializer().serializeToString(wrap);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${pageWPx}" height="${Math.max(totalH, 10)}"><foreignObject width="100%" height="100%">${xml}</foreignObject></svg>`;
    const svgUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    const img = new Image();
    img.src = svgUrl;
    await img.decode();

    const { PDFDocument } = getPdfLib();
    const pdfDoc = await PDFDocument.create();
    const scaleRatio = 1.75;
    const pageCount = slices.length;
    const wmText = (curDocx.watermark || '').trim();

    for (let p = 0; p < pageCount; p++) {
      const sl = slices[p];
      const sliceH = Math.max(1, Math.min(usableHPx, sl.yEnd - sl.yStart));
      const c = document.createElement('canvas');
      c.width = Math.round(pageWPx * scaleRatio);
      c.height = Math.round(pageHPx * scaleRatio);
      const ctx = c.getContext('2d');
      ctx.scale(scaleRatio, scaleRatio);
      ctx.fillStyle = curDocx.pageColor || '#FFFFFF';
      ctx.fillRect(0, 0, pageWPx, pageHPx);

      if (curDocx.hasBorder) {
        ctx.strokeStyle = '#1f1c17';
        ctx.lineWidth = 2.5;
        ctx.strokeRect(28, 28, pageWPx - 56, pageHPx - 56);
      }

      if (sl.yEnd > sl.yStart) {
        ctx.drawImage(img, 0, sl.yStart, pageWPx, sliceH, 0, marginTopPx, pageWPx, sliceH);
      }

      if (wmText) {
        ctx.save();
        ctx.translate(pageWPx / 2, pageHPx * 0.48);
        ctx.rotate((-32 * Math.PI) / 180);
        ctx.font = '800 72px "Cairo","Segoe UI",sans-serif';
        ctx.fillStyle = 'rgba(150,150,150,0.16)';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(wmText, 0, 0);
        ctx.restore();
      }

      if (hf.headerTxt && hf.headerTxt.trim()) {
        ctx.fillStyle = '#64748b';
        ctx.font = '12px "Cairo","Segoe UI",sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(hf.headerTxt.trim(), pageWPx / 2, 28);
      }
      if ((hf.footerTxt && hf.footerTxt.trim()) || hf.hasPageNum) {
        const fLabel = [hf.footerTxt?.trim(), hf.hasPageNum ? `${p + 1} / ${pageCount}` : ''].filter(Boolean).join('  |  ');
        ctx.fillStyle = '#64748b';
        ctx.font = '12px "Cairo","Segoe UI",sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(fLabel, pageWPx / 2, pageHPx - 20);
      }

      const pngBytes = await new Promise((resolve, reject) => {
        c.toBlob(async (b) => {
          if (!b) return reject(new Error('PNG slice failed'));
          resolve(new Uint8Array(await b.arrayBuffer()));
        }, 'image/png');
      });
      const pngImg = await pdfDoc.embedPng(pngBytes);
      const pdfPage = pdfDoc.addPage([pageWPt, pageHPt]);
      pdfPage.drawImage(pngImg, { x: 0, y: 0, width: pageWPt, height: pageHPt });
    }

    return await pdfDoc.save();
  } finally {
    if (holder) holder.remove();
  }
}

async function exportDocxToPdf(paper) {
  busy(true, 'جارٍ تجهيز ملف PDF…');
  try {
    const pdfBytes = await buildDocxPdfBytes(paper);
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const base = (fileName || 'document').replace(/\.(docx|docm)$/i, '') || 'document';
    a.href = url;
    a.download = `${base}.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast('تم تصدير ملف PDF بنجاح');
  } catch (err) {
    console.error('DOCX PDF export failed:', err);
    toast('حدث خطأ أثناء تصدير الملف');
  } finally {
    busy(false);
  }
}

async function convertDocxToPdfEditor(paper) {
  busy(true, 'جارٍ تجهيز ملف PDF…');
  try {
    const pdfBytes = await buildDocxPdfBytes(paper);
    const base = (fileName || 'document').replace(/\.(docx|docm)$/i, '') || 'document';
    const pdfFile = new File([pdfBytes], `${base}.pdf`, { type: 'application/pdf' });
    await handleFile(pdfFile);
    toast(tr('تم تحويل المستند إلى محرر PDF'));
  } catch (err) {
    console.error('DOCX to PDF editor conversion failed:', err);
    toast('حدث خطأ أثناء تصدير الملف');
  } finally {
    busy(false);
  }
}
