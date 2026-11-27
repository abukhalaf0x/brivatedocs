'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 5 من 6: التراجع، المسودات، شريط PPTX، ولوحة الخصائص
   ═════════════════════════════════════════════════════════════ */

let hist = [];
let histIdx = -1;
let histDebounceTimer = null;

function defaultDocxState() {
  return {
    activeTd: null,
    activeImg: null,
    zoom: '100',
    pageSize: 'a4',
    orientation: 'portrait',
    margins: 'normal',
    showHf: false,
    pageColor: '#FFFFFF',
    hasBorder: false,
    watermark: '',
    copiedStyle: null
  };
}

let curPptx = {
  deck: null,
  slideIdx: 0,
  selShapeIdx: null,
  multiSel: new Set(),
  selCell: null,
  activeTab: 'home',
  clipboardXml: null,
  clipboardRels: [],
  clipboardSourcePath: null,
  copiedStyle: null,
  viewMode: 'normal',
  zoom: '100',
  showGrid: false
};
let curDocx = defaultDocxState();

function showView(id) {
  $('#home').hidden = id !== 'home';
  $('#editor').hidden = id !== 'editor';
  if (id === 'home') checkSavedDraftBanner();
}

function showHomeError(msg) {
  const el = $('#homeError');
  if (!msg) { el.hidden = true; el.textContent = ''; return; }
  el.hidden = false;
  el.textContent = tr(msg);
}

function resetHistory() {
  clearTimeout(histDebounceTimer);
  hist = [];
  histIdx = -1;
  syncHistoryButtons();
}

let hasPendingHistory = false;

function pushHistory(label) {
  if (!pk) return;
  clearTimeout(histDebounceTimer);
  histDebounceTimer = null;
  hasPendingHistory = false;
  const snap = pk.snapshot();
  hist = hist.slice(0, histIdx + 1);
  hist.push({ snap, label: label || '' });
  if (hist.length > 60) hist.shift();
  histIdx = hist.length - 1;
  syncHistoryButtons();

  DraftDB.save({
    files: Array.from(snap.entries()),
    fileName,
    kind,
    mainPart: pk.mainPart,
    updatedAt: Date.now()
  });
}

function scheduleHistory(label, delay = 350) {
  clearTimeout(histDebounceTimer);
  hasPendingHistory = true;
  syncHistoryButtons();
  histDebounceTimer = setTimeout(() => pushHistory(label), delay);
}

function syncHistoryButtons() {
  if (!pk) hasPendingHistory = false;
  const u = $('#bUndo'), r = $('#bRedo');
  if (u) u.disabled = histIdx <= 0 && !hasPendingHistory;
  if (r) r.disabled = hasPendingHistory || histIdx >= hist.length - 1;
}

function applySnapshot(idx) {
  if (hasPendingHistory && pk) {
    const isUndo = idx < histIdx;
    pushHistory('pendingEdit');
    if (isUndo) idx = histIdx - 1;
  }
  if (idx < 0 || idx >= hist.length || !pk) return;
  clearTimeout(histDebounceTimer);
  histDebounceTimer = null;
  hasPendingHistory = false;
  histIdx = idx;
  pk.restore(hist[histIdx].snap);
  syncHistoryButtons();
  renderEditor();
  DraftDB.save({
    files: Array.from(pk.snapshot().entries()),
    fileName,
    kind,
    mainPart: pk.mainPart,
    updatedAt: Date.now()
  });
}

/* ── فحص واستعادة المشاريع الأخيرة (حتى 5 مسودات) من IndexedDB ── */
async function checkSavedDraftBanner() {
  const banner = $('#draftBanner');
  const listEl = $('#recentDraftsList');
  if (!banner || !listEl) return;
  const drafts = await DraftDB.loadAll();
  if (!drafts.length) {
    banner.hidden = true;
    return;
  }
  listEl.innerHTML = '';
  for (const d of drafts) {
    const timeStr = d.updatedAt ? new Date(d.updatedAt).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
    const icon = d.kind === 'pptx' ? '📊' : '📄';
    const row = h('div', { class: 'recent-draft-item' },
      h('div', { class: 'recent-draft-meta' },
        h('span', null, icon),
        h('strong', null, d.fileName || 'document'),
        h('small', { class: 'muted' }, timeStr)
      ),
      h('div', { class: 'recent-draft-actions' },
        btn('↩ استعادة المسودة', () => restoreSavedDraft(d), 'pri'),
        btn('🗑', async () => {
          await DraftDB.removeById(d.id);
          checkSavedDraftBanner();
        }, 'danger')
      )
    );
    listEl.append(row);
  }
  banner.hidden = false;
}

async function restoreSavedDraft(targetDraft) {
  busy(true, 'جارٍ فتح الملف…');
  try {
    const draft = targetDraft || (await DraftDB.loadAll())[0];
    if (!draft || !Array.isArray(draft.files)) return;
    activeDraftId = draft.id || ('draft_' + Date.now());
    pk = new Pkg();
    pk.files = new Map(draft.files);
    pk.mainPart = draft.mainPart || detectMainPart();
    kind = draft.kind || detectMainPart();
    fileName = draft.fileName || (kind === 'pptx' ? 'presentation.pptx' : 'document.docx');
    $('#fname').textContent = fileName;
    curPptx = { deck: null, slideIdx: 0, selShapeIdx: null, multiSel: new Set(), selCell: null, activeTab: 'home', clipboardXml: null, clipboardRels: [], clipboardSourcePath: null, copiedStyle: null, viewMode: 'normal', zoom: '100', showGrid: false };
    curDocx = defaultDocxState();
    resetHistory();
    pushHistory('restored');
    showView('editor');
    renderEditor();
    applyLanguage(currentLanguage);
    toast('تمت استعادة المسودة بنجاح');
  } catch (err) {
    showHomeError(err instanceof UserError ? err.message : ('حدث خطأ أثناء قراءة الملف: ' + (err && err.message || err)));
  } finally {
    busy(false);
  }
}

function detectMainPart() {
  if (pk.has('ppt/presentation.xml')) { pk.mainPart = 'ppt/presentation.xml'; return 'pptx'; }
  if (pk.has('word/document.xml')) { pk.mainPart = 'word/document.xml'; return 'docx'; }
  const ct = pk.doc('[Content_Types].xml').documentElement;
  for (const ov of kids(ct, null, 'Override')) {
    const t = ov.getAttribute('ContentType') || '';
    const pn = (ov.getAttribute('PartName') || '').replace(/^\//, '');
    if (t.includes('presentationml.presentation.main')) { pk.mainPart = pn; return 'pptx'; }
    if (t.includes('wordprocessingml.document.main')) { pk.mainPart = pn; return 'docx'; }
  }
  throw new UserError('تعذر التعرف على نوع الملف — يرجى التأكد من أنه مستند Word (.docx) أو عرض PowerPoint (.pptx) صحيح.');
}

async function handleFile(file) {
  if (!file) return;
  showHomeError('');
  activeDraftId = null;
  busy(true, 'جارٍ فتح الملف…');
  try {
    assertFileSize(file); const buf = await file.arrayBuffer();
    pk = await Pkg.load(buf);
    kind = detectMainPart();
    fileName = file.name || (kind === 'pptx' ? 'presentation.pptx' : 'document.docx');
    $('#fname').textContent = fileName;
    curPptx = { deck: null, slideIdx: 0, selShapeIdx: null, multiSel: new Set(), selCell: null, activeTab: 'home', clipboardXml: null, clipboardRels: [], clipboardSourcePath: null, copiedStyle: null, viewMode: 'normal', zoom: '100', showGrid: false };
    curDocx = defaultDocxState();
    resetHistory();
    pushHistory('initial');
    showView('editor');
    renderEditor();
    applyLanguage(currentLanguage);
    toast('تم فتح الملف بنجاح');
  } catch (err) {
    showView('home');
    showHomeError(err instanceof UserError ? err.message : ('حدث خطأ أثناء قراءة الملف: ' + (err && err.message || err)));
  } finally {
    busy(false);
  }
}

function startNewDeck(themeKey) {
  showHomeError('');
  activeDraftId = null;
  pk = createNewBlankDeck(themeKey);
  kind = 'pptx';
  fileName = 'presentation.pptx';
  $('#fname').textContent = fileName;
  curPptx = { deck: null, slideIdx: 0, selShapeIdx: null, multiSel: new Set(), selCell: null, activeTab: 'home', clipboardXml: null, clipboardRels: [], clipboardSourcePath: null, copiedStyle: null, viewMode: 'normal', zoom: '100', showGrid: false };
  resetHistory();
  pushHistory('newDeck');
  showView('editor');
  renderEditor();
  applyLanguage(currentLanguage);
  toast('تم إنشاء العرض التقديمي بنجاح');
}

function startNewDocx(templateKey = 'blank') {
  showHomeError('');
  activeDraftId = null;
  pk = createNewBlankDocx(templateKey);
  kind = 'docx';
  fileName = 'document.docx';
  $('#fname').textContent = fileName;
  curDocx = defaultDocxState();
  resetHistory();
  pushHistory('newDocx');
  showView('editor');
  renderEditor();
  applyLanguage(currentLanguage);
  toast('تم إنشاء مستند Word بنجاح');
}

/* ── تغيير مقاس الشريحة (16:9 / 4:3) ── */
function setDeckAspectRatio(deck, ratio) {
  const root = pk.doc(deck.pres).documentElement;
  let sz = kid(root, P, 'sldSz');
  if (!sz) {
    sz = mk(pk.doc(deck.pres), P, 'p:sldSz');
    root.appendChild(sz);
  }
  const newW = ratio === '4:3' ? 9144000 : 12192000;
  const newH = 6858000;
  const oldW = deck.W || newW;
  const scaleX = newW / oldW;

  sz.setAttribute('cx', String(newW));
  sz.setAttribute('cy', String(newH));
  deck.W = newW; deck.H = newH;
  if (ratio === '4:3') sz.setAttribute('type', 'screen4x3');
  else sz.setAttribute('type', 'screen16x9');

  if (Math.abs(scaleX - 1) > 0.001) {
    for (const sl of deck.slides) {
      const doc = pk.doc(sl.path);
      for (const sh of shapesOf(doc)) {
        const g = readXfrm(xfrmOf(sh));
        if (!g) continue;
        g.x = Math.round(g.x * scaleX);
        g.w = Math.round(g.w * scaleX);
        writeGeomToShape(sh, g);
      }
    }
    for (const [partPath, partData] of pk.files.entries()) {
      if (typeof partData !== 'string' || !/^ppt\/(slideMasters|slideLayouts)\/.*\.xml$/.test(partPath)) continue;
      const partDoc = pk.doc(partPath);
      for (const sh of shapesOf(partDoc)) {
        const g = readXfrm(xfrmOf(sh)); if (!g) continue;
        g.x = Math.round(g.x * scaleX); g.w = Math.round(g.w * scaleX); writeGeomToShape(sh, g);
      }
    }
  }

  pushHistory('aspectRatio:' + ratio);
  renderPptxEditor();
  toast('تم تغيير مقاس الشريحة');
}

/* ── التخطيطات، الملاحظات، الانتقالات، الإخفاء، والترقيم ── */
function availableLayouts() {
  const out = [];
  for (const [path, val] of pk.files.entries()) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path) && typeof val === 'string') {
      const d = pk.doc(path);
      const cSld = kid(d.documentElement, P, 'cSld');
      const name = (cSld && cSld.getAttribute('name')) || path.split('/').pop().replace('.xml', '');
      out.push({ path, name });
    }
  }
  out.sort((a, b) => a.path.localeCompare(b.path, undefined, { numeric: true }));
  return out;
}

function getSlideNotes(slidePath) {
  const rels = readRels(pk, slidePath);
  const nRel = relByType(rels, 'notesSlide');
  if (!nRel || !nRel.path || !pk.has(nRel.path)) return '';
  const doc = pk.doc(nRel.path);
  for (const sh of shapesOf(doc)) {
    const ph = phOf(sh);
    if (ph && ph.type === 'body') {
      const txb = txBodyOf(sh);
      if (txb) return kids(txb, A, 'p').map(p => desc(p, A, 't').map(t => t.textContent).join('')).join('\n');
    }
  }
  return '';
}

function setSlideNotes(slidePath, text) {
  const rels = readRels(pk, slidePath);
  let nRel = relByType(rels, 'notesSlide');
  let nPath = nRel && nRel.path;
  if (!nPath || !pk.has(nPath)) {
    nPath = uniquePath(pk, 'ppt/notesSlides', 'notesSlide', 'xml');
    const xml = XMLDECL + `<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t></a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
    pk.setText(nPath, xml);
    setOverride(pk, '/' + nPath, CT_NOTES);
    addRel(pk, slidePath, REL + 'notesSlide', nPath);
    addRel(pk, nPath, REL + 'slide', slidePath);
  }
  const doc = pk.doc(nPath);
  const targetSh = shapesOf(doc).find(s => phOf(s)?.type === 'body');
  if (targetSh) setShapePlainText(targetSh, text);
}

function getSlideTransition(slidePath) {
  const doc = pk.doc(slidePath);
  const tr = kid(doc.documentElement, P, 'transition');
  if (!tr) return 'none';
  if (kid(tr, P, 'fade')) return 'fade';
  if (kid(tr, P, 'push')) return 'push';
  if (kid(tr, P, 'wipe')) return 'wipe';
  if (kid(tr, P, 'zoom')) return 'zoom';
  return 'none';
}

function setSlideTransition(slidePath, type) {
  const doc = pk.doc(slidePath);
  const root = doc.documentElement;
  const old = kid(root, P, 'transition');
  if (old) removeNode(old);
  if (type && type !== 'none') {
    const tr = mk(doc, P, 'p:transition', { spd: 'med' });
    tr.appendChild(mk(doc, P, 'p:' + type));
    const clr = kid(root, P, 'clrMapOvr');
    if (clr) insertAfter(clr, tr);
    else root.appendChild(tr);
  }
}

function toggleHideSlide(deck, idx) {
  const sl = deck.slides[idx];
  if (!sl) return;
  if (sl.el.getAttribute('show') === '0') sl.el.removeAttribute('show');
  else sl.el.setAttribute('show', '0');
  pushHistory('toggleHide');
  renderPptxEditor();
  toast(sl.el.getAttribute('show') === '0' ? 'تم إخفاء الشريحة من العرض' : 'تم إظهار الشريحة');
}

function addSlideNumbersAndFooter(deck) {
  const promptMsg = currentLanguage === 'en'
    ? 'Enter footer text (leave empty for slide numbers only):'
    : 'اكتب نص التذييل (أو اتركه فارغاً لترقيم الشرائح فقط):';
  const footerTxt = prompt(promptMsg, currentLanguage === 'en' ? 'DOC STUDIO' : 'محرر المستندات');
  if (footerTxt == null) return;
  const total = deck.slides.length;
  deck.slides.forEach((sl, i) => {
    const doc = pk.doc(sl.path);
    const spTree = spTreeOf(doc);
    for (const s of shapesOf(doc)) {
      const name = kid(nvOf(s), P, 'cNvPr')?.getAttribute('name') || '';
      if (name === 'DocStudioFooter' || name === 'DocStudioSlideNum') removeNode(s);
    }
    const id1 = nextShapeId(spTree);
    const numXml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
      <p:nvSpPr><p:cNvPr id="${id1}" name="DocStudioSlideNum"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="${Math.round(deck.W * 0.85)}" y="${Math.round(deck.H * 0.92)}"/><a:ext cx="${Math.round(deck.W * 0.12)}" cy="${Math.round(deck.H * 0.06)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="r"/><a:r><a:rPr sz="1200"><a:solidFill><a:srgbClr val="7B7365"/></a:solidFill></a:rPr><a:t>${i + 1} / ${total}</a:t></a:r></a:p></p:txBody>
    </p:sp>`;
    spTree.appendChild(frag(doc, numXml));

    if (footerTxt.trim()) {
      const id2 = nextShapeId(spTree);
      const footXml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
        <p:nvSpPr><p:cNvPr id="${id2}" name="DocStudioFooter"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="${Math.round(deck.W * 0.04)}" y="${Math.round(deck.H * 0.92)}"/><a:ext cx="${Math.round(deck.W * 0.5)}" cy="${Math.round(deck.H * 0.06)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200"><a:solidFill><a:srgbClr val="7B7365"/></a:solidFill></a:rPr><a:t>${escXml(footerTxt.trim())}</a:t></a:r></a:p></p:txBody>
      </p:sp>`;
      spTree.appendChild(frag(doc, footXml));
    }
  });
  pushHistory('footerNum');
  renderPptxEditor();
  toast('تمت إضافة الترقيم والتذييل لجميع الشرائح');
}

/* ── قطارة الألوان ── */
async function pickColorFromScreenOrSlide(onColorPicked) {
  if (typeof window.EyeDropper === 'function') {
    try {
      const ed = new window.EyeDropper();
      const res = await ed.open();
      if (res && res.sRGBHex) {
        onColorPicked(res.sRGBHex.toUpperCase());
        toast('تم التقاط اللون');
      }
      return;
    } catch (_) {
      return;
    }
  }

  const stage = $('#app .stagewrap .slide');
  if (!stage || !curPptx.deck) return;
  toast('انقر على أي نقطة في الشريحة لالتقاط لونها');
  let canvas;
  try {
    canvas = await slideToCanvas(curPptx.deck, curPptx.slideIdx);
  } catch (_) {
    return;
  }
  const prevCursor = stage.style.cursor;
  stage.style.cursor = 'crosshair';
  const handler = (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    stage.style.cursor = prevCursor;
    stage.removeEventListener('pointerdown', handler, true);
    const rect = stage.getBoundingClientRect();
    const px = Math.max(0, Math.min(canvas.width - 1, Math.floor(((ev.clientX - rect.left) / (rect.width || 1)) * canvas.width)));
    const py = Math.max(0, Math.min(canvas.height - 1, Math.floor(((ev.clientY - rect.top) / (rect.height || 1)) * canvas.height)));
    const d = canvas.getContext('2d').getImageData(px, py, 1, 1).data;
    const hex = '#' + rgb2hex(d[0], d[1], d[2]);
    onColorPicked(hex);
    toast('تم التقاط اللون');
  };
  stage.addEventListener('pointerdown', handler, true);
}

function makeColorWithEyedropper(colorInput, onChange) {
  const dropBtn = h('button', {
    type: 'button',
    class: 'btn eyedrop-btn',
    title: tr('💉 قطارة الألوان'),
    onclick: () => pickColorFromScreenOrSlide((hex) => {
      colorInput.value = hex;
      if (onChange) onChange(hex);
    })
  }, '💉');
  return h('span', { class: 'color-pair' }, colorInput, dropBtn);
}

/* ── مكتبة الأيقونات الجاهزة ── */
const BUILTIN_ICONS = [
  { id: 'check', label: 'صح', glyph: '✔', path: 'M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z' },
  { id: 'close', label: 'خطأ', glyph: '✖', path: 'M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z' },
  { id: 'star', label: 'نجمة', glyph: '★', path: 'M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z' },
  { id: 'location', label: 'موقع', glyph: '📍', path: 'M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5a2.5 2.5 0 010-5 2.5 2.5 0 010 5z' },
  { id: 'phone', label: 'هاتف', glyph: '☎', path: 'M6.62 10.79a15.05 15.05 0 006.59 6.59l2.2-2.2a1 1 0 011.01-.24c1.12.37 2.33.57 3.58.57a1 1 0 011 1V20a1 1 0 01-1 1C10.07 21 3 13.93 3 5a1 1 0 011-1h3.5a1 1 0 011 1c0 1.25.2 2.46.57 3.58a1 1 0 01-.25 1.01l-2.2 2.2z' },
  { id: 'email', label: 'بريد', glyph: '✉', path: 'M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z' },
  { id: 'alert', label: 'تنبيه', glyph: '⚠', path: 'M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z' },
  { id: 'heart', label: 'قلب', glyph: '♥', path: 'M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z' },
  { id: 'gear', label: 'ترس', glyph: '⚙', path: 'M19.14 12.94c.04-.31.06-.63.06-.94s-.02-.63-.06-.94l2.03-1.58a.49.49 0 00.12-.61l-1.92-3.32a.49.49 0 00-.59-.22l-2.39.96a7.03 7.03 0 00-1.62-.94l-.36-2.54a.48.48 0 00-.48-.41h-3.84a.48.48 0 00-.48.41l-.36 2.54a7.03 7.03 0 00-1.62.94l-2.39-.96a.49.49 0 00-.59.22L2.74 8.87a.48.48 0 00.12.61l2.03 1.58c-.04.31-.06.63-.06.94s.02.63.06.94l-2.03 1.58a.48.48 0 00-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.04.7 1.62.94l.36 2.54c.05.24.26.41.48.41h3.84c.24 0 .44-.17.48-.41l.36-2.54c.58-.24 1.12-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32a.48.48 0 00-.12-.61l-2.01-1.58zM12 15.6A3.6 3.6 0 1112 8.4a3.6 3.6 0 010 7.2z' },
  { id: 'arrow', label: 'سهم', glyph: '➜', path: 'M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z' }
];

function closeIconsPopover() {
  const pop = $('#iconsPopover');
  if (pop) pop.hidden = true;
}

function toggleIconsPopover(anchorBtn, deck, slideIdx) {
  const pop = $('#iconsPopover');
  if (!pop) return;
  if (!pop.hidden) { pop.hidden = true; return; }
  pop.innerHTML = '';
  const head = h('div', { class: 'icons-popover-head' },
    h('span', null, tr('مكتبة الأيقونات')),
    h('button', { type: 'button', class: 'btn', style: { padding: '2px 6px', fontSize: '.75rem' }, onclick: closeIconsPopover }, '✕')
  );
  const grid = h('div', { class: 'icons-grid' });
  BUILTIN_ICONS.forEach(ic => {
    const svg = svgEl('svg', { viewBox: '0 0 24 24' });
    svg.appendChild(svgEl('path', { d: ic.path }));
    const cell = h('button', {
      type: 'button',
      class: 'icon-cell',
      title: tr(ic.label),
      onclick: () => {
        closeIconsPopover();
        addIconToSlide(deck, slideIdx, ic);
      }
    }, svg, h('span', null, tr(ic.label)));
    grid.append(cell);
  });
  pop.append(head, grid);
  const rect = anchorBtn.getBoundingClientRect();
  pop.hidden = false;
  const left = Math.max(10, Math.min(window.innerWidth - 285, rect.left));
  pop.style.left = left + 'px';
  pop.style.top = (rect.bottom + 6) + 'px';
}

/* ── قائمة الزر الأيمن المخصصة ── */
function closeContextMenu() {
  const m = $('#ctxMenu');
  if (m) m.hidden = true;
}

function openContextMenu(x, y, items) {
  const menu = $('#ctxMenu');
  if (!menu) return;
  menu.innerHTML = '';
  for (const it of items) {
    if (it === 'sep') {
      menu.append(h('div', { class: 'ctx-sep' }));
      continue;
    }
    const b = h('button', {
      type: 'button',
      class: 'ctx-item' + (it.danger ? ' danger' : ''),
      disabled: !!it.disabled,
      onclick: () => {
        closeContextMenu();
        if (it.action) it.action();
      }
    }, h('span', null, tr(it.label)), it.hint ? h('small', { class: 'muted' }, it.hint) : null);
    menu.append(b);
  }
  menu.hidden = false;
  const mw = menu.offsetWidth || 200;
  const mh = menu.offsetHeight || 220;
  const left = Math.max(8, Math.min(window.innerWidth - mw - 8, x));
  const top = Math.max(8, Math.min(window.innerHeight - mh - 8, y));
  menu.style.left = left + 'px';
  menu.style.top = top + 'px';
}

document.addEventListener('pointerdown', (e) => {
  const menu = $('#ctxMenu');
  if (menu && !menu.hidden && !menu.contains(e.target)) closeContextMenu();
  const pop = $('#iconsPopover');
  if (pop && !pop.hidden && !pop.contains(e.target)) closeIconsPopover();
});
window.addEventListener('blur', () => { closeContextMenu(); closeIconsPopover(); });

/* ── شريط تنسيق الكلمات المظللة داخل الشريحة ── */
function buildInlineTextToolbar() {
  const bar = h('div', { class: 'inline-fmt-bar', hidden: true, onpointerdown: e => e.preventDefault() });
  const col = h('input', { type: 'color', value: '#B4531A', title: tr('لون النص المظلل') });
  const sz = h('input', { type: 'number', min: 8, max: 200, value: 24, title: tr('حجم النص المظلل') });

  const execInline = (cmd, val) => {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd, false, val || null);
  };

  bar.append(
    h('span', { style: { fontSize: '.78rem', opacity: '.8' } }, tr('تنسيق المحدد:')),
    btn('B', () => execInline('bold'), '', { keep: true }),
    btn('I', () => execInline('italic'), '', { keep: true }),
    btn('U', () => execInline('underline'), '', { keep: true }),
    col,
    btn('لون', () => execInline('foreColor', col.value), '', { keep: true }),
    sz,
    btn('حجم', () => {
      const sel = window.getSelection();
      if (!sel || !sel.rangeCount || sel.isCollapsed) return;
      const range = sel.getRangeAt(0);
      const span = document.createElement('span');
      span.style.fontSize = `${(+sz.value || 24) / CUR_WPT * 100}cqw`;
      try {
        range.surroundContents(span);
      } catch (_) {
        const fragContent = range.extractContents();
        span.appendChild(fragContent);
        range.insertNode(span);
      }
    }, '', { keep: true })
  );
  return bar;
}

function parseCssColorToHex(str) {
  if (!str) return null;
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(str);
  if (m) return rgb2hex(+m[1], +m[2], +m[3]);
  if (/^#[0-9a-f]{6}$/i.test(str.trim())) return str.trim().slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(str.trim())) {
    const s = str.trim().slice(1);
    return (s[0] + s[0] + s[1] + s[1] + s[2] + s[2]).toUpperCase();
  }
  return null;
}

function syncRichHtmlToShapeXml(sh, tbEl, ctx) {
  const txb = txBodyOf(sh);
  if (!txb) return;
  const doc = sh.ownerDocument;
  const existingPs = kids(txb, A, 'p');
  const basePPr = existingPs[0] && kid(existingPs[0], A, 'pPr')?.cloneNode(true);
  const baseRPr = desc(txb, A, 'rPr')[0]?.cloneNode(true);

  for (const p of existingPs) removeNode(p);

  // Group direct child nodes into logical paragraph containers so text typed directly
  // into tbEl (e.g. on newly created empty shapes) is never lost alongside .tp divs.
  const paraContainers = [];
  let inlineBuf = null;
  for (const ch of Array.from(tbEl.childNodes)) {
    if (ch.nodeType === 1 && (ch.tagName === 'DIV' || ch.tagName === 'P')) {
      inlineBuf = null;
      paraContainers.push(ch);
    } else if (ch.nodeType === 1 && ch.tagName === 'BR') {
      inlineBuf = null;
      paraContainers.push(h('div', null));
    } else {
      const hasTxt = ch.nodeType === 3 ? !!ch.nodeValue?.replace(/\u200b/g, '') : !!ch.textContent?.replace(/\u200b/g, '');
      if (!hasTxt && !inlineBuf) continue;
      if (!inlineBuf) {
        inlineBuf = document.createElement('div');
        paraContainers.push(inlineBuf);
      }
      inlineBuf.appendChild(ch.cloneNode(true));
    }
  }
  if (!paraContainers.length) {
    setShapePlainText(sh, tbEl.innerText.replace(/\u200b/g, ''));
    return;
  }

  for (const div of paraContainers) {
    const p = mk(doc, A, 'a:p');
    if (basePPr) p.appendChild(basePPr.cloneNode(true));

    const walk = (node, inherited) => {
      if (node.nodeType === 3) {
        const txt = node.nodeValue.replace(/\u200b/g, '');
        if (!txt) return;
        const r = mk(doc, A, 'a:r');
        const rp = baseRPr ? baseRPr.cloneNode(true) : mk(doc, A, 'a:rPr', { sz: '2400' });
        if (inherited.b != null) rp.setAttribute('b', inherited.b ? '1' : '0');
        if (inherited.i != null) rp.setAttribute('i', inherited.i ? '1' : '0');
        if (inherited.u != null) rp.setAttribute('u', inherited.u ? 'sng' : 'none');
        if (inherited.sz) rp.setAttribute('sz', String(Math.round(inherited.sz * 100)));
        if (inherited.hex) {
          for (const f of kids(rp, A)) if (f.localName === 'solidFill') removeNode(f);
          const sf = mk(doc, A, 'a:solidFill');
          sf.appendChild(mk(doc, A, 'a:srgbClr', { val: inherited.hex }));
          rp.appendChild(sf);
        }
        r.appendChild(rp);
        const t = mk(doc, A, 'a:t');
        if (/^\s|\s$/.test(txt)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        t.textContent = txt;
        r.appendChild(t);
        p.appendChild(r);
      } else if (node.nodeType === 1) {
        if (node.classList?.contains('bu')) return;
        if (node.tagName === 'BR') {
          p.appendChild(mk(doc, A, 'a:br'));
          return;
        }
        const next = { ...inherited };
        const st = node.style || {};
        if (node.tagName === 'B' || node.tagName === 'STRONG' || st.fontWeight === '700' || st.fontWeight === 'bold') next.b = true;
        if (node.tagName === 'I' || node.tagName === 'EM' || st.fontStyle === 'italic') next.i = true;
        if (node.tagName === 'U' || (st.textDecoration && st.textDecoration.includes('underline'))) next.u = true;
        const colHex = parseCssColorToHex(st.color) || (node.tagName === 'FONT' ? parseCssColorToHex(node.getAttribute('color')) : null);
        if (colHex) next.hex = colHex;
        if (st.fontSize) {
          if (st.fontSize.endsWith('cqw')) {
            const cqw = parseFloat(st.fontSize);
            if (!isNaN(cqw)) next.sz = (cqw / 100) * CUR_WPT;
          } else if (st.fontSize.endsWith('pt')) {
            const pt = parseFloat(st.fontSize);
            if (!isNaN(pt)) next.sz = pt;
          }
        }
        for (const ch of Array.from(node.childNodes)) walk(ch, next);
      }
    };

    walk(div, {});
    if (!kids(p, A, 'r').length) {
      const r = mk(doc, A, 'a:r');
      if (baseRPr) r.appendChild(baseRPr.cloneNode(true));
      r.appendChild(mk(doc, A, 'a:t'));
      p.appendChild(r);
    }
    txb.appendChild(p);
  }
}

/* ── واجهة محرر PPTX وعرض الشبكة ── */
function renderPptxEditor() {
  syncEditorTopBarForKind();
  const app = $('#app');
  const prevThumbs = app.querySelector('.thumbs');
  const prevProps = app.querySelector('.props');
  const prevScroll = prevThumbs ? prevThumbs.scrollTop : 0;
  const prevPropsScroll = prevProps ? prevProps.scrollTop : 0;
  app.innerHTML = '';

  const deck = readDeck();
  curPptx.deck = deck;
  if (!deck.slides.length) {
    app.append(h('div', { class: 'doc' }, h('p', { class: 'alert' }, tr('لا يحتوي هذا الملف على شرائح قابلة للعرض.'))));
    return;
  }
  if (curPptx.slideIdx >= deck.slides.length) curPptx.slideIdx = deck.slides.length - 1;
  if (curPptx.slideIdx < 0) curPptx.slideIdx = 0;

  const currentSlide = deck.slides[curPptx.slideIdx];
  const wrap = h('div', { class: 'ppt' });
  const thumbs = h('div', { class: 'thumbs' });
  const center = h('div');
  const stageWrap = h('div', { class: 'stagewrap' });
  const props = h('div', { class: 'props' });

  deck.slides.forEach((sl, i) => {
    const trType = getSlideTransition(sl.path);
    const th = h('div', {
      class: 'thumb' + (i === curPptx.slideIdx ? ' on' : '') + (sl.hidden ? ' slide-hidden' : ''),
      draggable: true,
      onclick: () => {
        if (curPptx.slideIdx === i && curPptx.viewMode === 'normal') return;
        curPptx.slideIdx = i; curPptx.selShapeIdx = null; curPptx.multiSel.clear(); curPptx.selCell = null; curPptx.viewMode = 'normal'; renderPptxEditor();
      },
      oncontextmenu: (e) => {
        e.preventDefault();
        e.stopPropagation();
        curPptx.slideIdx = i;
        openContextMenu(e.clientX, e.clientY, [
          { label: '📄 نسخ', action: () => duplicateSlide(deck, i) },
          { label: '👁 إخفاء/إظهار', action: () => toggleHideSlide(deck, i) },
          'sep',
          { label: '🗑 حذف شريحة', danger: true, disabled: deck.slides.length <= 1, action: () => deleteSlide(deck, i) }
        ]);
      },
      ondragstart: (e) => { e.dataTransfer.setData('text/plain', String(i)); e.dataTransfer.effectAllowed = 'move'; th.classList.add('dragging'); },
      ondragend: () => th.classList.remove('dragging'),
      ondragover: (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; },
      ondrop: (e) => { e.preventDefault(); const from = Number(e.dataTransfer.getData('text/plain')); if (Number.isInteger(from) && from !== i) moveSlide(deck, from, i); }
    },
      h('span', { class: 'num' }, String(i + 1)),
      sl.hidden ? h('span', { class: 'hid-badge' }, tr('مخفية')) : null,
      trType !== 'none' ? h('span', { class: 'tr-badge' }, '✨ ' + trType) : null
    );
    try { th.append(renderSlideEl(deck, i, false)); } catch (_) { th.append(h('div', { class: 'slide' }, tr('شريحة ' + (i + 1)))); }
    thumbs.append(th);
  });

  const ribbon = buildPptxRibbon(deck, currentSlide, stageWrap);
  const inlineFmtBar = buildInlineTextToolbar();

  if (curPptx.viewMode === 'grid') {
    const grid = h('div', { class: 'grid-sorter' });
    deck.slides.forEach((sl, i) => {
      const card = h('div', {
        class: 'grid-card' + (i === curPptx.slideIdx ? ' on' : '') + (sl.hidden ? ' slide-hidden' : ''),
        onclick: () => { curPptx.slideIdx = i; curPptx.viewMode = 'normal'; renderPptxEditor(); }
      },
        renderSlideEl(deck, i, false),
        h('div', { class: 'grid-card-foot' },
          h('span', null, tr(`شريحة ${i + 1}${sl.hidden ? ' (مخفية)' : ''}`)),
          btn(sl.hidden ? 'إظهار' : 'إخفاء', (e) => { e.stopPropagation(); toggleHideSlide(deck, i); })
        )
      );
      grid.append(card);
    });
    center.append(ribbon, grid);
  } else {
    const stage = renderSlideEl(deck, curPptx.slideIdx, true);
    stage.style.width = curPptx.zoom + '%';
    stageWrap.append(stage);

    const notesPanel = h('details', { class: 'notes-panel' },
      h('summary', null, tr('📝 ملاحظات المحاضر (Speaker Notes)')),
      h('textarea', {
        placeholder: tr('اكتب ملاحظاتك لهذه الشريحة هنا…'),
        value: getSlideNotes(currentSlide.path),
        onchange: (e) => { setSlideNotes(currentSlide.path, e.target.value); pushHistory('notes'); toast('تم حفظ ملاحظات الشريحة'); syncPresenterWindow(); }
      })
    );

    center.append(ribbon, inlineFmtBar, stageWrap, notesPanel);
    bindSlideInteractions(stage, deck, curPptx.slideIdx, props, inlineFmtBar);
  }

  wrap.append(thumbs, center, props);
  app.append(wrap);
  thumbs.scrollTop = prevScroll;

  renderShapeProps(props, center.querySelector('.slide'), deck, curPptx.slideIdx);
  props.scrollTop = prevPropsScroll;
  syncPresenterWindow();
}

function refreshPptxRibbon(deck, slideIdx) {
  const oldRibbon = $('#app .ribbon');
  const stageWrap = $('#app .stagewrap');
  if (!oldRibbon || !deck || !deck.slides[slideIdx]) return;
  oldRibbon.replaceWith(buildPptxRibbon(deck, deck.slides[slideIdx], stageWrap));
}

function buildPptxRibbon(deck, currentSlide, stageWrap) {
  const ribbon = h('div', { class: 'ribbon' });
  const tabsBar = h('div', { class: 'ribbon-tabs' });
  const body = h('div', { class: 'ribbon-body' });

  [['home', 'الرئيسية'], ['insert', 'إدراج'], ['design', 'تصميم'], ['show', 'عرض وتصدير']].forEach(([id, label]) => {
    tabsBar.append(h('button', {
      class: 'rtab' + (curPptx.activeTab === id ? ' active' : ''),
      onclick: () => { curPptx.activeTab = id; renderPptxEditor(); }
    }, tr(label)));
  });

  const zoomSel = h('select', { class: 'zoom-select' }, ...['50', '75', '90', '100', '125', '150'].map(v => h('option', { value: v }, v + '%')));
  zoomSel.value = curPptx.zoom;
  zoomSel.onchange = () => {
    curPptx.zoom = zoomSel.value;
    const sl = stageWrap.querySelector('.slide');
    if (sl) sl.style.width = curPptx.zoom + '%';
  };

  tabsBar.append(
    h('span', { style: { flex: '1' } }),
    btn(curPptx.viewMode === 'grid' ? '📄 عرض عادي' : '⊞ عرض الشبكة', () => {
      curPptx.viewMode = curPptx.viewMode === 'grid' ? 'normal' : 'grid';
      renderPptxEditor();
    }),
    h('span', { class: 'bar-label' }, tr('تكبير')),
    zoomSel,
    btn('◀ السابق', () => { if (curPptx.slideIdx > 0) { curPptx.slideIdx--; curPptx.selShapeIdx = null; renderPptxEditor(); } }, '', { disabled: curPptx.slideIdx <= 0 }),
    btn('التالي ▶', () => { if (curPptx.slideIdx < deck.slides.length - 1) { curPptx.slideIdx++; curPptx.selShapeIdx = null; renderPptxEditor(); } }, '', { disabled: curPptx.slideIdx >= deck.slides.length - 1 })
  );

  const hasSel = curPptx.selShapeIdx != null;
  const hasMulti = curPptx.multiSel.size >= 2;
  const selSh = hasSel ? shapesOf(pk.doc(currentSlide.path))[curPptx.selShapeIdx] : null;
  const isGroup = selSh && selSh.localName === 'grpSp';

  if (curPptx.activeTab === 'home') {
    body.append(
      btn('➕ شريحة', () => addSlide(deck), 'pri'),
      btn('📄 نسخ', () => duplicateSlide(deck, curPptx.slideIdx)),
      btn('👁 إخفاء/إظهار', () => toggleHideSlide(deck, curPptx.slideIdx)),
      btn('🗑 حذف شريحة', () => deleteSlide(deck, curPptx.slideIdx), 'danger', { disabled: deck.slides.length <= 1 }),
      btn('⬆️ للأعلى', () => moveSlide(deck, curPptx.slideIdx, curPptx.slideIdx - 1), '', { disabled: curPptx.slideIdx <= 0 }),
      btn('⬇️ للأسفل', () => moveSlide(deck, curPptx.slideIdx, curPptx.slideIdx + 1), '', { disabled: curPptx.slideIdx >= deck.slides.length - 1 }),
      h('span', { class: 'bar-sep' }),
      btn('📋 نسخ عنصر', () => copySelectedShape(deck, curPptx.slideIdx), '', { disabled: !hasSel }),
      btn('📥 لصق', () => pasteShapeToSlide(deck, curPptx.slideIdx), '', { disabled: !curPptx.clipboardXml }),
      btn('🖌 نسخ التنسيق', () => startFormatPainter(deck, curPptx.slideIdx), curPptx.copiedStyle ? 'on' : '', { disabled: !hasSel && !curPptx.copiedStyle }),
      btn('🗑 حذف عنصر', () => deleteSelectedShape(deck, curPptx.slideIdx), 'danger', { disabled: !hasSel && !hasMulti }),
      h('span', { class: 'bar-sep' }),
      btn('⬆ للأمام', () => reorderShapeLayer(deck, curPptx.slideIdx, 'front'), '', { disabled: !hasSel }),
      btn('⬇ للخلف', () => reorderShapeLayer(deck, curPptx.slideIdx, 'back'), '', { disabled: !hasSel }),
      btn('🔗 تجميع', () => groupSelectedShapes(deck, curPptx.slideIdx), '', { disabled: !hasMulti }),
      btn('🔓 فك التجميع', () => ungroupSelectedShape(deck, curPptx.slideIdx), '', { disabled: !isGroup })
    );
  } else if (curPptx.activeTab === 'insert') {
    const iconBtn = h('button', {
      type: 'button',
      class: 'btn',
      onclick: (e) => { e.stopPropagation(); toggleIconsPopover(iconBtn, deck, curPptx.slideIdx); }
    }, tr('✨ أيقونات'));

    const saSel = h('select', { class: 'bar-select' },
      h('option', { value: '' }, tr('💠 مخطط SmartArt…')),
      h('option', { value: 'pyramid' }, tr('هرم تنظيمي')),
      h('option', { value: 'cycle' }, tr('دورة مستمرة')),
      h('option', { value: 'process' }, tr('خطوات متتالية'))
    );
    saSel.onchange = () => {
      if (saSel.value) insertSmartArtDiagram(saSel.value);
      saSel.value = '';
    };

    body.append(
      btn('➕ مربع نص', () => addTextBoxToSlide(deck, curPptx.slideIdx), 'pri'),
      btn('🖼 صورة', () => triggerImageInsert((info) => addImageToSlide(deck, curPptx.slideIdx, info))),
      btn('🖼+ صور لشرائح', () => triggerBatchImagesToSlides(deck)),
      btn('▦ جدول', () => addTableToSlide(deck, curPptx.slideIdx)),
      btn('📊 رسم بياني', () => addChartToSlide(deck, curPptx.slideIdx)),
      btn('🎬 فيديو/صوت', () => triggerMediaInsert((info) => addMediaToSlide(deck, curPptx.slideIdx, info))),
      h('span', { class: 'bar-sep' }),
      iconBtn,
      btn('✨ شرائح ذكية', () => openSmartSlidesModal(deck, curPptx.slideIdx), 'pri'),
      saSel,
      btn('▭ مستطيل', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'roundRect')),
      btn('◯ دائرة', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'ellipse')),
      btn('▲ مثلث', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'triangle')),
      btn('◆ معين', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'diamond')),
      btn('⬡ سداسي', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'hexagon')),
      btn('★ نجمة', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'star5')),
      btn('💬 فقاعة', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'wedgeRoundRectCallout')),
      btn('➡ سهم', () => addPresetShapeToSlide(deck, curPptx.slideIdx, 'rightArrow')),
      btn('╱ خط', () => addLineToSlide(deck, curPptx.slideIdx))
    );
  } else if (curPptx.activeTab === 'design') {
    const thSel = h('select', { class: 'bar-select' },
      h('option', { value: '' }, tr('🎨 ثيم العرض الكامل…')),
      ...Object.entries(PPTX_PRO_THEMES).map(([k, v]) => h('option', { value: k }, tr(v.name)))
    );
    thSel.onchange = () => {
      if (thSel.value) applyGlobalDeckTheme(thSel.value);
      thSel.value = '';
    };

    const layouts = availableLayouts();
    const curLayout = slideLayoutPath(currentSlide.path) || '';
    const laySel = h('select', { class: 'bar-select' }, ...layouts.map(l => h('option', { value: l.path }, l.name)));
    if (curLayout) laySel.value = curLayout;
    laySel.onchange = () => { changeSlideLayout(currentSlide.path, laySel.value); pushHistory('layout'); renderPptxEditor(); toast('تم تغيير تخطيط الشريحة'); };

    const is43 = Math.abs((deck.W / deck.H) - (4 / 3)) < 0.08;
    const sizeSel = h('select', { class: 'bar-select' },
      h('option', { value: '16:9' }, tr('16:9 شاشة عريضة')),
      h('option', { value: '4:3' }, tr('4:3 قياسي'))
    );
    sizeSel.value = is43 ? '4:3' : '16:9';
    sizeSel.onchange = () => setDeckAspectRatio(deck, sizeSel.value);

    const curBg = backgroundOf(buildCtx(deck, currentSlide.path))?.fill;
    const initBg1 = (curBg?.type === 'solid') ? ('#' + curBg.hex) : (curBg?.type === 'grad' && curBg.rawStops?.[0] ? '#' + curBg.rawStops[0].hex : '#FFFFFF');
    const initBg2 = (curBg?.type === 'grad' && curBg.rawStops?.[1]) ? ('#' + curBg.rawStops[1].hex) : '#E2E8F0';
    const initBgAng = (curBg?.type === 'grad') ? (curBg.angle || 0) : 90;

    const bgColor1 = h('input', { type: 'color', value: initBg1, title: tr('لون 1'), style: { width: '32px', height: '28px', border: '1px solid var(--line)', borderRadius: '6px' } });
    const bgColor2 = h('input', { type: 'color', value: initBg2, title: tr('لون 2'), style: { width: '32px', height: '28px', border: '1px solid var(--line)', borderRadius: '6px' } });
    const bgAngle = h('input', { type: 'number', min: 0, max: 360, value: initBgAng, title: tr('الزاوية (°)'), style: { width: '54px', border: '1px solid var(--line)', borderRadius: '6px', padding: '4px 5px', background: 'var(--card)', fontSize: '.82rem' } });

    const trSel = h('select', { class: 'bar-select' },
      h('option', { value: 'none' }, tr('بدون انتقال')),
      h('option', { value: 'fade' }, tr('تلاشي (Fade)')),
      h('option', { value: 'push' }, tr('دفع (Push)')),
      h('option', { value: 'wipe' }, tr('مسح (Wipe)')),
      h('option', { value: 'zoom' }, tr('تكبير (Zoom)'))
    );
    trSel.value = getSlideTransition(currentSlide.path);
    trSel.onchange = () => { setSlideTransition(currentSlide.path, trSel.value); pushHistory('transition'); renderPptxEditor(); toast('تم ضبط الانتقال'); };

    body.append(
      thSel,
      h('span', { class: 'bar-sep' }),
      btn('📐 شبكة التصميم', () => {
        curPptx.showGrid = !curPptx.showGrid;
        renderPptxEditor();
      }, curPptx.showGrid ? 'on' : ''),
      h('span', { class: 'bar-sep' }),
      h('span', { class: 'bar-label' }, tr('التخطيط:')), laySel,
      h('span', { class: 'bar-label' }, tr('مقاس الشريحة:')), sizeSel,
      h('span', { class: 'bar-sep' }),
      makeColorWithEyedropper(bgColor1),
      btn('🎨 لون الخلفية', () => setSlideBackgroundColor(currentSlide.path, bgColor1.value)),
      bgColor2,
      bgAngle,
      btn('🌈 تدرج الخلفية', () => setSlideBackgroundGradient(currentSlide.path, bgColor1.value, bgColor2.value, +bgAngle.value || 0)),
      btn('🖼 صورة خلفية', () => triggerBgImageInsert((info) => setSlideBackgroundImage(currentSlide.path, info))),
      btn('تطبيق على الكل', () => applyBgToAllSlides(deck, currentSlide.path)),
      h('span', { class: 'bar-sep' }),
      btn('🔢 ترقيم وتذييل الكل', () => addSlideNumbersAndFooter(deck)),
      h('span', { class: 'bar-sep' }),
      h('span', { class: 'bar-label' }, tr('انتقال الشريحة:')), trSel
    );
  } else if (curPptx.activeTab === 'show') {
    body.append(
      btn('▶ بدء العرض', () => startSlideshow(deck, curPptx.slideIdx), 'pri'),
      btn('🖥 وضع مقدم العرض', () => openPresenterView(deck, curPptx.slideIdx)),
      h('span', { class: 'bar-sep' }),
      btn('🖼 تصدير PNG', () => exportCurrentSlidePng(deck, curPptx.slideIdx)),
      btn('📦 تصدير الكل ZIP', () => exportAllSlidesZip(deck)),
      btn('📄 تصدير PDF', () => exportDeckToPdf(deck)),
      h('span', { class: 'bar-sep' }),
      btn('📄 تصدير إلى Word', () => exportPptxToDocx(deck))
    );
  }

  ribbon.append(tabsBar, body);
  return ribbon;
}

/* ── التفاعل المباشر مع العناصر + التعديل المزدوج + التحجيم بالماوس ── */
function bindSlideInteractions(stage, deck, slideIdx, propsPanel, inlineFmtBar) {
  const shapeEls = $$('.shape:not(.static)', stage);

  shapeEls.forEach((el) => {
    const idx = +(el.dataset.i || -1);
    if (idx >= 0 && curPptx.selShapeIdx === idx) el.classList.add('sel');
    if (idx >= 0 && curPptx.multiSel.has(idx)) el.classList.add('multi-sel');

    if (!el.querySelector('.rh-box')) {
      const box = h('div', { class: 'rh-box' },
        h('div', { class: 'rot-handle', title: tr('تدوير') }),
        ...['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(d => h('div', { class: 'rh-pt', 'data-dir': d }))
      );
      el.append(box);
    }

    el.addEventListener('contextmenu', (e) => {
      if (idx < 0) return;
      e.preventDefault();
      e.stopPropagation();
      if (!curPptx.multiSel.has(idx)) {
        curPptx.multiSel.clear();
        curPptx.selShapeIdx = idx;
      } else {
        curPptx.selShapeIdx = idx;
      }
      shapeEls.forEach(s => {
        const si = +(s.dataset.i || -1);
        s.classList.toggle('sel', si === curPptx.selShapeIdx);
        s.classList.toggle('multi-sel', curPptx.multiSel.has(si));
      });
      renderShapeProps(propsPanel, stage, deck, slideIdx);

      const sh = shapesOf(pk.doc(deck.slides[slideIdx].path))[idx];
      const isGrp = sh && sh.localName === 'grpSp';
      const canGroup = curPptx.multiSel.size >= 2;

      openContextMenu(e.clientX, e.clientY, [
        { label: '📋 نسخ عنصر', hint: 'Ctrl+C', action: () => copySelectedShape(deck, slideIdx) },
        { label: '📥 لصق', hint: 'Ctrl+V', disabled: !curPptx.clipboardXml, action: () => pasteShapeToSlide(deck, slideIdx) },
        { label: '📄 تكرار عنصر', action: () => duplicateSelectedShape(deck, slideIdx) },
        'sep',
        { label: '⬆ للأمام', action: () => reorderShapeLayer(deck, slideIdx, 'front') },
        { label: '⬇ للخلف', action: () => reorderShapeLayer(deck, slideIdx, 'back') },
        'sep',
        { label: '🔗 تجميع', disabled: !canGroup, action: () => groupSelectedShapes(deck, slideIdx) },
        { label: '🔓 فك التجميع', disabled: !isGrp, action: () => ungroupSelectedShape(deck, slideIdx) },
        'sep',
        { label: '🗑 حذف عنصر', hint: 'Del', danger: true, action: () => deleteSelectedShape(deck, slideIdx) }
      ]);
    });

    const enterInlineShapeEdit = (e) => {
      if (['table', 'pic', 'chart', 'group', 'line'].includes(el.dataset.kind)) return;
      if (e) e.stopPropagation();
      let tb = el.querySelector('.tb');
      if (!tb && el._sh && el._sh.localName === 'sp') {
        const doc = el._sh.ownerDocument;
        const txb = mk(doc, P, 'p:txBody');
        txb.append(
          mk(doc, A, 'a:bodyPr', { anchor: 'ctr' }),
          mk(doc, A, 'a:lstStyle'),
          frag(doc, `<a:p xmlns:a="${A}"><a:pPr algn="ctr"/><a:r><a:rPr sz="2000"><a:solidFill><a:srgbClr val="1F1C17"/></a:solidFill></a:rPr><a:t></a:t></a:r></a:p>`)
        );
        el._sh.appendChild(txb);
        tb = renderTextBody(stage._ctx, el._sh, txb);
        el.insertBefore(tb, el.querySelector('.rh-box') || null);
        el.dataset.kind = 'text';
      }
      if (!tb) return;

      el.classList.add('editing');
      tb.contentEditable = 'true';
      tb.focus();
      const targetNode = tb.querySelector('.tp span') || tb.querySelector('.tp') || tb;
      try {
        const sel = window.getSelection();
        if (sel && targetNode) {
          const rng = document.createRange();
          rng.selectNodeContents(targetNode);
          rng.collapse(false);
          sel.removeAllRanges();
          sel.addRange(rng);
        }
      } catch (_) {}
      if (inlineFmtBar) inlineFmtBar.hidden = false;

      const finishEdit = () => {
        tb.contentEditable = 'false';
        el.classList.remove('editing');
        if (inlineFmtBar) inlineFmtBar.hidden = true;
        syncRichHtmlToShapeXml(el._sh, tb, stage._ctx);
        pushHistory('inlineText');
        renderPptxEditor();
      };

      tb.onblur = (ev) => {
        if (inlineFmtBar && inlineFmtBar.contains(ev.relatedTarget)) return;
        finishEdit();
      };
      tb.onkeydown = (ev) => {
        if (ev.key === 'Escape') { ev.preventDefault(); tb.blur(); }
      };
    };

    el.addEventListener('dblclick', enterInlineShapeEdit);

    el.addEventListener('pointerdown', (e) => {
      if (el.classList.contains('editing')) {
        e.stopPropagation();
        return;
      }
      if (e.pointerType === 'touch' && !e.target?.closest?.('.rh-pt, .rot-handle')) {
        const now = Date.now();
        const prevTap = el._lastTouchTap;
        if (prevTap && (now - prevTap.t) < 340 && Math.hypot(e.clientX - prevTap.x, e.clientY - prevTap.y) < 24) {
          el._lastTouchTap = null;
          e.preventDefault();
          enterInlineShapeEdit(e);
          return;
        }
        el._lastTouchTap = { t: now, x: e.clientX, y: e.clientY };
      }
      if (!el._g) return;
      if (isShapeLocked(el._sh)) {
        e.stopPropagation();
        curPptx.multiSel.clear();
        curPptx.selShapeIdx = idx;
        shapeEls.forEach(s => s.classList.toggle('sel', +(s.dataset.i || -1) === idx));
        renderShapeProps(propsPanel, stage, deck, slideIdx);
        refreshPptxRibbon(deck, slideIdx);
        return;
      }
      if (e.button !== undefined && e.button !== 0 && e.pointerType === 'mouse') return;
      e.stopPropagation();

      if (curPptx.copiedStyle && idx >= 0) {
        applyFormatPainterToShape(deck, slideIdx, idx);
        return;
      }

      const handleEl = e.target?.closest?.('.rh-pt');
      const rotEl = e.target?.closest?.('.rot-handle');
      const dir = handleEl ? handleEl.dataset.dir : null;
      const isRot = !!rotEl;

      if (e.shiftKey && !dir && !isRot) {
        if (curPptx.multiSel.has(idx)) curPptx.multiSel.delete(idx);
        else curPptx.multiSel.add(idx);
        if (curPptx.selShapeIdx != null) curPptx.multiSel.add(curPptx.selShapeIdx);
        curPptx.selShapeIdx = idx;
      } else {
        curPptx.multiSel.clear();
        curPptx.selShapeIdx = idx;
      }

      shapeEls.forEach(s => {
        const si = +(s.dataset.i || -1);
        s.classList.toggle('sel', si === curPptx.selShapeIdx);
        s.classList.toggle('multi-sel', curPptx.multiSel.has(si));
      });
      renderShapeProps(propsPanel, stage, deck, slideIdx);
      refreshPptxRibbon(deck, slideIdx);

      const pointerId = e.pointerId;
      const startX = e.clientX, startY = e.clientY;
      const rect = stage.getBoundingClientRect();
      const orig = { ...el._g };
      let moved = false;

      const vTargets = [deck.W / 2];
      const hTargets = [deck.H / 2];
      shapeEls.forEach(other => {
        if (other === el || !other._g || other.style.display === 'none') return;
        const og = other._g;
        vTargets.push(og.x, og.x + og.w / 2, og.x + og.w);
        hTargets.push(og.y, og.y + og.h / 2, og.y + og.h);
      });

      const clearGuides = () => $$('.smart-guide', stage).forEach(g => g.remove());
      const showGuide = (axis, pct) => {
        const g = h('div', { class: 'smart-guide ' + axis });
        if (axis === 'v') g.style.left = pct + '%';
        else g.style.top = pct + '%';
        stage.append(g);
      };

      const onMove = (ev) => {
        if (ev.pointerId !== pointerId) return;
        if (!moved && Math.hypot(ev.clientX - startX, ev.clientY - startY) < 3) return;
        moved = true;
        clearGuides();

        if (isRot) {
          const elRect = el.getBoundingClientRect();
          const cx = elRect.left + elRect.width / 2;
          const cy = elRect.top + elRect.height / 2;
          const deg = Math.round(Math.atan2(ev.clientY - cy, ev.clientX - cx) * (180 / Math.PI) + 90);
          el._g.rot = (deg + 360) % 360;
          el.style.transform = `rotate(${el._g.rot}deg)`;
          return;
        }

        const dxEmu = ((ev.clientX - startX) / (rect.width || 1)) * deck.W;
        const dyEmu = ((ev.clientY - startY) / (rect.height || 1)) * deck.H;

        if (dir) {
          const minS = 140000;
          const rad = ((orig.rot || 0) * Math.PI) / 180;
          const cos = Math.cos(rad), sin = Math.sin(rad);
          const ldx = dxEmu * cos + dyEmu * sin;
          const ldy = -dxEmu * sin + dyEmu * cos;

          let nw = orig.w, nh = orig.h;
          let localShiftX = 0, localShiftY = 0;

          if (dir.includes('e')) { nw = Math.max(minS, orig.w + ldx); localShiftX = (nw - orig.w) / 2; }
          if (dir.includes('w')) { nw = Math.max(minS, orig.w - ldx); localShiftX = -(nw - orig.w) / 2; }
          if (dir.includes('s')) { nh = Math.max(minS, orig.h + ldy); localShiftY = (nh - orig.h) / 2; }
          if (dir.includes('n')) { nh = Math.max(minS, orig.h - ldy); localShiftY = -(nh - orig.h) / 2; }

          const origCx = orig.x + orig.w / 2;
          const origCy = orig.y + orig.h / 2;
          const newCx = origCx + localShiftX * cos - localShiftY * sin;
          const newCy = origCy + localShiftX * sin + localShiftY * cos;

          el._g.x = Math.round(newCx - nw / 2);
          el._g.y = Math.round(newCy - nh / 2);
          el._g.w = Math.round(nw);
          el._g.h = Math.round(nh);
        } else {
          let nx = Math.round(orig.x + dxEmu);
          let ny = Math.round(orig.y + dyEmu);
          const snapX = deck.W * 0.011;
          const snapY = deck.H * 0.011;

          const myXPoints = [
            { val: nx, offset: 0 },
            { val: nx + orig.w / 2, offset: orig.w / 2 },
            { val: nx + orig.w, offset: orig.w }
          ];
          let bestDx = Infinity, snappedX = nx, guideXPct = null;
          for (const mp of myXPoints) {
            for (const tx of vTargets) {
              const diff = Math.abs(mp.val - tx);
              if (diff < snapX && diff < bestDx) {
                bestDx = diff;
                snappedX = Math.round(tx - mp.offset);
                guideXPct = (tx / deck.W) * 100;
              }
            }
          }
          if (guideXPct !== null) { nx = snappedX; showGuide('v', guideXPct); }

          const myYPoints = [
            { val: ny, offset: 0 },
            { val: ny + orig.h / 2, offset: orig.h / 2 },
            { val: ny + orig.h, offset: orig.h }
          ];
          let bestDy = Infinity, snappedY = ny, guideYPct = null;
          for (const mp of myYPoints) {
            for (const ty of hTargets) {
              const diff = Math.abs(mp.val - ty);
              if (diff < snapY && diff < bestDy) {
                bestDy = diff;
                snappedY = Math.round(ty - mp.offset);
                guideYPct = (ty / deck.H) * 100;
              }
            }
          }
          if (guideYPct !== null) { ny = snappedY; showGuide('h', guideYPct); }

          el._g.x = nx;
          el._g.y = ny;
        }
        placeEl(el, el._g, { ox: 0, oy: 0, w: deck.W, h: deck.H });
      };

      const onUp = (ev) => {
        if (ev.pointerId !== pointerId) return;
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        window.removeEventListener('pointercancel', onUp);
        clearGuides();
        if (!moved) return;
        writeGeomToShape(el._sh, el._g);
        pushHistory('transform');
        renderPptxEditor();
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
      window.addEventListener('pointercancel', onUp);
    });
  });

  stage.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    openContextMenu(e.clientX, e.clientY, [
      { label: '📥 لصق', hint: 'Ctrl+V', disabled: !curPptx.clipboardXml, action: () => pasteShapeToSlide(deck, slideIdx) },
      { label: '➕ مربع نص', action: () => addTextBoxToSlide(deck, slideIdx) },
      { label: '📐 شبكة التصميم', action: () => { curPptx.showGrid = !curPptx.showGrid; renderPptxEditor(); } }
    ]);
  });

  stage.addEventListener('pointerdown', (e) => {
    if (e.target.closest('.editing')) return;
    if (e.button !== undefined && e.button !== 0 && e.pointerType === 'mouse') return;
    const pointerId = e.pointerId;
    const rect = stage.getBoundingClientRect();
    const startX = e.clientX - rect.left;
    const startY = e.clientY - rect.top;
    let marqueeEl = null;
    let dragged = false;

    const onMove = (ev) => {
      if (ev.pointerId !== pointerId) return;
      const curX = Math.max(0, Math.min(rect.width, ev.clientX - rect.left));
      const curY = Math.max(0, Math.min(rect.height, ev.clientY - rect.top));
      const left = Math.min(startX, curX);
      const top = Math.min(startY, curY);
      const width = Math.abs(curX - startX);
      const height = Math.abs(curY - startY);

      if (!dragged && (width > 5 || height > 5)) {
        dragged = true;
        marqueeEl = h('div', { class: 'marquee-box' });
        stage.append(marqueeEl);
      }
      if (!marqueeEl) return;

      marqueeEl.style.left = left + 'px';
      marqueeEl.style.top = top + 'px';
      marqueeEl.style.width = width + 'px';
      marqueeEl.style.height = height + 'px';

      const selEmuX1 = (left / (rect.width || 1)) * deck.W;
      const selEmuY1 = (top / (rect.height || 1)) * deck.H;
      const selEmuX2 = ((left + width) / (rect.width || 1)) * deck.W;
      const selEmuY2 = ((top + height) / (rect.height || 1)) * deck.H;

      const hitIndices = [];
      shapeEls.forEach(s => {
        const si = +(s.dataset.i || -1);
        if (si < 0 || !s._g || s.style.display === 'none') return;
        const g = s._g;
        const intersects = !(g.x + g.w < selEmuX1 || g.x > selEmuX2 || g.y + g.h < selEmuY1 || g.y > selEmuY2);
        if (intersects) hitIndices.push(si);
      });

      curPptx.multiSel = new Set(hitIndices);
      curPptx.selShapeIdx = hitIndices.length ? hitIndices[hitIndices.length - 1] : null;
      shapeEls.forEach(s => {
        const si = +(s.dataset.i || -1);
        s.classList.toggle('sel', si === curPptx.selShapeIdx);
        s.classList.toggle('multi-sel', curPptx.multiSel.has(si));
      });
    };

    const onUp = (ev) => {
      if (ev.pointerId !== pointerId) return;
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      if (marqueeEl) marqueeEl.remove();

      if (dragged) {
        if (curPptx.multiSel.size === 1) curPptx.multiSel.clear();
        curPptx.selCell = null;
        renderPptxEditor();
      } else if (curPptx.selShapeIdx != null || curPptx.multiSel.size || curPptx.selCell) {
        curPptx.selShapeIdx = null;
        curPptx.multiSel.clear();
        curPptx.selCell = null;
        renderPptxEditor();
      }
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  });
}

/* ── لوحة الخصائص الجانبية ── */
function writeGeomToShape(sh, g) {
  const doc = sh.ownerDocument;
  let x = xfrmOf(sh);
  if (!x) {
    const spPr = kid(sh, P, 'spPr') || sh.appendChild(mk(doc, P, 'p:spPr'));
    x = mk(doc, A, 'a:xfrm');
    spPr.insertBefore(x, spPr.firstChild);
  }
  if (g.rot != null) x.setAttribute('rot', String(Math.round((g.rot % 360) * 60000)));
  if (g.flipH != null) { if (g.flipH) x.setAttribute('flipH', '1'); else x.removeAttribute('flipH'); }
  if (g.flipV != null) { if (g.flipV) x.setAttribute('flipV', '1'); else x.removeAttribute('flipV'); }
  let off = kid(x, A, 'off'), ext = kid(x, A, 'ext');
  if (!off) { off = mk(doc, A, 'a:off', { x: '0', y: '0' }); x.insertBefore(off, x.firstChild); }
  if (!ext) { ext = mk(doc, A, 'a:ext', { cx: '914400', cy: '914400' }); insertAfter(off, ext); }
  off.setAttribute('x', String(Math.round(g.x)));
  off.setAttribute('y', String(Math.round(g.y)));
  ext.setAttribute('cx', String(Math.max(1, Math.round(g.w))));
  ext.setAttribute('cy', String(Math.max(1, Math.round(g.h))));
}

function buildGradFillXml(doc, hex1, hex2, angleDeg, alphaPct) {
  const gf = mk(doc, A, 'a:gradFill', { flip: 'none', rotWithShape: '1' });
  const gsLst = mk(doc, A, 'a:gsLst');

  const gs1 = mk(doc, A, 'a:gs', { pos: '0' });
  const c1 = mk(doc, A, 'a:srgbClr', { val: hex1.replace('#', '').toUpperCase() });
  if (alphaPct != null && alphaPct < 100) c1.appendChild(mk(doc, A, 'a:alpha', { val: String(alphaPct * 1000) }));
  gs1.appendChild(c1);

  const gs2 = mk(doc, A, 'a:gs', { pos: '100000' });
  const c2 = mk(doc, A, 'a:srgbClr', { val: hex2.replace('#', '').toUpperCase() });
  if (alphaPct != null && alphaPct < 100) c2.appendChild(mk(doc, A, 'a:alpha', { val: String(alphaPct * 1000) }));
  gs2.appendChild(c2);

  gsLst.append(gs1, gs2);
  gf.appendChild(gsLst);

  const ooxmlDeg = (((+angleDeg || 0) - 90) % 360 + 360) % 360;
  const angVal = Math.round(ooxmlDeg * 60000);
  gf.appendChild(mk(doc, A, 'a:lin', { ang: String(angVal), scaled: '1' }));
  return gf;
}

function applyTextFormatToTableCell(tc, opts) {
  if (!tc) return;
  const doc = tc.ownerDocument;
  let txb = txBodyOf(tc);
  if (!txb) {
    txb = mk(doc, A, 'a:txBody');
    txb.append(mk(doc, A, 'a:bodyPr'), mk(doc, A, 'a:lstStyle'), mk(doc, A, 'a:p'));
    tc.appendChild(txb);
  }
  const sz = Math.max(6, Math.min(300, +opts.sz || 16)) * 100;
  for (const p of kids(txb, A, 'p')) {
    let pp = kid(p, A, 'pPr');
    if (!pp) { pp = mk(doc, A, 'a:pPr'); p.insertBefore(pp, p.firstChild); }
    pp.setAttribute('algn', opts.algn);

    let runs = kids(p, A, 'r');
    if (!runs.length) {
      const r = mk(doc, A, 'a:r');
      r.appendChild(mk(doc, A, 'a:t'));
      p.appendChild(r);
      runs = [r];
    }
    for (const r of runs) {
      let rp = kid(r, A, 'rPr');
      if (!rp) { rp = mk(doc, A, 'a:rPr'); r.insertBefore(rp, r.firstChild); }
      rp.setAttribute('sz', String(sz));
      rp.setAttribute('b', opts.bold ? '1' : '0');
      rp.setAttribute('i', opts.italic ? '1' : '0');
      for (const f of kids(rp, A)) if (f.localName === 'solidFill') removeNode(f);
      const sf = mk(doc, A, 'a:solidFill');
      sf.appendChild(mk(doc, A, 'a:srgbClr', { val: opts.color.replace('#', '').toUpperCase() }));
      rp.appendChild(sf);
    }
  }
}

async function applyPptxImageFilter(slidePath, shape, imageEl, filter) {
  const blip = desc(shape, A, 'blip')[0];
  if (!blip) return;
  const cNvPr = kid(nvOf(shape), P, 'cNvPr');
  const originalMarker = cNvPr?.getAttribute('descr') || '';
  const markerPrefix = 'docstudio-original-image:';
  const rels = readRels(pk, slidePath);
  const currentRel = rels.get(blip.getAttributeNS(R, 'embed'));
  if (!currentRel) return;
  if (filter === 'none') {
    const originalPath = originalMarker.startsWith(markerPrefix) ? originalMarker.slice(markerPrefix.length) : '';
    if (originalPath && pk.has(originalPath)) {
      const newRid = addRel(pk, slidePath, REL + 'image', originalPath);
      const oldRid = blip.getAttributeNS(R, 'embed');
      blip.setAttributeNS(R, 'r:embed', newRid);
      cNvPr?.removeAttribute('descr');
      removeRelIfUnused(pk, slidePath, oldRid, pk.doc(slidePath));
      pushHistory('imageFilterReset'); renderPptxEditor();
    }
    return;
  }
  if (!imageEl || !imageEl.complete || !imageEl.naturalWidth) { toast(currentLanguage === 'en' ? 'Image is not ready yet' : 'الصورة غير جاهزة'); return; }
  try {
    const originalPath = originalMarker.startsWith(markerPrefix) ? originalMarker.slice(markerPrefix.length) : currentRel.path;
    const canvas = document.createElement('canvas');
    canvas.width = imageEl.naturalWidth; canvas.height = imageEl.naturalHeight;
    const ctx = canvas.getContext('2d'); ctx.filter = filter; ctx.drawImage(imageEl, 0, 0);
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    if (!blob) throw new Error('Image conversion failed');
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const mediaPath = uniquePath(pk, 'ppt/media', 'filtered', 'png');
    pk.setBin(mediaPath, bytes); ensureDefault(pk, 'png', 'image/png');
    const newRid = addRel(pk, slidePath, REL + 'image', mediaPath);
    const oldRid = blip.getAttributeNS(R, 'embed');
    blip.setAttributeNS(R, 'r:embed', newRid);
    if (cNvPr && !originalMarker.startsWith(markerPrefix)) cNvPr.setAttribute('descr', markerPrefix + originalPath);
    removeRelIfUnused(pk, slidePath, oldRid, pk.doc(slidePath));
    pushHistory('imageFilter'); renderPptxEditor();
  } catch (err) {
    console.error(err); toast(currentLanguage === 'en' ? 'Could not apply image filter' : 'تعذر تطبيق فلتر الصورة');
  }
}

function renderShapeProps(panel, stage, deck, slideIdx) {
  panel.innerHTML = '';
  panel.append(h('h3', null, tr('خصائص العنصر')));

  if (curPptx.multiSel.size >= 2) {
    panel.append(
      h('h4', null, tr('محاذاة العناصر المحددة')),
      h('div', { class: 'align-grid' },
        btn('يسار', () => alignSelectedShapes(deck, slideIdx, 'left')),
        btn('توسيط أفقي', () => alignSelectedShapes(deck, slideIdx, 'center')),
        btn('يمين', () => alignSelectedShapes(deck, slideIdx, 'right')),
        btn('أعلى', () => alignSelectedShapes(deck, slideIdx, 'top')),
        btn('توسيط عمودي', () => alignSelectedShapes(deck, slideIdx, 'middle')),
        btn('أسفل', () => alignSelectedShapes(deck, slideIdx, 'bottom')),
        btn('توزيع أفقي', () => distributeSelectedShapes(deck, slideIdx, 'h')),
        btn('توزيع عمودي', () => distributeSelectedShapes(deck, slideIdx, 'v'))
      ),
      h('hr')
    );
  }

  if (curPptx.selShapeIdx == null) {
    panel.append(h('p', { class: 'hint' }, tr('انقر مرتين على أي نص للتعديل المباشر، أو حدد عنصراً لتعديل خصائصه.')));
    return;
  }

  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const sh = shapesOf(doc)[curPptx.selShapeIdx];
  if (!sh) { panel.append(h('p', { class: 'hint' }, tr('العنصر غير متاح.'))); return; }

  const ctx = (stage && stage._ctx) || buildCtx(deck, slPath);
  const g = geomOf(sh, ctx) || { x: 0, y: 0, w: deck.W * 0.25, h: deck.H * 0.15, rot: 0 };
  const addField = (label, input) => panel.append(h('div', { class: 'fld' }, h('label', null, tr(label)), input));

  if (sh.localName === 'sp') {
    const txtArea = h('textarea', {
      placeholder: tr('اكتب النص هنا…'),
      value: getShapePlainText(sh)
    });
    const saveTxt = () => {
      let txb = txBodyOf(sh);
      if (!txb) {
        txb = mk(doc, P, 'p:txBody');
        txb.append(
          mk(doc, A, 'a:bodyPr', { anchor: 'ctr' }),
          mk(doc, A, 'a:lstStyle'),
          frag(doc, `<a:p xmlns:a="${A}"><a:pPr algn="ctr"/><a:r><a:rPr sz="2000"><a:solidFill><a:srgbClr val="1F1C17"/></a:solidFill></a:rPr><a:t></a:t></a:r></a:p>`)
        );
        sh.appendChild(txb);
      }
      setShapePlainText(sh, txtArea.value);
      pushHistory('editShapeText');
      renderPptxEditor();
      toast('تم تحديث النص');
    };
    panel.append(
      addField('محتوى النص', txtArea),
      btn('حفظ النص', saveTxt, 'pri'),
      h('hr')
    );
  }

  const xIn = h('input', { type: 'number', value: (g.x / 9525).toFixed(0), step: 1 });
  const yIn = h('input', { type: 'number', value: (g.y / 9525).toFixed(0), step: 1 });
  const wIn = h('input', { type: 'number', value: (g.w / 9525).toFixed(0), step: 1 });
  const hIn = h('input', { type: 'number', value: (g.h / 9525).toFixed(0), step: 1 });
  const rIn = h('input', { type: 'number', value: Math.round(g.rot || 0), min: 0, max: 360, step: 1 });

  const applyGeom = () => {
    const next = {
      ...g,
      x: Math.round(+xIn.value * 9525),
      y: Math.round(+yIn.value * 9525),
      w: Math.max(140000, Math.round(+wIn.value * 9525)),
      h: Math.max(140000, Math.round(+hIn.value * 9525)),
      rot: (+rIn.value || 0) % 360
    };
    writeGeomToShape(sh, next);
    pushHistory('geometry');
    renderPptxEditor();
    toast('تم تطبيق الأبعاد');
  };

  panel.append(
    h('div', { class: 'fld-grid' },
      h('div', { class: 'fld' }, h('label', null, 'X (px)'), xIn),
      h('div', { class: 'fld' }, h('label', null, 'Y (px)'), yIn),
      h('div', { class: 'fld' }, h('label', null, tr('العرض (px)')), wIn),
      h('div', { class: 'fld' }, h('label', null, tr('الارتفاع (px)')), hIn)
    ),
    addField('التدوير (°)', rIn),
    h('div', { class: 'row' },
      btn('تطبيق الأبعاد', applyGeom, 'pri'),
      btn('🔒 قفل/فتح العنصر', () => toggleLockShape(deck, slideIdx), isShapeLocked(sh) ? 'on' : '')
    )
  );

  /* محاذاة العنصر الواحد مع الشريحة + حركة الدخول (Animations) */
  const cNvPrAnim = kid(nvOf(sh), P, 'cNvPr');
  if (cNvPrAnim) {
    const curAnim = cNvPrAnim.getAttribute('data-anim') || 'none';
    const animSel = h('select', null,
      h('option', { value: 'none' }, tr('بدون حركة دخول')),
      h('option', { value: 'fade' }, tr('تلاشي ناعم (Fade)')),
      h('option', { value: 'fly' }, tr('انزلاق للأعلى (Fly In)')),
      h('option', { value: 'zoom' }, tr('تكبير (Zoom In)'))
    );
    animSel.value = curAnim;
    animSel.onchange = () => {
      if (animSel.value === 'none') cNvPrAnim.removeAttribute('data-anim');
      else cNvPrAnim.setAttribute('data-anim', animSel.value);
      pushHistory('shapeAnim');
      const domEl = $(`.stagewrap .shape[data-i="${curPptx.selShapeIdx}"]`);
      if (domEl && animSel.value !== 'none') {
        domEl.classList.remove('anim-preview-fade', 'anim-preview-fly', 'anim-preview-zoom');
        void domEl.offsetWidth;
        domEl.classList.add('anim-preview-' + animSel.value);
      }
    };
    panel.append(
      h('hr'),
      h('h4', null, tr('محاذاة مع الشريحة وحركة الدخول')),
      h('div', { class: 'row' },
        btn('↔ توسيط أفقي للشريحة', () => centerSingleShapeToSlide('h')),
        btn('↕ توسيط عمودي للشريحة', () => centerSingleShapeToSlide('v'))
      ),
      addField('حركة ظهور العنصر في العرض', animSel)
    );
  }

  if (sh.localName === 'pic') {
    const domImg = $(`.stagewrap .shape[data-i="${curPptx.selShapeIdx}"] img`);
    panel.append(
      h('hr'),
      h('h4', null, tr('أدوات الصورة')),
      h('div', { class: 'row' },
        btn('🔄 تبديل الصورة', () => triggerReplaceImage((info) => replaceSelectedImage(slPath, sh, info)), 'pri'),
        btn('↔ قلب أفقي', () => { writeGeomToShape(sh, { ...g, flipH: !g.flipH }); pushHistory('flipH'); renderPptxEditor(); }),
        btn('↕ قلب عمودي', () => { writeGeomToShape(sh, { ...g, flipV: !g.flipV }); pushHistory('flipV'); renderPptxEditor(); }),
        btn('⭕ قص دائري', () => { setPicCropPreset(sh, 'ellipse'); pushHistory('cropCircle'); renderPptxEditor(); }),
        btn('⬜ إعادة تعيين القص', () => { setPicCropPreset(sh, 'rect'); pushHistory('cropRect'); renderPptxEditor(); })
      ),
      h('h4', null, tr('تأثيرات وفلاتر الصورة')),
      h('div', { class: 'row' },
        btn('أبيض وأسود', () => applyPptxImageFilter(slPath, sh, domImg, 'grayscale(100%)')),
        btn('تباين عالي', () => applyPptxImageFilter(slPath, sh, domImg, 'contrast(135%) saturate(120%)')),
        btn('إضاءة ناعمة', () => applyPptxImageFilter(slPath, sh, domImg, 'brightness(115%)')),
        btn('أصلي', () => applyPptxImageFilter(slPath, sh, domImg, 'none'))
      )
    );
  }

  const txb = txBodyOf(sh);
  if (txb) {
    panel.append(h('hr'), h('h4', null, tr('تنسيق النص')));
    const firstR = desc(txb, A, 'rPr')[0];
    const firstP = kid(txb, A, 'p');
    const firstPPr = firstP && kid(firstP, A, 'pPr');
    const resolvedRp = runProps(ctx, sh, firstR, +(firstPPr?.getAttribute('lvl') || 0));

    const size = h('input', { type: 'number', min: 6, max: 300, value: Math.round((resolvedRp.sz || 2400) / 100) });
    const fontList = [
      'Cairo', 'Tajawal', 'Noto Sans Arabic', 'IBM Plex Sans Arabic', 'Amiri', 'Traditional Arabic', 'Sakkal Majalla',
      'Tahoma', 'Segoe UI', 'Inter', 'Calibri', 'Arial', 'Verdana', 'Trebuchet MS',
      'Times New Roman', 'Georgia', 'Cambria', 'Garamond', 'Courier New', 'Consolas'
    ];
    if (resolvedRp.font && !resolvedRp.font.startsWith('+') && !fontList.includes(resolvedRp.font)) {
      fontList.unshift(resolvedRp.font);
    }
    const font = h('select', null, ...fontList.map(n => h('option', { value: n }, n)));
    if (resolvedRp.font && !resolvedRp.font.startsWith('+')) font.value = resolvedRp.font;

    const align = h('select', null, ...[['l', 'يسار'], ['ctr', 'وسط'], ['r', 'يمين'], ['just', 'ضبط']].map(([v, n]) => h('option', { value: v }, tr(n))));
    if (firstPPr && firstPPr.getAttribute('algn')) align.value = firstPPr.getAttribute('algn');

    const lineSpacing = h('select', null,
      h('option', { value: '100000' }, '1.0'),
      h('option', { value: '115000' }, '1.15'),
      h('option', { value: '150000' }, '1.5'),
      h('option', { value: '200000' }, '2.0')
    );
    const curLnSpc = firstPPr && kid(firstPPr, A, 'lnSpc') && kid(kid(firstPPr, A, 'lnSpc'), A, 'spcPct');
    if (curLnSpc) {
      const v = curLnSpc.getAttribute('val');
      if (['100000', '115000', '150000', '200000'].includes(v)) lineSpacing.value = v;
    }

    const listMode = h('select', null,
      h('option', { value: 'none' }, tr('بدون تعداد')),
      h('option', { value: 'bullet' }, tr('تعداد نقطي')),
      h('option', { value: 'number' }, tr('تعداد رقمي'))
    );
    if (firstPPr) {
      if (kid(firstPPr, A, 'buChar')) listMode.value = 'bullet';
      else if (kid(firstPPr, A, 'buAutoNum')) listMode.value = 'number';
    }

    const dirMode = h('select', null,
      h('option', { value: 'rtl' }, tr('اتجاه RTL')),
      h('option', { value: 'ltr' }, tr('اتجاه LTR'))
    );
    if (firstPPr && firstPPr.getAttribute('rtl') === '0') dirMode.value = 'ltr';

    const curTextHex = resolvedRp.color ? ('#' + resolvedRp.color.hex) : '#1F1C17';
    const color = h('input', { type: 'color', value: curTextHex });
    const bold = h('input', { type: 'checkbox', checked: resolvedRp.b });
    const italic = h('input', { type: 'checkbox', checked: resolvedRp.i });
    const under = h('input', { type: 'checkbox', checked: resolvedRp.u });

    panel.append(
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('حجم الخط')), size),
        h('div', { class: 'fld' }, h('label', null, tr('نوع الخط')), font)
      ),
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('محاذاة')), align),
        h('div', { class: 'fld' }, h('label', null, tr('اتجاه النص')), dirMode)
      ),
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('التعداد')), listMode),
        h('div', { class: 'fld' }, h('label', null, tr('تباعد الأسطر')), lineSpacing)
      ),
      h('div', { class: 'row format-checks' },
        h('label', null, bold, ' ', tr('عريض')),
        h('label', null, italic, ' ', tr('مائل')),
        h('label', null, under, ' ', tr('تحته خط'))
      ),
      addField('لون الخط', makeColorWithEyedropper(color)),
      btn('تطبيق تنسيق النص', () => {
        const sz = Math.max(6, Math.min(300, +size.value || 24)) * 100;
        for (const r of desc(txb, A, 'r')) {
          let rp = kid(r, A, 'rPr');
          if (!rp) { rp = mk(doc, A, 'a:rPr'); r.insertBefore(rp, r.firstChild); }
          rp.setAttribute('sz', String(sz));
          rp.setAttribute('b', bold.checked ? '1' : '0');
          rp.setAttribute('i', italic.checked ? '1' : '0');
          rp.setAttribute('u', under.checked ? 'sng' : 'none');
          for (const f of kids(rp, A)) if (['solidFill', 'latin', 'ea', 'cs'].includes(f.localName)) removeNode(f);
          const sf = mk(doc, A, 'a:solidFill');
          sf.appendChild(mk(doc, A, 'a:srgbClr', { val: color.value.slice(1).toUpperCase() }));
          rp.appendChild(sf);
          rp.appendChild(mk(doc, A, 'a:latin', { typeface: font.value }));
          rp.appendChild(mk(doc, A, 'a:ea', { typeface: font.value }));
          rp.appendChild(mk(doc, A, 'a:cs', { typeface: font.value }));
        }
        for (const p of kids(txb, A, 'p')) {
          let pp = kid(p, A, 'pPr');
          if (!pp) { pp = mk(doc, A, 'a:pPr'); p.insertBefore(pp, p.firstChild); }
          pp.setAttribute('algn', align.value);
          pp.setAttribute('rtl', dirMode.value === 'rtl' ? '1' : '0');

          const oldLnSpc = kid(pp, A, 'lnSpc');
          if (oldLnSpc) removeNode(oldLnSpc);
          const lnSpcEl = mk(doc, A, 'a:lnSpc');
          lnSpcEl.appendChild(mk(doc, A, 'a:spcPct', { val: lineSpacing.value }));
          pp.appendChild(lnSpcEl);

          for (const b of kids(pp, A)) if (/^bu/.test(b.localName)) removeNode(b);
          if (listMode.value === 'bullet') {
            pp.setAttribute('marL', '285750'); pp.setAttribute('indent', '-285750');
            pp.appendChild(mk(doc, A, 'a:buChar', { char: '•' }));
          } else if (listMode.value === 'number') {
            pp.setAttribute('marL', '285750'); pp.setAttribute('indent', '-285750');
            pp.appendChild(mk(doc, A, 'a:buAutoNum', { type: 'arabicPeriod', startAt: '1' }));
          } else {
            pp.removeAttribute('marL'); pp.removeAttribute('indent');
            pp.appendChild(mk(doc, A, 'a:buNone'));
          }
        }
        pushHistory('formatText');
        renderPptxEditor();
        toast('تم تطبيق تنسيق النص');
      }, 'pri')
    );
  }

  if (sh.localName === 'sp' || sh.localName === 'cxnSp' || sh.localName === 'pic') {
    panel.append(h('hr'), h('h4', null, tr('المظهر والحدود')));
    const appState = readShapeAppearance(sh, ctx);

    const fillType = h('select', null,
      h('option', { value: 'solid' }, tr('لون ثابت')),
      h('option', { value: 'grad' }, tr('تدرج لوني'))
    );
    fillType.value = appState.fillType;

    const fill = h('input', { type: 'color', value: appState.fillHex });
    const gradCol2 = h('input', { type: 'color', value: appState.gradHex2 });
    const gradAngle = h('input', { type: 'number', min: 0, max: 360, value: appState.gradAngle });
    const noFill = h('input', { type: 'checkbox', checked: appState.noFill });

    const gradControls = h('div', { class: 'fld-grid', style: { marginTop: '6px' }, hidden: appState.fillType !== 'grad' },
      h('div', { class: 'fld' }, h('label', null, tr('لون 2')), makeColorWithEyedropper(gradCol2)),
      h('div', { class: 'fld' }, h('label', null, tr('الزاوية (°)')), gradAngle)
    );
    fillType.onchange = () => {
      gradControls.hidden = fillType.value !== 'grad';
      noFill.checked = false;
    };
    fill.oninput = () => { noFill.checked = false; };

    const borderCol = h('input', { type: 'color', value: appState.borderHex });
    const borderW = h('input', { type: 'number', min: 0, max: 30, value: appState.borderW });
    const opacity = h('input', { type: 'number', min: 10, max: 100, value: appState.opacity });
    const shadow = h('input', { type: 'checkbox', checked: appState.hasShadow });
    const prstGeomEl = kid(kid(sh, P, 'spPr'), A, 'prstGeom');
    const isRoundRect = prstGeomEl?.getAttribute('prst') === 'roundRect';
    const curGd = isRoundRect ? kid(kid(prstGeomEl, A, 'avLst'), A, 'gd') : null;
    const curRadM = curGd ? /val\s+(\d+)/.exec(curGd.getAttribute('fmla') || '') : null;
    const cornerRad = h('input', { type: 'number', min: 0, max: 50, value: curRadM ? Math.round((+curRadM[1] / 50000) * 50) : 12 });

    panel.append(
      sh.localName !== 'pic' ? h('div', null,
        addField('نوع التعبئة:', fillType),
        h('div', { class: 'row' },
          h('label', null, makeColorWithEyedropper(fill, () => { noFill.checked = false; }), ' ', tr('لون 1')),
          h('label', null, noFill, ' ', tr('بدون تعبئة'))
        ),
        gradControls
      ) : null,
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('لون الحدود')), makeColorWithEyedropper(borderCol)),
        h('div', { class: 'fld' }, h('label', null, tr('سمك الحدود (px)')), borderW)
      ),
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('الشفافية (%)')), opacity),
        h('div', { class: 'fld' }, h('label', null, tr('تأثير')), h('label', { style: { marginTop: '6px', display: 'flex', gap: '5px' } }, shadow, ' ', tr('ظل خفيف')))
      ),
      isRoundRect ? addField('استدارة الزوايا (%)', cornerRad) : null,
      btn('تطبيق المظهر', () => {
        const spPr = kid(sh, P, 'spPr');
        if (!spPr) return;

        if (sh.localName !== 'pic') {
          for (const f of kids(spPr, A)) if (['solidFill', 'noFill', 'gradFill'].includes(f.localName)) removeNode(f);
          const opVal = Math.max(10, Math.min(100, +opacity.value || 100));
          let fEl;
          if (noFill.checked) {
            fEl = mk(doc, A, 'a:noFill');
          } else if (fillType.value === 'grad') {
            fEl = buildGradFillXml(doc, fill.value, gradCol2.value, +gradAngle.value || 0, opVal);
          } else {
            fEl = mk(doc, A, 'a:solidFill');
            const srgb = mk(doc, A, 'a:srgbClr', { val: fill.value.slice(1).toUpperCase() });
            if (opVal < 100) srgb.appendChild(mk(doc, A, 'a:alpha', { val: String(opVal * 1000) }));
            fEl.appendChild(srgb);
          }
          const geom = kid(spPr, A, 'prstGeom');
          spPr.insertBefore(fEl, geom || null);
        }

        const oldLn = kid(spPr, A, 'ln');
        if (oldLn) removeNode(oldLn);
        const bw = +borderW.value || 0;
        const lnEl = mk(doc, A, 'a:ln', { w: String(Math.round(bw * 12700)) });
        if (bw > 0) {
          const sf = mk(doc, A, 'a:solidFill');
          sf.appendChild(mk(doc, A, 'a:srgbClr', { val: borderCol.value.slice(1).toUpperCase() }));
          lnEl.appendChild(sf);
        } else lnEl.appendChild(mk(doc, A, 'a:noFill'));
        spPr.appendChild(lnEl);

        if (isRoundRect && prstGeomEl) {
          let avLst = kid(prstGeomEl, A, 'avLst');
          if (!avLst) { avLst = mk(doc, A, 'a:avLst'); prstGeomEl.appendChild(avLst); }
          avLst.innerHTML = '';
          const adjVal = Math.round((Math.max(0, Math.min(50, +cornerRad.value || 0)) / 50) * 50000);
          avLst.appendChild(mk(doc, A, 'a:gd', { name: 'adj', fmla: `val ${adjVal}` }));
        }
        const oldEff = kid(spPr, A, 'effectLst');
        if (oldEff) removeNode(oldEff);
        if (shadow.checked) {
          const eff = mk(doc, A, 'a:effectLst');
          const shd = mk(doc, A, 'a:outerShdw', { blurRad: '50800', dist: '38100', dir: '2700000' });
          const sc = mk(doc, A, 'a:srgbClr', { val: '000000' });
          sc.appendChild(mk(doc, A, 'a:alpha', { val: '35000' }));
          shd.appendChild(sc);
          eff.appendChild(shd);
          spPr.appendChild(eff);
        }

        pushHistory('appearance');
        renderPptxEditor();
        toast('تم تطبيق المظهر');
      }, 'pri')
    );
  }

  if (sh.localName === 'graphicFrame' && desc(sh, A, 'tbl').length) {
    const tbl = desc(sh, A, 'tbl')[0];
    panel.append(h('hr'), h('h4', null, tr('إدارة الجدول')));
    const cellColor = h('input', { type: 'color', value: '#D9E1F2' });

    const getSelectedTc = () => {
      if (!curPptx.selCell) return null;
      const trRow = kids(tbl, A, 'tr')[curPptx.selCell.r];
      return trRow && kids(trRow, A, 'tc')[curPptx.selCell.c];
    };

    const selTc = getSelectedTc();
    const selRPr = selTc && desc(txBodyOf(selTc), A, 'rPr')[0];
    const selPPr = selTc && kid(kid(txBodyOf(selTc), A, 'p'), A, 'pPr');

    const cFontSize = h('input', { type: 'number', min: 6, max: 200, value: selRPr && selRPr.getAttribute('sz') ? Math.round(+selRPr.getAttribute('sz') / 100) : 16 });
    const cAlign = h('select', null, ...[['l', 'يسار'], ['ctr', 'وسط'], ['r', 'يمين']].map(([v, n]) => h('option', { value: v }, tr(n))));
    if (selPPr && selPPr.getAttribute('algn')) cAlign.value = selPPr.getAttribute('algn');
    const cTextColor = h('input', { type: 'color', value: '#1F1C17' });
    const cBold = h('input', { type: 'checkbox', checked: selRPr?.getAttribute('b') === '1' });
    const cItalic = h('input', { type: 'checkbox', checked: selRPr?.getAttribute('i') === '1' });

    panel.append(
      h('div', { class: 'row' },
        btn('➕ صف', () => modifyTableStructure(tbl, 'addRow')),
        btn('➖ صف', () => modifyTableStructure(tbl, 'delRow'), 'danger'),
        btn('➕ عمود', () => modifyTableStructure(tbl, 'addCol')),
        btn('➖ عمود', () => modifyTableStructure(tbl, 'delCol'), 'danger')
      ),
      addField('لون الخلية المحددة', makeColorWithEyedropper(cellColor)),
      btn('تطبيق لون الخلية', () => {
        const tc = getSelectedTc();
        if (!tc) { toast('انقر على خلية داخل الجدول أولاً'); return; }
        let tcPr = kid(tc, A, 'tcPr');
        if (!tcPr) { tcPr = mk(doc, A, 'a:tcPr'); tc.insertBefore(tcPr, tc.firstChild); }
        for (const f of kids(tcPr, A)) if (['solidFill', 'noFill'].includes(f.localName)) removeNode(f);
        const sf = mk(doc, A, 'a:solidFill');
        sf.appendChild(mk(doc, A, 'a:srgbClr', { val: cellColor.value.slice(1).toUpperCase() }));
        tcPr.appendChild(sf);
        pushHistory('cellFill');
        renderPptxEditor();
        toast('تم تلوين الخلية');
      }),
      h('hr'),
      h('h4', null, tr('تنسيق نص الخلية المحددة')),
      h('div', { class: 'fld-grid' },
        h('div', { class: 'fld' }, h('label', null, tr('حجم الخط')), cFontSize),
        h('div', { class: 'fld' }, h('label', null, tr('محاذاة')), cAlign)
      ),
      h('div', { class: 'row format-checks' },
        h('label', null, cBold, ' ', tr('عريض')),
        h('label', null, cItalic, ' ', tr('مائل')),
        h('label', null, makeColorWithEyedropper(cTextColor), ' ', tr('لون الخط'))
      ),
      btn('تطبيق على الخلية', () => {
        const tc = getSelectedTc();
        if (!tc) { toast('انقر على خلية داخل الجدول أولاً'); return; }
        applyTextFormatToTableCell(tc, {
          sz: +cFontSize.value || 16,
          algn: cAlign.value,
          bold: cBold.checked,
          italic: cItalic.checked,
          color: cTextColor.value
        });
        pushHistory('cellTextFmt');
        renderPptxEditor();
        toast('تم تطبيق التنسيق على نص الخلية');
      }, 'pri')
    );
  }

  if (sh.localName === 'graphicFrame' && desc(sh, C, 'chart').length) {
    const cEl = desc(sh, C, 'chart')[0];
    const rel = readRels(pk, slPath).get(cEl.getAttributeNS(R, 'id'));
    const cData = rel && parseChartData(rel.path);
    if (cData) {
      panel.append(h('hr'), h('h4', null, tr('بيانات الرسم البياني')));
      const typeSel = h('select', null,
        h('option', { value: 'bar' }, tr('أعمدة')),
        h('option', { value: 'line' }, tr('خطي')),
        h('option', { value: 'pie' }, tr('دائري'))
      );
      typeSel.value = cData.type;
      panel.append(addField('نوع الرسم', typeSel));

      let hasSer2 = cData.series.length > 1;
      const s1Name = cData.series[0]?.name || tr('السلسلة 1');
      const s2Name = cData.series[1]?.name || tr('السلسلة 2');
      const rowsState = cData.cats.map((cat, i) => ({
        cat,
        v1: cData.series[0]?.vals[i] ?? 0,
        v2: cData.series[1]?.vals[i] ?? 40
      }));
      const tbl = h('table', { class: 'chart-data-table' });

      const rebuildChartTable = () => {
        tbl.innerHTML = '';
        tbl.append(h('tr', null,
          h('th', null, tr('التسمية')),
          h('th', null, s1Name),
          hasSer2 ? h('th', null, s2Name) : null,
          h('th', { style: { width: '26px' } }, '')
        ));
        rowsState.forEach((row, idx) => {
          const cIn = h('input', { type: 'text', value: row.cat, oninput: () => { row.cat = cIn.value; } });
          const v1In = h('input', { type: 'number', value: row.v1, oninput: () => { row.v1 = +v1In.value || 0; } });
          const v2In = hasSer2 ? h('input', { type: 'number', value: row.v2, oninput: () => { row.v2 = +v2In.value || 0; } }) : null;
          const delBtn = h('button', {
            type: 'button',
            class: 'chart-del-btn',
            title: tr('حذف'),
            disabled: rowsState.length <= 1,
            onclick: () => {
              if (rowsState.length <= 1) return;
              rowsState.splice(idx, 1);
              rebuildChartTable();
            }
          }, '🗑');
          tbl.append(h('tr', null, h('td', null, cIn), h('td', null, v1In), hasSer2 ? h('td', null, v2In) : null, h('td', null, delBtn)));
        });
      };
      rebuildChartTable();

      const addRowBtn = btn('➕ إضافة صف/بيان', () => {
        const nextNum = rowsState.length + 1;
        rowsState.push({ cat: (currentLanguage === 'en' ? 'Item ' : 'بيان ') + nextNum, v1: 50, v2: 65 });
        rebuildChartTable();
      });

      const toggleSer2Btn = btn('➕ سلسلة مقارنة', () => {
        hasSer2 = !hasSer2;
        rebuildChartTable();
      }, hasSer2 ? 'on' : '');

      panel.append(
        tbl,
        h('div', { class: 'row' },
          addRowBtn,
          typeSel.value !== 'pie' ? toggleSer2Btn : null,
          btn('تحديث الرسم', () => {
            const seriesPayload = [{ name: s1Name, vals: rowsState.map(x => +x.v1 || 0) }];
            if (hasSer2 && typeSel.value !== 'pie') {
              seriesPayload.push({ name: s2Name, vals: rowsState.map(x => +x.v2 || 0) });
            }
            updateChartXml(cData.chartPath, typeSel.value, rowsState.map(x => x.cat), seriesPayload);
            pushHistory('editChart');
            renderPptxEditor();
            toast('تم تحديث الرسم البياني');
          }, 'pri')
        )
      );
    }
  }

  const cNvPr = kid(nvOf(sh), P, 'cNvPr');
  if (cNvPr) {
    panel.append(h('hr'), h('h4', null, tr('رابط تشعبي (Hyperlink)')));
    const existingLink = kid(cNvPr, A, 'hlinkClick');
    const curUrl = existingLink ? (readRels(pk, slPath).get(existingLink.getAttributeNS(R, 'id'))?.target || '') : '';
    const linkIn = h('input', { type: 'text', placeholder: 'https://example.com', value: curUrl });

    panel.append(
      linkIn,
      h('div', { class: 'row' },
        btn('حفظ الرابط', () => {
          const url = safeUrl(linkIn.value);
          if (!linkIn.value.trim() || url === '#') { toast(currentLanguage === 'en' ? 'Enter a valid web or email address' : 'أدخل رابطاً أو بريداً إلكترونياً صحيحاً'); return; }
          const old = kid(cNvPr, A, 'hlinkClick');
          if (old) { const oldId = old.getAttributeNS(R, 'id'); removeNode(old); removeRelIfUnused(pk, slPath, oldId, doc); }
          const rid = addRel(pk, slPath, REL + 'hyperlink', url, true);
          const hl = mk(doc, A, 'a:hlinkClick');
          hl.setAttributeNS(R, 'r:id', rid);
          cNvPr.appendChild(hl);
          pushHistory('hyperlink');
          renderPptxEditor();
          toast('تم ربط العنصر بالرابط');
        }, 'pri'),
        curUrl ? btn('إزالة الرابط', () => {
          const old = kid(cNvPr, A, 'hlinkClick');
          if (old) { const oldId = old.getAttributeNS(R, 'id'); removeNode(old); removeRelIfUnused(pk, slPath, oldId, doc); }
          pushHistory('removeLink');
          renderPptxEditor();
          toast('تمت إزالة الرابط');
        }, 'danger') : null
      )
    );
  }
}
