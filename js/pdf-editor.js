'use strict';
/* ═════════════════════════════════════════════════════════════
   محرر PDF المتكامل محلياً (PDF.js + pdf-lib)
   الجزء 1 من 2: التهيئة، المصغرات، الحذف، الترتيب، التدوير، والدمج
   ═════════════════════════════════════════════════════════════ */

let curPdf = {
  bytes: null,
  pageCount: 0,
  pageIdx: 0,
  rotations: [],
  zoom: 1.15,
  _effectiveZoom: 1.15,
  tool: 'none', // 'none' | 'edit-text' | 'text' | 'pen' | 'highlight' | 'whiteout'
  penColor: '#e11d48',
  penSize: 3,
  highlightColor: '#facc15',
  textColor: '#1f1c17',
  textSize: 18,
  drawings: {},  // { [pageIdx]: [{ pts: [{ x, y }], color, size }] }
  texts: {},     // { [pageIdx]: [{ id, xPct, yPct, text, color, size }] }
  signatures: {},// { [pageIdx]: [{ id, sigImgDataUrl, stampDataUrl, xPct, yPct, wPct, hPct, aspect, baseRot, showStamp, signerName, signerReason, signedAt, sha256Hex }] }
  textEdits: {}, // { [pageIdx]: { [key]: { key, xPct, yPct, wPct, hPct, baseRot, origText, text, color, bgColor, fontSizePt, dir, bold } } }
  activeOrigEditRef: null
};

function getEffectivePdfZoom() {
  return typeof curPdf.zoom === 'number' ? curPdf.zoom : (curPdf._effectiveZoom || 1);
}

function configurePdfJsWorker() {
  const lib = window.pdfjsLib;
  if (lib && lib.GlobalWorkerOptions && !lib.GlobalWorkerOptions.workerSrc) {
    lib.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.js';
  }
  return lib || null;
}

function getPdfLib() {
  const pl = window.PDFLib;
  if (!pl || !pl.PDFDocument) {
    throw new UserError('لم يتم تحميل مكتبة pdf-lib.min.js بشكل صحيح.');
  }
  return pl;
}

// Keep one PDF.js document alive while the underlying PDF bytes stay unchanged.
// Re-parsing the complete file on every page click caused avoidable multi-second delays.
let pdfViewCacheBytes = null;
let pdfViewCachePromise = null;
async function loadPdfDocumentForView(bytes) {
  const lib = configurePdfJsWorker();
  if (!lib) throw new UserError('لم يتم تحميل مكتبة pdf.min.js بشكل صحيح.');
  if (pdfViewCacheBytes === bytes && pdfViewCachePromise) return pdfViewCachePromise;
  const copy = bytes instanceof Uint8Array ? bytes.slice() : new Uint8Array(bytes);
  const task = lib.getDocument({ data: copy, isEvalSupported: false });
  pdfViewCacheBytes = bytes;
  pdfViewCachePromise = task.promise;
  try {
    return await pdfViewCachePromise;
  } catch (err) {
    if (pdfViewCacheBytes === bytes) {
      pdfViewCacheBytes = null;
      pdfViewCachePromise = null;
    }
    throw err;
  }
}

function resetPdfStateForBytes(bytes, keepPageIdx = 0) {
  const u8 = bytes instanceof Uint8Array ? bytes.slice() : new Uint8Array(bytes);
  curPdf.bytes = u8;
  curPdf.rotations = [];
  curPdf.drawings = {};
  curPdf.texts = {};
  curPdf.signatures = {};
  curPdf.textEdits = {};
  curPdf.activeOrigEditRef = null;
  curPdf.pageIdx = Math.max(0, keepPageIdx);
}

async function initPdfFromFile(file) {
  assertFileSize(file); const buf = await file.arrayBuffer();
  const u8 = new Uint8Array(buf);
  const { PDFDocument } = getPdfLib();
  const srcDoc = await PDFDocument.load(u8, { ignoreEncryption: true });
  const count = srcDoc.getPageCount();
  if (!count) throw new UserError('ملف PDF فارغ ولا يحتوي على صفحات.');

  resetPdfStateForBytes(u8, 0);
  curPdf.zoom = window.innerWidth <= 680 ? 'fit' : 1.15;
  curPdf.pageCount = count;
  curPdf.rotations = srcDoc.getPages().map(page => ((page.getRotation()?.angle || 0) % 360 + 360) % 360);
  kind = 'pdf';
  fileName = file.name || 'document.pdf';
  if ($('#fname')) $('#fname').textContent = fileName;
  resetHistory();
  showView('editor');
  await renderPdfEditor();
  applyLanguage(currentLanguage);
  toast('تم فتح الملف بنجاح');
}

// نافذة ترتيب ملفات الدمج: نظام ويندوز يرجّع الملفات المحددة بترتيب الأسماء مو بترتيب اختيارك،
// فنخلي المستخدم يرتبها بالسحب (نفس طريقة سحب مصغرات صفحات المحرر) قبل الدمج.
function openMergeOrderDialog(files) {
  return new Promise((resolve) => {
    const en = currentLanguage === 'en';
    let current = files.map((f, i) => ({ f, id: i }));
    let dragFrom = -1;
    let scrollSpeed = 0;
    let scrollRaf = 0;
    const prevFocus = document.activeElement;

    const overlay = h('div', { class: 'merge-order-overlay' });
    const box = h('div', { class: 'merge-order-box', role: 'dialog', 'aria-modal': 'true' });
    const list = h('div', { class: 'merge-order-list' });

    const fmtSize = (n) => n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';

    // تمرير تلقائي أثناء السحب: كل ما قرّبت من حافة القائمة (أو طلعت عنها) يسرع التمرير.
    const ZONE = 110;
    const onDocDragOver = (e) => {
      if (dragFrom < 0) return;
      const r = list.getBoundingClientRect();
      let v = 0;
      if (e.clientY < r.top + ZONE) {
        const d = Math.min(1, (r.top + ZONE - e.clientY) / ZONE);
        v = -(4 + d * d * 46);
      } else if (e.clientY > r.bottom - ZONE) {
        const d = Math.min(1, (e.clientY - (r.bottom - ZONE)) / ZONE);
        v = 4 + d * d * 46;
      }
      scrollSpeed = v;
    };
    const onDocWheel = (e) => {
      if (dragFrom < 0) return;
      e.preventDefault();
      list.scrollTop += e.deltaY;
    };
    const scrollTick = () => {
      if (scrollSpeed) list.scrollTop += scrollSpeed;
      scrollRaf = requestAnimationFrame(scrollTick);
    };
    const startAutoScroll = () => {
      document.addEventListener('dragover', onDocDragOver, true);
      document.addEventListener('wheel', onDocWheel, { passive: false, capture: true });
      cancelAnimationFrame(scrollRaf);
      scrollRaf = requestAnimationFrame(scrollTick);
    };
    const stopAutoScroll = () => {
      scrollSpeed = 0;
      cancelAnimationFrame(scrollRaf);
      document.removeEventListener('dragover', onDocDragOver, true);
      document.removeEventListener('wheel', onDocWheel, { capture: true });
    };

    const close = (result) => {
      stopAutoScroll();
      document.removeEventListener('keydown', onKey, true);
      overlay.remove();
      try { prevFocus && prevFocus.focus && prevFocus.focus(); } catch (_) {}
      resolve(result);
    };
    const onKey = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(null); } };
    document.addEventListener('keydown', onKey, true);

    const clearMarks = () => list.querySelectorAll('.drop-before,.drop-after').forEach(r => r.classList.remove('drop-before', 'drop-after'));

    function draw() {
      list.replaceChildren();
      list.append(h('div', { class: 'merge-order-row fixed' },
        h('span', { class: 'merge-order-num' }, '•'),
        h('span', { class: 'merge-order-name' }, en ? 'Current PDF (stays first)' : 'الملف الحالي (يبقى أول شي)')
      ));
      current.forEach((it, i) => {
        const row = h('div', {
          class: 'merge-order-row',
          draggable: true,
          'data-index': String(i),
          ondragstart: (e) => {
            dragFrom = i;
            e.dataTransfer.setData('text/plain', String(i));
            e.dataTransfer.effectAllowed = 'move';
            row.classList.add('dragging');
            startAutoScroll();
          },
          ondragend: () => { dragFrom = -1; stopAutoScroll(); row.classList.remove('dragging'); clearMarks(); },
          ondragover: (e) => {
            if (dragFrom < 0) return;
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            clearMarks();
            if (dragFrom !== i) row.classList.add(dragFrom > i ? 'drop-before' : 'drop-after');
          },
          ondragleave: (e) => { if (!row.contains(e.relatedTarget)) row.classList.remove('drop-before', 'drop-after'); },
          ondrop: (e) => {
            e.preventDefault();
            e.stopPropagation();
            const from = Number(e.dataTransfer.getData('text/plain'));
            clearMarks();
            stopAutoScroll();
            if (Number.isInteger(from) && from >= 0 && from < current.length && from !== i) {
              const [moved] = current.splice(from, 1);
              current.splice(i, 0, moved);
              const keep = list.scrollTop;
              draw();
              list.scrollTop = keep;
            }
          }
        },
          h('span', { class: 'merge-order-grip', title: en ? 'Drag to reorder' : 'اسحب للترتيب' }, '⋮⋮'),
          h('span', { class: 'merge-order-num' }, String(i + 1)),
          h('span', { class: 'merge-order-name', title: it.f.name }, it.f.name),
          h('span', { class: 'merge-order-size' }, fmtSize(it.f.size)),
          h('button', {
            type: 'button', class: 'merge-order-btn danger', title: en ? 'Remove' : 'إزالة',
            onclick: () => {
              current = current.filter(x => x.id !== it.id);
              if (!current.length) { close(null); return; }
              const keep = list.scrollTop;
              draw();
              list.scrollTop = keep;
            }
          }, '✕')
        );
        list.append(row);
      });
    }

    const okBtn = h('button', { type: 'button', class: 'merge-order-ok', onclick: () => close(current.map(x => x.f)) },
      en ? 'Merge in this order' : 'ادمج بهذا الترتيب');
    const cancelBtn = h('button', { type: 'button', class: 'merge-order-cancel', onclick: () => close(null) },
      en ? 'Cancel' : 'إلغاء');

    box.append(
      h('h3', null, en ? 'Arrange files before merging' : 'رتّب الملفات قبل الدمج'),
      h('p', { class: 'merge-order-hint' }, en
        ? 'Drag a file and drop it on the position you want. Files are added after the current PDF, top to bottom.'
        : 'اسحب الملف وأفلته على المكان اللي تبيه. الملفات تنضاف بعد الملف الحالي من فوق لتحت.'),
      list,
      h('div', { class: 'merge-order-actions' }, cancelBtn, okBtn)
    );
    overlay.append(box);
    overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(null); });
    document.body.append(overlay);
    draw();
    okBtn.focus();
  });
}

async function mergePdfFiles(filesList) {
  const files = Array.from(filesList || []).filter(f => f && f.size > 0);
  if (!files.length || !curPdf.bytes) return;

  busy(true, 'جارٍ دمج ملفات PDF…');
  try {
    const { PDFDocument } = getPdfLib();
    const baseDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
    const oldCount = baseDoc.getPageCount();

    for (const f of files) {
      assertFileSize(f); const addBytes = new Uint8Array(await f.arrayBuffer());
      const donorDoc = await PDFDocument.load(addBytes, { ignoreEncryption: true });
      const copiedPages = await baseDoc.copyPages(donorDoc, donorDoc.getPageIndices());
      copiedPages.forEach(pg => baseDoc.addPage(pg));
    }

    const mergedBytes = await baseDoc.save();
    const newCount = baseDoc.getPageCount();
    const prevRotations = curPdf.rotations.slice(0, oldCount);
    while (prevRotations.length < newCount) prevRotations.push(0);

    curPdf.bytes = mergedBytes;
    curPdf.pageCount = newCount;
    curPdf.rotations = prevRotations;
    await renderPdfEditor();
    toast(currentLanguage === 'en' ? 'PDF files merged successfully' : 'تم دمج ملفات PDF بنجاح');
  } catch (err) {
    toast(err instanceof UserError ? err.message : 'حدث خطأ أثناء دمج ملفات PDF');
  } finally {
    busy(false);
  }
}

async function deletePdfPageAt(targetIdx) {
  if (!curPdf.bytes || curPdf.pageCount <= 1) {
    toast(currentLanguage === 'en' ? 'Cannot delete the last page' : 'لا يمكن حذف آخر صفحة');
    return;
  }
  if (targetIdx < 0 || targetIdx >= curPdf.pageCount) return;

  busy(true, 'جارٍ حذف الصفحة…');
  try {
    const { PDFDocument } = getPdfLib();
    const srcDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
    const newDoc = await PDFDocument.create();
    const indicesToKeep = [];
    for (let i = 0; i < srcDoc.getPageCount(); i++) {
      if (i !== targetIdx) indicesToKeep.push(i);
    }
    const copied = await newDoc.copyPages(srcDoc, indicesToKeep);
    copied.forEach(pg => newDoc.addPage(pg));

    const updatedBytes = await newDoc.save();
    curPdf.rotations.splice(targetIdx, 1);

    const shiftMap = (obj) => {
      const next = {};
      for (const k of Object.keys(obj || {})) {
        const idx = +k;
        if (idx === targetIdx) continue;
        next[idx > targetIdx ? idx - 1 : idx] = obj[k];
      }
      return next;
    };
    curPdf.drawings = shiftMap(curPdf.drawings);
    curPdf.texts = shiftMap(curPdf.texts);
    curPdf.signatures = shiftMap(curPdf.signatures);
    curPdf.textEdits = shiftMap(curPdf.textEdits);
    curPdf.activeOrigEditRef = null;

    curPdf.bytes = updatedBytes;
    curPdf.pageCount = newDoc.getPageCount();
    if (curPdf.pageIdx >= curPdf.pageCount) {
      curPdf.pageIdx = Math.max(0, curPdf.pageCount - 1);
    }
    await renderPdfEditor();
    toast(currentLanguage === 'en' ? 'Page deleted' : 'تم حذف الصفحة بنجاح');
  } catch (err) {
    console.error('PDF page deletion failed:', err);
    toast(currentLanguage === 'en' ? 'Failed to delete page' : 'تعذر حذف الصفحة');
  } finally {
    busy(false);
  }
}

async function movePdfPage(fromIdx, toIdx) {
  if (!curPdf.bytes || fromIdx === toIdx) return;
  if (fromIdx < 0 || fromIdx >= curPdf.pageCount || toIdx < 0 || toIdx >= curPdf.pageCount) return;

  busy(true, 'جارٍ إعادة ترتيب الصفحات…');
  try {
    const { PDFDocument } = getPdfLib();
    const srcDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
    const order = Array.from({ length: srcDoc.getPageCount() }, (_, i) => i);
    const [moved] = order.splice(fromIdx, 1);
    order.splice(toIdx, 0, moved);

    const newDoc = await PDFDocument.create();
    const copied = await newDoc.copyPages(srcDoc, order);
    copied.forEach(pg => newDoc.addPage(pg));

    const nextRots = order.map(oldI => curPdf.rotations[oldI] || 0);
    const nextDraw = {};
    const nextTxt = {};
    const nextSigs = {};
    const nextEdits = {};
    order.forEach((oldI, newI) => {
      if (curPdf.drawings[oldI]) nextDraw[newI] = curPdf.drawings[oldI];
      if (curPdf.texts[oldI]) nextTxt[newI] = curPdf.texts[oldI];
      if (curPdf.signatures[oldI]) nextSigs[newI] = curPdf.signatures[oldI];
      if (curPdf.textEdits[oldI]) nextEdits[newI] = curPdf.textEdits[oldI];
    });

    curPdf.bytes = await newDoc.save();
    curPdf.pageCount = newDoc.getPageCount();
    curPdf.rotations = nextRots;
    curPdf.drawings = nextDraw;
    curPdf.texts = nextTxt;
    curPdf.signatures = nextSigs;
    curPdf.textEdits = nextEdits;
    curPdf.activeOrigEditRef = null;
    curPdf.pageIdx = toIdx;
    await renderPdfEditor();
    toast(currentLanguage === 'en' ? 'Page order updated' : 'تم تحديث ترتيب الصفحات');
  } catch (err) {
    console.error('PDF page reordering failed:', err);
    toast(currentLanguage === 'en' ? 'Failed to reorder pages' : 'تعذر إعادة ترتيب الصفحات');
  } finally {
    busy(false);
  }
}

function rotateCurrentPdfPage(degDelta) {
  if (!curPdf.pageCount) return;
  const idx = curPdf.pageIdx;
  const cur = curPdf.rotations[idx] || 0;
  curPdf.rotations[idx] = (((cur + degDelta) % 360) + 360) % 360;
  renderPdfEditor();
  toast(currentLanguage === 'en' ? 'Page rotated' : 'تم تدوير الصفحة');
}

async function insertBlankPdfPageAfterCurrent() {
  if (!curPdf.bytes) return false;
  busy(true, currentLanguage === 'en' ? 'Adding blank page…' : 'جارٍ إضافة صفحة فارغة…');
  try {
    const { PDFDocument } = getPdfLib();
    const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
    const curIdx = Math.max(0, Math.min(pdfDoc.getPageCount() - 1, curPdf.pageIdx || 0));
    const refPage = pdfDoc.getPages()[curIdx];
    const { width, height } = refPage ? refPage.getSize() : { width: 595.28, height: 841.89 };
    const insertAt = curIdx + 1;
    pdfDoc.insertPage(insertAt, [width, height]);

    const shiftUp = (obj) => {
      const next = {};
      for (const k of Object.keys(obj || {})) {
        const idx = +k;
        next[idx >= insertAt ? idx + 1 : idx] = obj[k];
      }
      return next;
    };
    curPdf.rotations.splice(insertAt, 0, 0);
    curPdf.drawings = shiftUp(curPdf.drawings);
    curPdf.texts = shiftUp(curPdf.texts);
    curPdf.signatures = shiftUp(curPdf.signatures);
    curPdf.textEdits = shiftUp(curPdf.textEdits);
    curPdf.activeOrigEditRef = null;
    curPdf.bytes = await pdfDoc.save();
    curPdf.pageCount = pdfDoc.getPageCount();
    curPdf.pageIdx = insertAt;
    await renderPdfEditor();
    toast(tr('تمت إضافة صفحة فارغة'));
    return true;
  } catch (err) {
    console.error('Insert blank PDF page failed:', err);
    toast(currentLanguage === 'en' ? 'Could not insert blank page' : 'تعذر إضافة صفحة فارغة');
    return false;
  } finally {
    busy(false);
  }
}

async function extractCurrentPdfPages() {
  if (!curPdf.bytes || !curPdf.pageCount) return;
  const total = curPdf.pageCount;
  const raw = prompt(
    currentLanguage === 'en'
      ? `Enter page numbers or ranges to extract into a new PDF (1-${total}), e.g. 1-3, 5:`
      : `أدخل أرقام أو نطاقات الصفحات لاستخراجها في ملف PDF مستقل (1-${total})، مثال: 1-3, 5:`,
    String((curPdf.pageIdx || 0) + 1)
  );
  if (raw == null || !raw.trim()) return;

  const indices = [];
  for (const part of raw.split(',').map(s => s.trim()).filter(Boolean)) {
    const m = /^(\d+)\s*(?:-\s*(\d+))?$/.exec(part);
    if (!m) {
      toast(currentLanguage === 'en' ? `Invalid range: ${part}` : `نطاق غير صالح: ${part}`);
      return;
    }
    const a = +m[1], b = +(m[2] || m[1]);
    if (a < 1 || b < a || b > total) {
      toast(currentLanguage === 'en' ? `Range out of bounds (1-${total})` : `النطاق خارج حدود الصفحات (1-${total})`);
      return;
    }
    for (let p = a - 1; p <= b - 1; p++) indices.push(p);
  }
  if (!indices.length) return;

  busy(true, currentLanguage === 'en' ? 'Extracting pages…' : 'جارٍ استخراج الصفحات…');
  try {
    const mergedBytes = await buildMergedPdfBytesWithAnnotations();
    const { PDFDocument } = getPdfLib();
    const srcDoc = await PDFDocument.load(mergedBytes, { ignoreEncryption: true });
    const outDoc = await PDFDocument.create();
    const copied = await outDoc.copyPages(srcDoc, indices);
    copied.forEach(pg => outDoc.addPage(pg));
    const outBytes = await outDoc.save();
    const blob = new Blob([outBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const base = (fileName || 'document').replace(/\.pdf$/i, '');
    a.href = url;
    a.download = `${base}-extracted.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast(tr('تم استخراج الصفحات بنجاح'));
  } catch (err) {
    console.error('Extract PDF pages failed:', err);
    toast(currentLanguage === 'en' ? 'Could not extract pages' : 'تعذر استخراج الصفحات');
  } finally {
    busy(false);
  }
}

async function openPdfSearchModal() {
  if (!curPdf.bytes) return;
  $('#pdfSearchBackdrop')?.remove();

  const searchIn = h('input', {
    type: 'text',
    placeholder: tr('ابحث في صفحات PDF…'),
    style: { flex: '1', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px 10px', background: 'var(--soft)', color: 'var(--ink)' }
  });
  const resultsBox = h('div', {
    style: { display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '52vh', overflow: 'auto', marginTop: '4px' }
  });

  const close = () => {
    document.removeEventListener('keydown', onEsc, true);
    backdrop.remove();
  };
  const onEsc = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  document.addEventListener('keydown', onEsc, true);

  const runSearch = async () => {
    const q = searchIn.value.trim();
    if (!q) { searchIn.focus(); return; }
    resultsBox.replaceChildren(h('small', { class: 'muted' }, tr('جارٍ التحميل…')));
    try {
      const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
      const qLower = q.toLowerCase();
      const hits = [];
      for (let i = 0; i < pdfDoc.numPages; i++) {
        const page = await pdfDoc.getPage(i + 1);
        const tc = await page.getTextContent();
        const streamTxt = (tc?.items || []).map(it => normalizeExtractedPdfStr(it.str || '')).join(' ');
        const extraTxt = [
          ...Object.values(curPdf.textEdits[i] || {}).map(e => e.text || ''),
          ...(curPdf.texts[i] || []).map(t => t.text || '')
        ].join(' ');
        const combined = (streamTxt + ' ' + extraTxt).replace(/\s+/g, ' ').trim();
        const at = combined.toLowerCase().indexOf(qLower);
        if (at !== -1) {
          const start = Math.max(0, at - 36);
          const end = Math.min(combined.length, at + q.length + 48);
          hits.push({
            pageIdx: i,
            snippet: (start > 0 ? '…' : '') + combined.slice(start, end) + (end < combined.length ? '…' : '')
          });
        }
      }
      resultsBox.replaceChildren();
      if (!hits.length) {
        resultsBox.append(h('div', { class: 'muted', style: { padding: '8px 2px' } }, tr('لم يتم العثور على نتائج في هذا الملف')));
        return;
      }
      hits.forEach(hit => {
        const row = h('button', {
          type: 'button',
          class: 'btn',
          style: { textAlign: 'start', display: 'flex', flexDirection: 'column', gap: '2px', whiteSpace: 'normal', padding: '8px 10px' },
          onclick: () => {
            curPdf.pageIdx = hit.pageIdx;
            close();
            renderPdfEditor();
          }
        },
          h('strong', { style: { color: 'var(--acc)', fontSize: '.82rem' } }, (currentLanguage === 'en' ? 'Page ' : 'صفحة ') + (hit.pageIdx + 1)),
          h('span', { style: { fontSize: '.84rem' } }, hit.snippet)
        );
        resultsBox.append(row);
      });
    } catch (err) {
      console.error('PDF search failed:', err);
      resultsBox.replaceChildren(h('div', { class: 'muted' }, tr('لم يتم العثور على نتائج في هذا الملف')));
    }
  };

  searchIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); runSearch(); } });

  const backdrop = h('div', { id: 'pdfSearchBackdrop', class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) close(); } },
    h('div', { class: 'md-modal', style: { width: 'min(500px, 95vw)' } },
      h('div', { class: 'row', style: { justifyContent: 'space-between' } },
        h('strong', null, '🔎 ' + tr('نتائج البحث في PDF')),
        h('button', { type: 'button', class: 'btn', style: { padding: '2px 8px' }, onclick: close }, '✕')
      ),
      h('div', { class: 'row' }, searchIn, btn('🔎', runSearch, 'pri')),
      resultsBox
    )
  );
  document.body.append(backdrop);
  searchIn.focus();
}

async function insertPdfImageFileOnCurrentPage(file, fromClipboard = false) {
  if (!file || !curPdf.bytes) return false;
  busy(true, 'جارٍ تجهيز الصورة…');
  try {
    const processed = await processUploadedSignatureImage(file, false);
    if (!processed || !processed.dataUrl) throw new Error('Empty image');
    await insertPdfSignatureOnCurrentPage({
      sigImgDataUrl: processed.dataUrl,
      showStamp: false
    });
    toast(fromClipboard ? 'تم لصق الصورة من الحافظة' : 'تمت إضافة الصورة');
    return true;
  } catch (err) {
    toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة');
    return false;
  } finally {
    busy(false);
  }
}

function triggerPdfImageInsert() {
  const inp = $('#pdfImgIn') || $('#imgIn');
  if (!inp) return;
  inp.value = '';
  inp.onchange = async () => {
    const f = inp.files?.[0];
    if (f) await insertPdfImageFileOnCurrentPage(f, false);
    inp.value = '';
  };
  inp.click();
}
/* ═════════════════════════════════════════════════════════════
   محرر PDF المتكامل محلياً (PDF.js + pdf-lib)
   الجزء 2 من 2: إضافة النصوص، القلم الحر، التصدير، والربط بالمحرر
   ═════════════════════════════════════════════════════════════ */

function ensurePdfPageArrays(pageIdx) {
  if (!curPdf.drawings[pageIdx]) curPdf.drawings[pageIdx] = [];
  if (!curPdf.texts[pageIdx]) curPdf.texts[pageIdx] = [];
  if (!curPdf.signatures[pageIdx]) curPdf.signatures[pageIdx] = [];
  if (!curPdf.textEdits[pageIdx]) curPdf.textEdits[pageIdx] = {};
}

function projectPdfBoxFromRot(box, fromRot, toRot) {
  const delta = (((toRot - (fromRot || 0)) % 360) + 360) % 360;
  const { xPct, yPct, wPct, hPct } = box;
  if (delta === 90) return { xPct: 100 - yPct - hPct, yPct: xPct, wPct: hPct, hPct: wPct };
  if (delta === 180) return { xPct: 100 - xPct - wPct, yPct: 100 - yPct - hPct, wPct, hPct };
  if (delta === 270) return { xPct: yPct, yPct: 100 - xPct - wPct, wPct: hPct, hPct: wPct };
  return { xPct, yPct, wPct, hPct };
}

function rgbBytesToHex(r, g, b) {
  return '#' + [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');
}

function samplePdfTextColors(renderCtx, cw, ch, xPct, yPct, wPct, hPct) {
  try {
    const x0 = Math.max(0, Math.floor((xPct / 100) * cw));
    const y0 = Math.max(0, Math.floor((yPct / 100) * ch));
    const w = Math.max(2, Math.min(cw - x0, Math.ceil((wPct / 100) * cw)));
    const hPx = Math.max(2, Math.min(ch - y0, Math.ceil((hPct / 100) * ch)));
    const sx = Math.max(0, x0 - 2);
    const sy = Math.max(0, y0 - 2);
    const sw = Math.max(2, Math.min(cw - sx, w + 4));
    const sh = Math.max(2, Math.min(ch - sy, hPx + 4));
    const data = renderCtx.getImageData(sx, sy, sw, sh).data;

    const buckets = new Map();
    let bestBucket = null;
    for (let py = 0; py < sh; py++) {
      for (let px = 0; px < sw; px++) {
        if (py > 0 && py < sh - 1 && px > 0 && px < sw - 1) continue;
        const off = (py * sw + px) * 4;
        const r = data[off], g = data[off + 1], b = data[off + 2];
        const key = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
        let bkt = buckets.get(key);
        if (!bkt) { bkt = { count: 0, rSum: 0, gSum: 0, bSum: 0 }; buckets.set(key, bkt); }
        bkt.count++; bkt.rSum += r; bkt.gSum += g; bkt.bSum += b;
        if (!bestBucket || bkt.count > bestBucket.count) bestBucket = bkt;
      }
    }
    const bgR = bestBucket ? Math.round(bestBucket.rSum / bestBucket.count) : 255;
    const bgG = bestBucket ? Math.round(bestBucket.gSum / bestBucket.count) : 255;
    const bgB = bestBucket ? Math.round(bestBucket.bSum / bestBucket.count) : 255;
    const bgColor = rgbBytesToHex(bgR, bgG, bgB);

    let maxDist = 0;
    const stepX = sw > 120 ? 2 : 1;
    for (let py = 1; py < sh - 1; py++) {
      for (let px = 1; px < sw - 1; px += stepX) {
        const off = (py * sw + px) * 4;
        const dist = Math.abs(data[off] - bgR) + Math.abs(data[off + 1] - bgG) + Math.abs(data[off + 2] - bgB);
        if (dist > maxDist) maxDist = dist;
      }
    }
    let inkColor = (bgR * 0.299 + bgG * 0.587 + bgB * 0.114) > 128 ? '#1f1c17' : '#ffffff';
    if (maxDist >= 85) {
      const thresh = maxDist * 0.8;
      let rSum = 0, gSum = 0, bSum = 0, cnt = 0;
      for (let py = 1; py < sh - 1; py++) {
        for (let px = 1; px < sw - 1; px += stepX) {
          const off = (py * sw + px) * 4;
          const r = data[off], g = data[off + 1], b = data[off + 2];
          const dist = Math.abs(r - bgR) + Math.abs(g - bgG) + Math.abs(b - bgB);
          if (dist >= thresh) { rSum += r; gSum += g; bSum += b; cnt++; }
        }
      }
      if (cnt > 0) inkColor = rgbBytesToHex(rSum / cnt, gSum / cnt, bSum / cnt);
    }
    return { bgColor, color: inkColor };
  } catch (_) {
    return { bgColor: '#ffffff', color: '#1f1c17' };
  }
}

/* ── محرك تفكيك وتعديل أوامر رسم النصوص في PDF (PDF Content Stream Surgical Engine) ── */
function mulPdfMat(m1, m2) {
  return [
    m1[0] * m2[0] + m1[1] * m2[2],
    m1[0] * m2[1] + m1[1] * m2[3],
    m1[2] * m2[0] + m1[3] * m2[2],
    m1[2] * m2[1] + m1[3] * m2[3],
    m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
    m1[4] * m2[1] + m1[5] * m2[3] + m2[5]
  ];
}

function distPointToSeg(px, py, x0, y0, x1, y1) {
  const dx = x1 - x0, dy = y1 - y0;
  const lenSq = dx * dx + dy * dy;
  if (lenSq < 1e-6) return Math.hypot(px - x0, py - y0);
  const t = Math.max(0, Math.min(1, ((px - x0) * dx + (py - y0) * dy) / lenSq));
  return Math.hypot(px - (x0 + t * dx), py - (y0 + t * dy));
}

function bytesToLatin1Str(u8) {
  let s = '';
  for (let i = 0; i < u8.length; i += 8192) {
    s += String.fromCharCode.apply(null, u8.subarray(i, i + 8192));
  }
  return s;
}

function latin1StrToBytes(s) {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

function isPdfWsCode(c) {
  return c === 32 || c === 10 || c === 13 || c === 9 || c === 12 || c === 0;
}

function isPdfDelimCode(c) {
  return c === 40 || c === 41 || c === 60 || c === 62 || c === 91 || c === 93 || c === 123 || c === 125 || c === 47 || c === 37;
}

function tokenizePdfStreamStr(src) {
  const tokens = [];
  const n = src.length;
  let i = 0;

  const readOneToken = () => {
    while (i < n) {
      const c = src.charCodeAt(i);
      if (isPdfWsCode(c)) { i++; continue; }
      if (c === 37) { // '%' comment
        while (i < n && src.charCodeAt(i) !== 10 && src.charCodeAt(i) !== 13) i++;
        continue;
      }
      break;
    }
    if (i >= n) return null;

    const start = i;
    const c = src.charCodeAt(i);

    // Literal string (...)
    if (c === 40) {
      let depth = 1;
      let charLen = 0;
      i++;
      while (i < n && depth > 0) {
        const ch = src.charCodeAt(i);
        if (ch === 92) { i += 2; charLen++; continue; }
        if (ch === 40) { depth++; charLen++; i++; continue; }
        if (ch === 41) { depth--; if (depth > 0) charLen++; i++; continue; }
        charLen++;
        i++;
      }
      return { type: 'str', len: charLen, rawStart: start, rawEnd: i };
    }

    // Dict << >> or Hex string <...>
    if (c === 60) {
      if (i + 1 < n && src.charCodeAt(i + 1) === 60) {
        i += 2;
        return { type: 'delim', val: '<<', rawStart: start, rawEnd: i };
      }
      i++;
      let hexDigits = 0;
      while (i < n && src.charCodeAt(i) !== 62) {
        if (!isPdfWsCode(src.charCodeAt(i))) hexDigits++;
        i++;
      }
      if (i < n) i++;
      return { type: 'hex', hexDigits, rawStart: start, rawEnd: i };
    }

    if (c === 62 && i + 1 < n && src.charCodeAt(i + 1) === 62) {
      i += 2;
      return { type: 'delim', val: '>>', rawStart: start, rawEnd: i };
    }

    // Array [...]
    if (c === 91) {
      i++;
      const items = [];
      while (i < n) {
        while (i < n && isPdfWsCode(src.charCodeAt(i))) i++;
        if (i >= n) break;
        if (src.charCodeAt(i) === 93) { i++; break; }
        const sub = readOneToken();
        if (!sub) break;
        items.push(sub);
      }
      return { type: 'arr', items, rawStart: start, rawEnd: i };
    }

    if (c === 93) {
      i++;
      return { type: 'delim', val: ']', rawStart: start, rawEnd: i };
    }

    // Name /Foo
    if (c === 47) {
      i++;
      while (i < n && !isPdfWsCode(src.charCodeAt(i)) && !isPdfDelimCode(src.charCodeAt(i))) i++;
      return { type: 'name', val: src.slice(start + 1, i), rawStart: start, rawEnd: i };
    }

    // Number or Keyword Operator
    while (i < n && !isPdfWsCode(src.charCodeAt(i)) && !isPdfDelimCode(src.charCodeAt(i))) i++;
    const word = src.slice(start, i);
    if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(word)) {
      return { type: 'num', val: parseFloat(word), rawStart: start, rawEnd: i };
    }

    // Skip inline image binary payload BI ... ID <bytes> EI safely
    if (word === 'BI') {
      while (i < n) {
        if (src.charCodeAt(i) === 73 && src.charCodeAt(i + 1) === 68 && isPdfWsCode(src.charCodeAt(i - 1)) && isPdfWsCode(src.charCodeAt(i + 2))) {
          i += 3;
          break;
        }
        i++;
      }
      while (i < n) {
        if (src.charCodeAt(i) === 69 && src.charCodeAt(i + 1) === 73 && isPdfWsCode(src.charCodeAt(i - 1)) && (i + 2 >= n || isPdfWsCode(src.charCodeAt(i + 2)) || isPdfDelimCode(src.charCodeAt(i + 2)))) {
          i += 2;
          break;
        }
        i++;
      }
      return { type: 'inlineImg', rawStart: start, rawEnd: i };
    }

    return { type: 'op', val: word, rawStart: start, rawEnd: i };
  };

  while (i < n) {
    const tk = readOneToken();
    if (!tk) break;
    tokens.push(tk);
  }
  return tokens;
}

function estimatePdfShowAdvance(opName, operands, fontSize, horizScale, isMultiByteFont) {
  const bytesPerGlyph = isMultiByteFont ? 2 : 1;
  const avgGlyphFactor = 0.52;
  const tokenAdvance = (tk) => {
    if (!tk) return 0;
    if (tk.type === 'str') return (tk.len / bytesPerGlyph) * fontSize * avgGlyphFactor * horizScale;
    if (tk.type === 'hex') return ((tk.hexDigits / 2) / bytesPerGlyph) * fontSize * avgGlyphFactor * horizScale;
    return 0;
  };
  if (opName === 'TJ') {
    const arr = operands[operands.length - 1];
    if (!arr || arr.type !== 'arr') return 0;
    let total = 0;
    for (const sub of arr.items) {
      if (sub.type === 'num') total -= (sub.val / 1000) * fontSize * horizScale;
      else total += tokenAdvance(sub);
    }
    return total;
  }
  return tokenAdvance(operands[operands.length - 1]);
}

function stripTextOpsFromStreamBytes(streamBytes, targetSegs, initialCtm, multiByteFontNames, onInvokeXObject) {
  const src = bytesToLatin1Str(streamBytes);
  const tokens = tokenizePdfStreamStr(src);
  const patches = [];
  let ctm = [...initialCtm];
  const ctmStack = [];
  let tm = [1, 0, 0, 1, 0, 0];
  let tlm = [1, 0, 0, 1, 0, 0];
  let tl = 0;
  let fontSize = 12;
  let activeFontName = '';
  let textRise = 0;
  let horizScale = 1;
  let operands = [];

  for (const tk of tokens) {
    if (tk.type === 'inlineImg') {
      operands = [];
      continue;
    }
    if (tk.type !== 'op') {
      operands.push(tk);
      continue;
    }

    const op = tk.val;
    const stmtStart = operands.length ? operands[0].rawStart : tk.rawStart;
    const stmtEnd = tk.rawEnd;

    if (op === 'q') {
      ctmStack.push({ ctm: [...ctm], fontSize, activeFontName, tl, textRise, horizScale });
    } else if (op === 'Q') {
      const prev = ctmStack.pop();
      if (prev) {
        ctm = prev.ctm; fontSize = prev.fontSize; activeFontName = prev.activeFontName;
        tl = prev.tl; textRise = prev.textRise; horizScale = prev.horizScale;
      }
    } else if (op === 'cm' && operands.length >= 6) {
      const nums = operands.slice(-6).map(o => o.val || 0);
      ctm = mulPdfMat(nums, ctm);
    } else if (op === 'BT') {
      tm = [1, 0, 0, 1, 0, 0];
      tlm = [1, 0, 0, 1, 0, 0];
    } else if (op === 'Tf' && operands.length >= 2) {
      const fnTk = operands[operands.length - 2];
      const szTk = operands[operands.length - 1];
      if (fnTk?.type === 'name') activeFontName = fnTk.val;
      if (szTk?.type === 'num') fontSize = szTk.val;
    } else if (op === 'Tm' && operands.length >= 6) {
      const nums = operands.slice(-6).map(o => o.val || 0);
      tm = [...nums];
      tlm = [...nums];
    } else if (op === 'Td' && operands.length >= 2) {
      const tx = operands[operands.length - 2].val || 0;
      const ty = operands[operands.length - 1].val || 0;
      tlm = mulPdfMat([1, 0, 0, 1, tx, ty], tlm);
      tm = [...tlm];
    } else if (op === 'TD' && operands.length >= 2) {
      const tx = operands[operands.length - 2].val || 0;
      const ty = operands[operands.length - 1].val || 0;
      tl = -ty;
      tlm = mulPdfMat([1, 0, 0, 1, tx, ty], tlm);
      tm = [...tlm];
    } else if (op === 'TL' && operands.length >= 1) {
      tl = operands[operands.length - 1].val || 0;
    } else if (op === 'T*') {
      tlm = mulPdfMat([1, 0, 0, 1, 0, -tl], tlm);
      tm = [...tlm];
    } else if (op === 'Ts' && operands.length >= 1) {
      textRise = operands[operands.length - 1].val || 0;
    } else if (op === 'Tz' && operands.length >= 1) {
      horizScale = (operands[operands.length - 1].val || 100) / 100;
    } else if (op === 'Do' && operands.length >= 1 && onInvokeXObject) {
      const nm = operands[operands.length - 1];
      if (nm?.type === 'name') onInvokeXObject(nm.val, [...ctm]);
    } else if (op === 'Tj' || op === 'TJ' || op === "'" || op === '"') {
      if (op === "'" || op === '"') {
        tlm = mulPdfMat([1, 0, 0, 1, 0, -tl], tlm);
        tm = [...tlm];
      }
      const mUser = mulPdfMat([fontSize * horizScale, 0, 0, fontSize, 0, textRise], mulPdfMat(tm, ctm));
      const ox0 = mUser[4], oy0 = mUser[5];
      const isCid = multiByteFontNames && multiByteFontNames.has(activeFontName);
      const advLocal = estimatePdfShowAdvance(op, operands, fontSize, horizScale, isCid);
      const mEnd = mulPdfMat([fontSize * horizScale, 0, 0, fontSize, 0, textRise], mulPdfMat(mulPdfMat([1, 0, 0, 1, advLocal, 0], tm), ctm));
      const ox1 = mEnd[4], oy1 = mEnd[5];
      const midX = (ox0 + ox1) / 2, midY = (oy0 + oy1) / 2;

      let matched = false;
      for (const seg of targetSegs) {
        const tol = Math.max(2.6, (seg.uFontH || 12) * 0.42);
        if (
          distPointToSeg(ox0, oy0, seg.ux0, seg.uy0, seg.ux1, seg.uy1) <= tol ||
          distPointToSeg(midX, midY, seg.ux0, seg.uy0, seg.ux1, seg.uy1) <= tol
        ) {
          matched = true;
          break;
        }
      }

      if (matched) {
        const denom = fontSize * horizScale;
        const kernShift = Math.abs(denom) > 1e-4 ? -Math.round((advLocal / denom) * 1000) : 0;
        const rep = (op === "'" || op === '"')
          ? (kernShift ? `T* [ ${kernShift} ] TJ` : 'T*')
          : (kernShift ? `[ ${kernShift} ] TJ` : '() Tj');
        patches.push({ start: stmtStart, end: stmtEnd, rep });
      }
      tm = mulPdfMat([1, 0, 0, 1, advLocal, 0], tm);
    }

    operands = [];
  }

  if (!patches.length) return { removed: 0, bytes: streamBytes };
  let out = '';
  let cursor = 0;
  for (const p of patches) {
    out += src.slice(cursor, p.start) + p.rep;
    cursor = p.end;
  }
  out += src.slice(cursor);
  return { removed: patches.length, bytes: latin1StrToBytes(out) };
}

function collectMultiByteFontNames(pdfDoc, resourcesDict) {
  const set = new Set();
  try {
    const { PDFName, PDFDict } = getPdfLib();
    if (!resourcesDict) return set;
    const res = pdfDoc.context.lookup(resourcesDict);
    if (!(res instanceof PDFDict)) return set;
    const fonts = pdfDoc.context.lookup(res.get(PDFName.of('Font')));
    if (!(fonts instanceof PDFDict)) return set;
    for (const [k, v] of fonts.entries()) {
      const fDict = pdfDoc.context.lookup(v);
      if (fDict instanceof PDFDict) {
        const sub = fDict.get(PDFName.of('Subtype'));
        const enc = fDict.get(PDFName.of('Encoding'));
        if (String(sub) === '/Type0' || /Identity/i.test(String(enc || ''))) {
          set.add(k.decodeText ? k.decodeText() : String(k).replace(/^\//, ''));
        }
      }
    }
  } catch (_) {}
  return set;
}

function surgicallyStripPdfPageText(pdfDoc, page, targetSegs) {
  if (!targetSegs || !targetSegs.length) return 0;
  const pl = getPdfLib();
  const { PDFName, PDFArray, PDFRawStream, PDFDict, decodePDFRawStream } = pl;
  if (typeof decodePDFRawStream !== 'function') return 0;

  const contentsEntry = page.node.Contents();
  if (!contentsEntry) return 0;

  const rawStreams = [];
  if (contentsEntry instanceof PDFArray) {
    for (let i = 0; i < contentsEntry.size(); i++) {
      const st = pdfDoc.context.lookup(contentsEntry.get(i));
      if (st instanceof PDFRawStream) rawStreams.push(st);
    }
  } else {
    const st = pdfDoc.context.lookup(contentsEntry);
    if (st instanceof PDFRawStream) rawStreams.push(st);
  }
  if (!rawStreams.length) return 0;

  const decodedParts = [];
  let totalLen = 0;
  for (const st of rawStreams) {
    const dec = decodePDFRawStream(st).decode();
    decodedParts.push(dec);
    totalLen += dec.length + 1;
  }
  const merged = new Uint8Array(totalLen);
  let offset = 0;
  for (const part of decodedParts) {
    merged.set(part, offset);
    offset += part.length;
    merged[offset++] = 10; // '\n'
  }

  const resDict = page.node.Resources();
  const cidFonts = collectMultiByteFontNames(pdfDoc, resDict);
  const formInvocations = [];
  const mainRes = stripTextOpsFromStreamBytes(
    merged,
    targetSegs,
    [1, 0, 0, 1, 0, 0],
    cidFonts,
    (xobjName, invokeCtm) => formInvocations.push({ xobjName, invokeCtm })
  );

  let totalRemoved = mainRes.removed;
  if (mainRes.removed > 0) {
    const flate = pdfDoc.context.flateStream(mainRes.bytes);
    const ref = pdfDoc.context.register(flate);
    page.node.set(PDFName.of('Contents'), ref);
  }

  // Also inspect invoked Form XObjects if present
  if (formInvocations.length && resDict) {
    const xobjs = pdfDoc.context.lookup(resDict.get(PDFName.of('XObject')));
    if (xobjs instanceof PDFDict) {
      for (const inv of formInvocations) {
        const keyName = PDFName.of(inv.xobjName);
        const xobj = pdfDoc.context.lookup(xobjs.get(keyName));
        if (!(xobj instanceof PDFRawStream)) continue;
        if (String(xobj.dict.get(PDFName.of('Subtype'))) !== '/Form') continue;
        let formMat = [1, 0, 0, 1, 0, 0];
        const matArr = pdfDoc.context.lookup(xobj.dict.get(PDFName.of('Matrix')));
        if (matArr instanceof PDFArray && matArr.size() >= 6) {
          formMat = Array.from({ length: 6 }, (_, idx) => matArr.get(idx)?.asNumber?.() ?? 0);
        }
        const formBytes = decodePDFRawStream(xobj).decode();
        const formRes = xobj.dict.get(PDFName.of('Resources')) || resDict;
        const formCid = collectMultiByteFontNames(pdfDoc, formRes);
        const patchedForm = stripTextOpsFromStreamBytes(formBytes, targetSegs, mulPdfMat(formMat, inv.invokeCtm), formCid, null);
        if (patchedForm.removed > 0) {
          totalRemoved += patchedForm.removed;
          const extraDict = { Type: 'XObject', Subtype: 'Form' };
          for (const copyKey of ['BBox', 'Resources', 'Matrix', 'Group']) {
            const val = xobj.dict.get(PDFName.of(copyKey));
            if (val) extraDict[copyKey] = val;
          }
          const newFormRef = pdfDoc.context.register(pdfDoc.context.flateStream(patchedForm.bytes, extraDict));
          xobjs.set(keyName, newFormRef);
        }
      }
    }
  }

  return totalRemoved;
}

function canEncodeWinAnsi(text) {
  if (!text) return true;
  if (/[\u0590-\u08FF\u0900-\uFFFF]/.test(text)) return false;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 10 || code === 13 || code === 9) continue;
    if (code >= 0x20 && code <= 0x7E) continue;
    if (code >= 0xA0 && code <= 0xFF) continue;
    return false;
  }
  return true;
}

async function applySurgicalPdfTextEdit(pageIdx, block, newText, colorHex, fontSizePt) {
  if (!curPdf.bytes || pageIdx < 0 || pageIdx >= curPdf.pageCount) return false;
  const cleanText = String(newText ?? '').replace(/\r\n/g, '\n');
  const activeColor = colorHex || block.color || '#1f1c17';
  const activeSize = Math.max(6, Math.min(96, +(fontSizePt || block.fontSizePt || 12)));
  const curRot = curPdf.rotations[pageIdx] || 0;
  ensurePdfPageArrays(pageIdx);

  busy(true, currentLanguage === 'en' ? 'Updating PDF stream…' : 'جارٍ تحديث نص الـ PDF…');
  try {
    const { PDFDocument, StandardFonts, degrees } = getPdfLib();
    const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
    const page = pdfDoc.getPages()[pageIdx];
    if (!page) return false;

    let removedFromStream = block.isUnicodeLive ? (block.strippedFromStream !== false) : false;
    if (!block.isUnicodeLive && block.rawSegments?.length) {
      const removed = surgicallyStripPdfPageText(pdfDoc, page, block.rawSegments);
      removedFromStream = removed > 0;
    }

    // Remove previous live unicode entry if this was already a unicode-live block
    if (block.key && curPdf.textEdits[pageIdx]?.[block.key]) {
      delete curPdf.textEdits[pageIdx][block.key];
    }

    const { width, height } = page.getSize();
    const viewW = curRot % 180 ? height : width;
    const viewH = curRot % 180 ? width : height;
    const proj = projectPdfBoxFromRot(block, block.baseRot ?? curRot, curRot);

    // Rare fallback only if PDF text was drawn as outlined paths / non-standard operators
    if (!removedFromStream && !block.isUnicodeLive) {
      const padX = (1.8 / Math.max(1, viewW)) * 100;
      const padY = (1.4 / Math.max(1, viewH)) * 100;
      const c1 = pdfViewportPointToPage(Math.max(0, (proj.xPct - padX) / 100), Math.max(0, (proj.yPct - padY) / 100), width, height, curRot);
      const c2 = pdfViewportPointToPage(Math.min(1, (proj.xPct + proj.wPct + padX) / 100), Math.min(1, (proj.yPct + proj.hPct + padY) / 100), width, height, curRot);
      page.drawRectangle({
        x: Math.min(c1.x, c2.x),
        y: Math.min(c1.y, c2.y),
        width: Math.abs(c2.x - c1.x),
        height: Math.abs(c2.y - c1.y),
        color: hexToPdfRgb(block.bgColor || '#ffffff')
      });
    }

    if (cleanText.trim()) {
      if (canEncodeWinAnsi(cleanText)) {
        // True native Vector PDF Text (<BT /Font sz Tf ... Tj ET>) — 100% sharp, selectable, searchable, and re-editable!
        const fam = String(block.fontFamily || '');
        const isBold = !!block.bold;
        const stdName = /courier|mono|consolas/i.test(fam)
          ? (isBold ? StandardFonts.CourierBold : StandardFonts.Courier)
          : (!/sans/i.test(fam) && /times|serif|georgia|garamond/i.test(fam))
            ? (isBold ? StandardFonts.TimesRomanBold : StandardFonts.TimesRoman)
            : (isBold ? StandardFonts.HelveticaBold : StandardFonts.Helvetica);
        const pdfFont = await pdfDoc.embedFont(stdName);
        const lines = cleanText.split('\n');
        const baseYView = block.baselineYPct != null && (block.baseRot ?? curRot) === curRot
          ? (block.baselineYPct / 100) * viewH
          : (proj.yPct / 100) * viewH + activeSize * 0.86;
        const startXView = (proj.xPct / 100) * viewW;
        const rotAngle = degrees((360 - (curRot % 360)) % 360);

        lines.forEach((ln, li) => {
          if (!ln) return;
          const pt = pdfViewportPointToPage(startXView / viewW, (baseYView + li * activeSize * 1.18) / viewH, width, height, curRot);
          page.drawText(ln, {
            x: pt.x,
            y: pt.y,
            size: activeSize,
            font: pdfFont,
            color: hexToPdfRgb(activeColor),
            rotate: rotAngle
          });
        });
      } else {
        // Non-WinAnsi (Arabic / Unicode): keep underlying stream surgically stripped & maintain transparent live editable item
        const uKey = block.isUnicodeLive && block.key ? block.key : ('u_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));
        const nextDir = /[\u0590-\u08FF]/.test(cleanText) ? 'rtl' : 'ltr';
        curPdf.textEdits[pageIdx][uKey] = {
          key: uKey,
          isUnicodeLive: true,
          strippedFromStream: true,
          xPct: proj.xPct,
          yPct: proj.yPct,
          wPct: proj.wPct,
          hPct: proj.hPct,
          baseRot: curRot,
          origText: cleanText,
          text: cleanText,
          color: activeColor,
          bgColor: block.bgColor || '#ffffff',
          fontSizePt: activeSize,
          dir: nextDir,
          bold: !!block.bold
        };
      }
    }

    curPdf.bytes = await pdfDoc.save();
    curPdf.activeOrigEditRef = null;
    await renderPdfEditor();
    toast(tr('تم تحديث نص الـ PDF بنجاح'));
    return true;
  } catch (err) {
    console.error('Surgical PDF text edit failed:', err);
    toast(currentLanguage === 'en' ? 'Failed to update PDF text' : 'تعذر تحديث النص في ملف PDF');
    return false;
  } finally {
    busy(false);
  }
}

function normalizeExtractedPdfStr(s) {
  if (!s) return '';
  const lib = window.pdfjsLib;
  let out = typeof lib?.normalizeUnicode === 'function' ? lib.normalizeUnicode(s) : s;
  out = out.replace(/[\uFB50-\uFDFF\uFE70-\uFEFF]+/g, m => m.normalize('NFKC'));
  return out;
}

async function extractPdfPageTextBlocks(activePage, curRot, renderCanvas) {
  const lib = window.pdfjsLib;
  if (!lib || typeof activePage.getTextContent !== 'function') return [];
  let tc;
  try {
    tc = await activePage.getTextContent();
  } catch (err) {
    console.warn('PDF text extraction failed:', err);
    return [];
  }
  if (!tc || !Array.isArray(tc.items) || !tc.items.length) return [];

  const vp1 = activePage.getViewport({ scale: 1, rotation: curRot });
  const pageW = vp1.width || 1;
  const pageH = vp1.height || 1;
  const styles = tc.styles || {};
  const rawRuns = [];

  for (const it of tc.items) {
    if (!it || typeof it.str !== 'string') continue;
    const clean = normalizeExtractedPdfStr(it.str);
    if (!clean.trim()) continue;
    const tx = lib.Util ? lib.Util.transform(vp1.transform, it.transform) : it.transform;
    const [ta, tb, tcMat, td, te, tf] = tx;
    if (Math.abs(tb) > Math.max(1, Math.abs(ta)) * 0.45) continue;

    const fontH = Math.max(6, Math.min(120, Math.hypot(tcMat, td) || Math.hypot(ta, tb) || it.height || 12));
    const runW = Math.max(fontH * 0.35, Math.abs(it.width || 0));
    const bx = ta >= 0 ? te : te - runW;
    const by = tf - fontH * 0.86;
    const bh = fontH * 1.15;
    const stInfo = styles[it.fontName] || {};
    const fontFamily = String(it.fontName || '') + ' ' + String(stInfo.fontFamily || '');
    const isBold = /bold|black|heavy/i.test(fontFamily);

    // Unrotated PDF user-space baseline segment for surgical Content Stream operator matching
    const [ia, ib, ic, id, ie, if_] = it.transform;
    const lenA = Math.hypot(ia, ib) || 1;
    const uFontH = Math.hypot(ic, id) || lenA || fontH;
    const rawW = Math.abs(it.width || 0);

    rawRuns.push({
      str: clean,
      dir: it.dir || '',
      bx,
      by,
      bw: runW,
      bh,
      tf,
      fontH,
      bold: isBold,
      fontFamily,
      uSeg: {
        ux0: ie,
        uy0: if_,
        ux1: ie + (ia / lenA) * rawW,
        uy1: if_ + (ib / lenA) * rawW,
        uFontH
      }
    });
  }

  if (!rawRuns.length) return [];

  rawRuns.sort((a, b) => (Math.abs(a.tf - b.tf) > Math.max(a.fontH, b.fontH) * 0.35 ? a.tf - b.tf : a.bx - b.bx));

  const groups = [];
  for (const run of rawRuns) {
    let placed = false;
    for (let gIdx = groups.length - 1; gIdx >= 0; gIdx--) {
      const grp = groups[gIdx];
      const maxFH = Math.max(grp.fontH, run.fontH);
      if (Math.abs(run.tf - grp.tf) > maxFH * 0.38) {
        if (run.tf - grp.tf > maxFH * 1.5) break;
        continue;
      }
      const gap = Math.max(run.bx - grp.maxX, grp.minX - (run.bx + run.bw));
      if (gap <= maxFH * 1.65) {
        grp.runs.push(run);
        grp.minX = Math.min(grp.minX, run.bx);
        grp.maxX = Math.max(grp.maxX, run.bx + run.bw);
        grp.minY = Math.min(grp.minY, run.by);
        grp.maxY = Math.max(grp.maxY, run.by + run.bh);
        grp.fontH = Math.max(grp.fontH, run.fontH);
        grp.bold = grp.bold || run.bold;
        placed = true;
        break;
      }
    }
    if (!placed) {
      groups.push({
        runs: [run],
        tf: run.tf,
        minX: run.bx,
        maxX: run.bx + run.bw,
        minY: run.by,
        minFontFamily: run.fontFamily,
        maxY: run.by + run.bh,
        fontH: run.fontH,
        bold: run.bold
      });
    }
  }

  const renderCtx = renderCanvas ? renderCanvas.getContext('2d', { willReadFrequently: true }) : null;
  const cw = renderCanvas ? renderCanvas.width : 0;
  const ch = renderCanvas ? renderCanvas.height : 0;

  return groups.map((grp) => {
    const sampleText = grp.runs.map(r => r.str).join('');
    const isRtl = /[\u0590-\u08FF]/.test(sampleText) || grp.runs.some(r => r.dir === 'rtl');
    const ordered = [...grp.runs].sort((a, b) => isRtl ? ((b.bx + b.bw) - (a.bx + a.bw)) : (a.bx - b.bx));

    let fullText = '';
    for (let i = 0; i < ordered.length; i++) {
      const cur = ordered[i];
      if (i > 0) {
        const prev = ordered[i - 1];
        const edgeGap = isRtl ? (prev.bx - (cur.bx + cur.bw)) : (cur.bx - (prev.bx + prev.bw));
        const needSpace = edgeGap > grp.fontH * 0.16 && !/\s$/.test(fullText) && !/^\s/.test(cur.str);
        if (needSpace) fullText += ' ';
      }
      fullText += cur.str;
    }

    const xPct = Math.max(0, Math.min(99.5, (grp.minX / pageW) * 100));
    const yPct = Math.max(0, Math.min(99.5, (grp.minY / pageH) * 100));
    const wPct = Math.max(0.6, Math.min(100 - xPct, ((grp.maxX - grp.minX) / pageW) * 100));
    const hPct = Math.max(0.8, Math.min(100 - yPct, ((grp.maxY - grp.minY) / pageH) * 100));
    const baselineYPct = Math.max(0, Math.min(100, (grp.tf / pageH) * 100));
    const key = `b_${curRot}_${Math.round(grp.minX * 10)}_${Math.round(grp.minY * 10)}_${Math.round((grp.maxX - grp.minX) * 10)}`;
    const sampled = renderCtx && cw > 0 && ch > 0
      ? samplePdfTextColors(renderCtx, cw, ch, xPct, yPct, wPct, hPct)
      : { bgColor: '#ffffff', color: '#1f1c17' };

    return {
      key,
      xPct,
      yPct,
      wPct,
      hPct,
      baselineYPct,
      baseRot: curRot,
      origText: fullText,
      fontSizePt: Math.round(grp.fontH * 10) / 10,
      fontFamily: grp.minFontFamily || '',
      dir: isRtl ? 'rtl' : 'ltr',
      bold: !!grp.bold,
      color: sampled.color,
      bgColor: sampled.bgColor,
      rawSegments: grp.runs.map(r => r.uSeg)
    };
  });
}

function renderPdfOrigTextLayer(origLayer, blocks, pageIdx, curRot, opts) {
  if (!origLayer) return;
  origLayer.innerHTML = '';
  ensurePdfPageArrays(pageIdx);
  const unicodeMap = curPdf.textEdits[pageIdx] || {};
  const isEditMode = curPdf.tool === 'edit-text';
  origLayer.classList.toggle('active-edit-text', isEditMode);

  // Combine transparent live Unicode items (Arabic/non-WinAnsi) + freshly extracted PDF vector blocks (in edit-text mode)
  const itemsToRender = [
    ...Object.values(unicodeMap),
    ...(isEditMode ? blocks : [])
  ];

  for (const item of itemsToRender) {
    const isUnicodeLive = !!item.isUnicodeLive;
    const proj = projectPdfBoxFromRot(item, item.baseRot ?? curRot, curRot);
    const currentText = isUnicodeLive ? item.text : item.origText;
    const inkColor = item.color || '#1f1c17';
    const fontSizePt = item.fontSizePt || 12;
    const isBold = !!item.bold;
    const dir = item.dir || (/[\u0590-\u08FF]/.test(currentText) ? 'rtl' : 'ltr');

    const posStyle = {
      top: proj.yPct + '%',
      minWidth: proj.wPct + '%',
      minHeight: proj.hPct + '%',
      fontSize: Math.max(8, Math.round(fontSizePt * getEffectivePdfZoom() * 10) / 10) + 'px',
      fontWeight: isBold ? '700' : '400',
      direction: dir,
      textAlign: dir === 'rtl' ? 'right' : 'left'
    };
    if (dir === 'rtl') {
      posStyle.right = Math.max(0, 100 - (proj.xPct + proj.wPct)) + '%';
    } else {
      posStyle.left = proj.xPct + '%';
    }
    if (isUnicodeLive) {
      posStyle.color = inkColor;
    }

    const box = h('div', {
      class: 'pdf-orig-item ' + (isUnicodeLive ? 'unicode-live' : 'unmodified'),
      contenteditable: isEditMode ? 'true' : 'false',
      spellcheck: 'false',
      title: isEditMode ? tr('انقر على أي نص في الصفحة لتعديله مباشرة') : null,
      style: posStyle
    });

    const delBtn = h('button', {
      type: 'button',
      class: 'pdf-orig-del',
      contenteditable: 'false',
      title: currentLanguage === 'en' ? 'Delete text from PDF' : 'حذف هذا النص من الـ PDF',
      onclick: (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        committed = true;
        applySurgicalPdfTextEdit(pageIdx, item, '', inkColor, fontSizePt);
      }
    }, '✕');

    const textNode = document.createTextNode(currentText);
    box.append(textNode, delBtn);
    origLayer.append(box);

    if (!isEditMode) continue;

    let activeColor = inkColor;
    let activeSizePt = fontSizePt;
    let styleChanged = false;
    let committed = false;

    const readBoxPlainText = () => {
      let s = '';
      for (const ch of Array.from(box.childNodes)) {
        if (ch === delBtn) continue;
        if (ch.nodeType === 1 && ch.tagName === 'BR') { s += '\n'; continue; }
        s += ch.textContent || '';
      }
      return s.replace(/\u00A0/g, ' ');
    };

    const commitIfModified = () => {
      if (committed) return;
      const val = readBoxPlainText();
      const textChanged = val !== currentText;
      if (!textChanged && !styleChanged) {
        box.classList.remove('is-editing');
        if (!isUnicodeLive) {
          box.classList.add('unmodified');
          box.style.color = '';
        }
        return;
      }
      committed = true;
      applySurgicalPdfTextEdit(pageIdx, item, val, activeColor, activeSizePt);
    };

    box.addEventListener('pointerdown', (ev) => {
      if (ev.target === delBtn) return;
      ev.stopPropagation();
    });

    box.addEventListener('focus', () => {
      box.classList.add('is-editing');
      box.classList.remove('unmodified');
      box.style.color = activeColor;
      if (opts?.syncToolbarInputs) opts.syncToolbarInputs(activeColor, Math.round(activeSizePt));

      curPdf.activeOrigEditRef = {
        pageIdx,
        key: item.key,
        applyStyleUpdate(nextColor, nextSize) {
          if (nextColor) {
            activeColor = nextColor;
            box.style.color = nextColor;
            styleChanged = true;
          }
          if (nextSize) {
            activeSizePt = nextSize;
            box.style.fontSize = Math.max(8, Math.round(nextSize * getEffectivePdfZoom() * 10) / 10) + 'px';
            styleChanged = true;
          }
          if (document.activeElement !== box) {
            commitIfModified();
          }
        }
      };
    });

    box.addEventListener('input', () => {
      const val = readBoxPlainText();
      const nextDir = /[\u0590-\u08FF]/.test(val) ? 'rtl' : (val.trim() ? 'ltr' : dir);
      box.style.direction = nextDir;
      box.style.textAlign = nextDir === 'rtl' ? 'right' : 'left';
    });

    box.addEventListener('blur', (ev) => {
      // Allow clicking toolbar color/size inputs without prematurely committing before change fires
      const rel = ev.relatedTarget;
      if (rel && rel.closest && rel.closest('.pdf-toolbar')) {
        box.classList.remove('is-editing');
        return;
      }
      commitIfModified();
    });

    box.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' && !ev.shiftKey) {
        ev.preventDefault();
        commitIfModified();
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        committed = true;
        renderPdfEditor();
      }
    });
  }
}

/* ── تحويل ملف PDF الحالي إلى مستند Word (.docx) قابل للتحرير الكامل ── */
async function convertCurrentPdfToDocx() {
  if (!curPdf.bytes) return;
  busy(true, tr('جارٍ تحويل ملف PDF إلى Word…'));
  try {
    const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
    const total = pdfDoc.numPages;
    const newPk = createNewBlankDocx('blank');
    const doc = newPk.doc('word/document.xml');
    const body = kid(doc.documentElement, W, 'body');
    const sectPr = kid(body, W, 'sectPr');
    for (const child of Array.from(kids(body, W))) {
      if (child.localName !== 'sectPr') removeNode(child);
    }

    let picCounter = 1;

    for (let pIdx = 0; pIdx < total; pIdx++) {
      const page = await pdfDoc.getPage(pIdx + 1);
      const rot = curPdf.rotations[pIdx] || 0;
      const vp = page.getViewport({ scale: 1.35, rotation: rot });
      const sampleCvs = h('canvas', { width: Math.max(1, Math.round(vp.width)), height: Math.max(1, Math.round(vp.height)) });
      await page.render({ canvasContext: sampleCvs.getContext('2d'), viewport: vp }).promise;

      const streamBlocks = await extractPdfPageTextBlocks(page, rot, sampleCvs);
      const unicodeItems = Object.values(curPdf.textEdits[pIdx] || {});
      const floatItems = (curPdf.texts[pIdx] || []).filter(t => t.text && t.text.trim()).map(t => ({
        xPct: t.xPct,
        yPct: t.yPct,
        wPct: 25,
        hPct: 4,
        origText: t.text.trim(),
        fontSizePt: t.size || 14,
        dir: /[\u0590-\u08FF]/.test(t.text) ? 'rtl' : 'ltr',
        bold: false,
        color: t.color || '#1f1c17'
      }));

      const allLines = [
        ...streamBlocks,
        ...unicodeItems.map(u => ({ ...u, origText: u.text })),
        ...floatItems
      ].filter(b => (b.origText || '').trim());

      if (!allLines.length) {
        // Scanned / image-only PDF page: embed rendered page image cleanly into Word
        const pngBlob = await new Promise(res => sampleCvs.toBlob(res, 'image/png'));
        if (pngBlob) {
          const imgBytes = new Uint8Array(await pngBlob.arrayBuffer());
          const mediaPath = uniquePath(newPk, 'word/media', 'pdfpage', 'png');
          newPk.setBin(mediaPath, imgBytes);
          ensureDefault(newPk, 'png', 'image/png');
          const rId = addRel(newPk, newPk.mainPart, REL + 'image', mediaPath);
          const cx = 5486400;
          const cy = Math.round(cx * (sampleCvs.height / Math.max(1, sampleCvs.width)));
          const imgP = mk(doc, W, 'w:p');
          const pPr = mk(doc, W, 'w:pPr');
          pPr.appendChild(mk(doc, W, 'w:jc', { 'w:val': 'center' }));
          imgP.appendChild(pPr);
          imgP.appendChild(buildDocxDrawingRun(doc, rId, picCounter++, cx, cy));
          body.insertBefore(imgP, sectPr || null);
        }
      } else {
        allLines.sort((a, b) => (Math.abs(a.yPct - b.yPct) > 1.2 ? a.yPct - b.yPct : a.xPct - b.xPct));

        // Group blocks sharing the same horizontal row so multi-column rows become clean Word table rows
        const rows = [];
        for (const blk of allLines) {
          const lastRow = rows[rows.length - 1];
          if (lastRow && Math.abs(blk.yPct - lastRow.yPct) <= 1.25) {
            lastRow.cells.push(blk);
          } else {
            rows.push({ yPct: blk.yPct, cells: [blk] });
          }
        }

        const buildWordParaNode = (blk) => {
          const txt = (blk.origText || '').trim();
          const isRtl = blk.dir === 'rtl' || /[\u0590-\u08FF]/.test(txt);
          const szPt = Math.max(9, Math.min(36, Math.round(blk.fontSizePt || 12)));
          const centerOffset = Math.abs((blk.xPct + (blk.wPct || 0) / 2) - 50);
          const jcVal = (centerOffset < 10 && (blk.wPct || 0) < 75) ? 'center' : (isRtl ? 'right' : 'left');
          const hexCol = String(blk.color || '#1f1c17').replace('#', '').toUpperCase();

          const wP = mk(doc, W, 'w:p');
          const pPr = mk(doc, W, 'w:pPr');
          if (szPt >= 19) pPr.appendChild(mk(doc, W, 'w:pStyle', { 'w:val': 'Heading1' }));
          else if (szPt >= 15 && blk.bold) pPr.appendChild(mk(doc, W, 'w:pStyle', { 'w:val': 'Heading2' }));
          pPr.appendChild(isRtl ? mk(doc, W, 'w:bidi') : mk(doc, W, 'w:bidi', { 'w:val': '0' }));
          pPr.appendChild(mk(doc, W, 'w:jc', { 'w:val': jcVal }));
          wP.appendChild(pPr);

          const wR = mk(doc, W, 'w:r');
          const rPr = mk(doc, W, 'w:rPr');
          if (blk.bold || szPt >= 16) rPr.appendChild(mk(doc, W, 'w:b'));
          if (/^[0-9A-F]{6}$/.test(hexCol) && hexCol !== '000000' && hexCol !== '1F1C17') {
            rPr.appendChild(mk(doc, W, 'w:color', { 'w:val': hexCol }));
          }
          const halfPt = String(szPt * 2);
          rPr.appendChild(mk(doc, W, 'w:sz', { 'w:val': halfPt }));
          rPr.appendChild(mk(doc, W, 'w:szCs', { 'w:val': halfPt }));
          wR.appendChild(rPr);
          const wT = mk(doc, W, 'w:t');
          if (/^\s|\s$/.test(txt)) wT.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
          wT.textContent = txt;
          wR.appendChild(wT);
          wP.appendChild(wR);
          return wP;
        };

        for (const row of rows) {
          if (row.cells.length >= 2) {
            const orderedCells = [...row.cells].sort((a, b) => a.xPct - b.xPct);
            const tbl = mk(doc, W, 'w:tbl');
            const tblPr = mk(doc, W, 'w:tblPr');
            tblPr.appendChild(frag(doc, `<w:tblBorders xmlns:w="${W}"><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>`));
            tbl.appendChild(tblPr);
            const wTr = mk(doc, W, 'w:tr');
            for (const cellBlk of orderedCells) {
              const wTc = mk(doc, W, 'w:tc');
              wTc.appendChild(buildWordParaNode(cellBlk));
              wTr.appendChild(wTc);
            }
            tbl.appendChild(wTr);
            body.insertBefore(tbl, sectPr || null);
          } else {
            body.insertBefore(buildWordParaNode(row.cells[0]), sectPr || null);
          }
        }
      }

      if (pIdx < total - 1) {
        const brP = mk(doc, W, 'w:p');
        const brR = mk(doc, W, 'w:r');
        brR.appendChild(mk(doc, W, 'w:br', { 'w:type': 'page' }));
        brP.appendChild(brR);
        body.insertBefore(brP, sectPr || null);
      }
    }

    normalizeWordXmlPackage(newPk);
    pk = newPk;
    activeDraftId = null;
    kind = 'docx';
    fileName = (fileName || 'document').replace(/\.pdf$/i, '') + '.docx';
    if ($('#fname')) $('#fname').textContent = fileName;
    curDocx = defaultDocxState();
    resetHistory();
    pushHistory('pdfToDocx');
    renderEditor();
    toast(tr('تم تحويل ملف PDF إلى مستند Word بنجاح'));
  } catch (err) {
    console.error('PDF to DOCX conversion failed:', err);
    toast(currentLanguage === 'en' ? 'Failed to convert PDF to Word' : 'تعذر تحويل ملف PDF إلى Word');
  } finally {
    busy(false);
  }
}

function redrawPdfDrawCanvas(drawCanvas, pageIdx) {
  if (!drawCanvas) return;
  const ctx = drawCanvas.getContext('2d');
  ctx.clearRect(0, 0, drawCanvas.width, drawCanvas.height);
  const strokes = curPdf.drawings[pageIdx] || [];
  const scaleFactor = Math.max(0.35, (drawCanvas.width || 595) / 595);
  for (const st of strokes) {
    if (!st.pts || st.pts.length < 2) continue;
    ctx.save();
    if (st.mode === 'whiteout') {
      const a = st.pts[0];
      const b = st.pts[st.pts.length - 1];
      const x = Math.min(a.x, b.x) * drawCanvas.width;
      const y = Math.min(a.y, b.y) * drawCanvas.height;
      const w = Math.abs(b.x - a.x) * drawCanvas.width;
      const hPx = Math.abs(b.y - a.y) * drawCanvas.height;
      ctx.fillStyle = st.color || '#ffffff';
      ctx.fillRect(x, y, Math.max(2, w), Math.max(2, hPx));
      ctx.restore();
      continue;
    }
    if (st.mode === 'highlight') {
      ctx.globalAlpha = 0.36;
      ctx.strokeStyle = st.color || '#facc15';
      ctx.lineWidth = Math.max(8, (st.size || 18) * scaleFactor);
      ctx.lineCap = 'butt';
      ctx.lineJoin = 'round';
    } else {
      ctx.strokeStyle = st.color || '#e11d48';
      ctx.lineWidth = Math.max(1, (st.size || 3) * scaleFactor);
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
    }
    ctx.beginPath();
    ctx.moveTo(st.pts[0].x * drawCanvas.width, st.pts[0].y * drawCanvas.height);
    for (let i = 1; i < st.pts.length; i++) {
      ctx.lineTo(st.pts[i].x * drawCanvas.width, st.pts[i].y * drawCanvas.height);
    }
    ctx.stroke();
    ctx.restore();
  }
}

function renderPdfTextOverlays(textLayer, pageIdx, onAnyChange) {
  if (!textLayer) return;
  textLayer.innerHTML = '';
  const items = curPdf.texts[pageIdx] || [];

  items.forEach((item) => {
    const box = h('div', {
      class: 'pdf-text-item',
      contenteditable: 'true',
      style: {
        left: item.xPct + '%',
        top: item.yPct + '%',
        color: item.color || '#1f1c17',
        fontSize: Math.round((item.size || 18) * (4 / 3) * getEffectivePdfZoom()) + 'px'
      }
    });

    const delBtn = h('button', {
      type: 'button',
      class: 'pdf-txt-del',
      contenteditable: 'false',
      title: currentLanguage === 'en' ? 'Delete text' : 'حذف النص',
      onclick: (ev) => {
        ev.stopPropagation();
        curPdf.texts[pageIdx] = (curPdf.texts[pageIdx] || []).filter(t => t.id !== item.id);
        renderPdfTextOverlays(textLayer, pageIdx, onAnyChange);
        if (onAnyChange) onAnyChange();
      }
    }, '✕');

    const textNode = document.createTextNode(item.text || '');
    box.append(textNode, delBtn);

    const readBoxText = () => {
      let s = '';
      for (const ch of Array.from(box.childNodes)) {
        if (ch === delBtn) continue;
        s += ch.textContent || '';
      }
      return s;
    };

    box.addEventListener('input', () => {
      item.text = readBoxText();
      if (onAnyChange) onAnyChange();
    });

    box.addEventListener('pointerdown', (ev) => {
      if (ev.target === delBtn) return;
      if (curPdf.tool === 'pen') return;
      ev.stopPropagation();

      const startX = ev.clientX;
      const startY = ev.clientY;
      const origXPct = item.xPct;
      const origYPct = item.yPct;
      const rect = textLayer.getBoundingClientRect();
      let dragged = false;

      const onMove = (mv) => {
        const dx = mv.clientX - startX;
        const dy = mv.clientY - startY;
        if (!dragged && Math.hypot(dx, dy) < 4) return;
        dragged = true;
        const nextX = origXPct + (dx / (rect.width || 1)) * 100;
        const nextY = origYPct + (dy / (rect.height || 1)) * 100;
        item.xPct = Math.max(0, Math.min(92, nextX));
        item.yPct = Math.max(0, Math.min(94, nextY));
        box.style.left = item.xPct + '%';
        box.style.top = item.yPct + '%';
      };

      const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        if (dragged && onAnyChange) onAnyChange();
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    });

    textLayer.append(box);
  });
}

/* ── التوقيع الرقمي والإلكتروني المتكامل مع ختم التوثيق (SHA-256) ── */
const SAVED_PDF_SIGS_KEY = 'doc-studio-saved-signatures';

function loadSavedPdfSignatures() {
  try {
    const raw = localStorage.getItem(SAVED_PDF_SIGS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(x => x && typeof x.sigImgDataUrl === 'string' && x.sigImgDataUrl.startsWith('data:image/')).slice(0, 8);
  } catch (_) {
    return [];
  }
}

function savePdfSignatureLocally(sigImgDataUrl, label) {
  if (!sigImgDataUrl || !sigImgDataUrl.startsWith('data:image/')) return;
  try {
    const cur = loadSavedPdfSignatures().filter(x => x.sigImgDataUrl !== sigImgDataUrl);
    cur.unshift({
      id: 'ssig_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
      sigImgDataUrl,
      label: String(label || '').trim().slice(0, 40),
      createdAt: Date.now()
    });
    localStorage.setItem(SAVED_PDF_SIGS_KEY, JSON.stringify(cur.slice(0, 8)));
  } catch (err) {
    console.warn('Could not persist saved signature:', err);
  }
}

function removeSavedPdfSignature(id) {
  try {
    const next = loadSavedPdfSignatures().filter(x => x.id !== id);
    localStorage.setItem(SAVED_PDF_SIGS_KEY, JSON.stringify(next));
  } catch (_) {}
}

async function computePdfBytesSha256Hex(u8) {
  if (!u8 || !u8.length) return '0'.repeat(64);
  try {
    if (window.crypto && window.crypto.subtle && typeof window.crypto.subtle.digest === 'function') {
      const buf = u8.byteOffset === 0 && u8.byteLength === u8.buffer.byteLength
        ? u8.buffer
        : u8.slice().buffer;
      const digest = await window.crypto.subtle.digest('SHA-256', buf);
      return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
    }
  } catch (_) {}
  // Fallback deterministic digest for non-secure HTTP contexts where subtle crypto is unavailable
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < u8.length; i++) {
    h1 = Math.imul(h1 ^ u8[i], 0x01000193) >>> 0;
    h2 = Math.imul(h2 ^ ((u8[i] + i) & 0xff), 0x85ebca6b) >>> 0;
  }
  return (h1.toString(16).padStart(8, '0') + h2.toString(16).padStart(8, '0')).repeat(4);
}

function autoCropSignatureCanvas(srcCanvas, padPx = 10) {
  if (!srcCanvas || !srcCanvas.width || !srcCanvas.height) return null;
  const w = srcCanvas.width, hPx = srcCanvas.height;
  const ctx = srcCanvas.getContext('2d', { willReadFrequently: true });
  const data = ctx.getImageData(0, 0, w, hPx).data;
  let minX = w, minY = hPx, maxX = -1, maxY = -1;
  for (let y = 0; y < hPx; y++) {
    for (let x = 0; x < w; x++) {
      const a = data[(y * w + x) * 4 + 3];
      if (a > 14) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < minX || maxY < minY) return null;
  const x0 = Math.max(0, minX - padPx);
  const y0 = Math.max(0, minY - padPx);
  const cw = Math.min(w - x0, (maxX - minX + 1) + padPx * 2);
  const ch = Math.min(hPx - y0, (maxY - minY + 1) + padPx * 2);
  const out = document.createElement('canvas');
  out.width = Math.max(16, cw);
  out.height = Math.max(16, ch);
  out.getContext('2d').drawImage(srcCanvas, x0, y0, cw, ch, 0, 0, cw, ch);
  return {
    dataUrl: out.toDataURL('image/png'),
    width: out.width,
    height: out.height,
    aspect: out.width / Math.max(1, out.height)
  };
}

async function processUploadedSignatureImage(file, removeWhiteBg) {
  assertFileSize(file);
  const url = URL.createObjectURL(file);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const maxDim = 1100;
    const ratio = Math.min(1, maxDim / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
    const cw = Math.max(16, Math.round((img.naturalWidth || 300) * ratio));
    const ch = Math.max(16, Math.round((img.naturalHeight || 150) * ratio));
    const cvs = document.createElement('canvas');
    cvs.width = cw;
    cvs.height = ch;
    const ctx = cvs.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(img, 0, 0, cw, ch);
    if (removeWhiteBg) {
      const imgData = ctx.getImageData(0, 0, cw, ch);
      const d = imgData.data;
      for (let i = 0; i < d.length; i += 4) {
        const r = d[i], g = d[i + 1], b = d[i + 2], a = d[i + 3];
        if (a === 0) continue;
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        const sat = Math.max(r, g, b) - Math.min(r, g, b);
        if (lum > 236 && sat < 26) {
          d[i + 3] = 0;
        } else if (lum > 192 && sat < 38) {
          d[i + 3] = Math.max(0, Math.min(a, Math.round(((236 - lum) / 44) * a)));
        }
      }
      ctx.putImageData(imgData, 0, 0);
    }
    return autoCropSignatureCanvas(cvs, 8) || {
      dataUrl: cvs.toDataURL('image/png'),
      width: cw,
      height: ch,
      aspect: cw / Math.max(1, ch)
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function composeSignatureStampDataUrl(rawSigDataUrl, opts) {
  const img = new Image();
  img.src = rawSigDataUrl;
  await img.decode();
  const rawW = img.naturalWidth || img.width || 300;
  const rawH = img.naturalHeight || img.height || 110;
  const rawAspect = rawW / Math.max(1, rawH);
  if (!opts || !opts.showStamp) {
    return { dataUrl: rawSigDataUrl, aspect: rawAspect };
  }

  const signerName = String(opts.signerName || '').trim();
  const signerReason = String(opts.signerReason || '').trim();
  const signedAt = String(opts.signedAt || '').trim();
  const shaShort = String(opts.sha256Hex || '').slice(0, 20);
  const isRtl = /[\u0590-\u08FF]/.test(signerName + ' ' + signerReason) || (!signerName && currentLanguage === 'ar');

  const W = 720;
  const maxSigH = 210;
  let drawW = Math.min(W - 48, Math.round(maxSigH * rawAspect));
  let drawH = Math.round(drawW / Math.max(0.15, rawAspect));
  if (drawH < 88) {
    drawH = 88;
    drawW = Math.min(W - 48, Math.round(drawH * rawAspect));
  }
  const hasSignerLine = !!(signerName || signerReason);
  const sealH = hasSignerLine ? 102 : 76;
  const H = drawH + 16 + sealH + 8;

  const cvs = document.createElement('canvas');
  cvs.width = W;
  cvs.height = H;
  const ctx = cvs.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  ctx.drawImage(img, Math.round((W - drawW) / 2), 4, drawW, drawH);

  const bx = 8, by = drawH + 14, bw = W - 16, bh = sealH, rad = 12;
  ctx.save();
  ctx.beginPath();
  ctx.moveTo(bx + rad, by);
  ctx.lineTo(bx + bw - rad, by);
  ctx.quadraticCurveTo(bx + bw, by, bx + bw, by + rad);
  ctx.lineTo(bx + bw, by + bh - rad);
  ctx.quadraticCurveTo(bx + bw, by + bh, bx + bw - rad, by + bh);
  ctx.lineTo(bx + rad, by + bh);
  ctx.quadraticCurveTo(bx, by + bh, bx, by + bh - rad);
  ctx.lineTo(bx, by + rad);
  ctx.quadraticCurveTo(bx, by, bx + rad, by);
  ctx.closePath();
  ctx.fillStyle = 'rgba(248, 250, 252, 0.94)';
  ctx.fill();
  ctx.lineWidth = 2;
  ctx.strokeStyle = '#cbd5e1';
  ctx.stroke();

  // Accent verification stripe
  ctx.fillStyle = '#059669';
  if (isRtl) ctx.fillRect(bx + bw - 8, by + 6, 5, bh - 12);
  else ctx.fillRect(bx + 3, by + 6, 5, bh - 12);

  const padX = isRtl ? (bx + bw - 22) : (bx + 22);
  ctx.direction = isRtl ? 'rtl' : 'ltr';
  ctx.textAlign = isRtl ? 'right' : 'left';
  ctx.textBaseline = 'top';

  const badgeTitle = isRtl ? '✔ معتمد وموقّع إلكترونياً' : '✔ Digitally Signed & Verified';
  ctx.font = '700 21px "Cairo","Segoe UI",Tahoma,Arial,sans-serif';
  ctx.fillStyle = '#047857';
  ctx.fillText(badgeTitle, padX, by + 11);

  let nextY = by + 39;
  if (hasSignerLine) {
    const line2 = [signerName, signerReason].filter(Boolean).join('  •  ');
    ctx.font = '700 20px "Cairo","Segoe UI",Tahoma,Arial,sans-serif';
    ctx.fillStyle = '#0f172a';
    ctx.fillText(line2, padX, nextY, bw - 40);
    nextY += 28;
  }

  ctx.direction = 'ltr';
  ctx.textAlign = isRtl ? 'right' : 'left';
  ctx.font = '600 16px ui-monospace,SFMono-Regular,Consolas,monospace';
  ctx.fillStyle = '#475569';
  ctx.fillText(`${signedAt}  |  SHA-256: ${shaShort}…`, padX, nextY, bw - 40);
  ctx.restore();

  return {
    dataUrl: cvs.toDataURL('image/png'),
    aspect: W / H
  };
}

async function insertPdfSignatureOnCurrentPage(sigPayload) {
  if (!curPdf.bytes || !sigPayload || !sigPayload.sigImgDataUrl) return false;
  const pageIdx = curPdf.pageIdx || 0;
  const curRot = curPdf.rotations[pageIdx] || 0;
  ensurePdfPageArrays(pageIdx);

  const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
  const page = await pdfDoc.getPage(pageIdx + 1);
  const vp = page.getViewport({ scale: 1, rotation: curRot });
  const pageAspect = (vp.width || 595) / Math.max(1, vp.height || 842);

  const composed = await composeSignatureStampDataUrl(sigPayload.sigImgDataUrl, sigPayload);
  const targetWPct = sigPayload.showStamp ? 32 : 24;
  const rawHPct = (targetWPct * pageAspect) / Math.max(0.2, composed.aspect || 2.2);
  const targetHPct = Math.max(4, Math.min(48, rawHPct));
  const existingCount = (curPdf.signatures[pageIdx] || []).length;
  const offsetPct = (existingCount % 4) * 3;

  curPdf.signatures[pageIdx].push({
    id: 'psig_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
    sigImgDataUrl: sigPayload.sigImgDataUrl,
    stampDataUrl: composed.dataUrl,
    aspect: composed.aspect || 2.2,
    xPct: Math.min(96 - targetWPct, 58 + offsetPct),
    yPct: Math.min(95 - targetHPct, 74 + offsetPct),
    wPct: targetWPct,
    hPct: targetHPct,
    baseRot: curRot,
    showStamp: !!sigPayload.showStamp,
    signerName: sigPayload.signerName || '',
    signerReason: sigPayload.signerReason || '',
    signedAt: sigPayload.signedAt || '',
    sha256Hex: sigPayload.sha256Hex || ''
  });

  curPdf.tool = 'none';
  curPdf.activeOrigEditRef = null;
  await renderPdfEditor();
  toast(tr('تم إدراج التوقيع بنجاح — اسحبه أو غيّر حجمه بحرية'));
  return true;
}

async function applyPdfSignatureToAllPages(sourceSig, sourcePageIdx) {
  if (!curPdf.bytes || !sourceSig || curPdf.pageCount <= 1) return false;
  const srcRot = curPdf.rotations[sourcePageIdx] || 0;
  const normBox = projectPdfBoxFromRot(sourceSig, sourceSig.baseRot ?? srcRot, srcRot);

  for (let p = 0; p < curPdf.pageCount; p++) {
    if (p === sourcePageIdx) continue;
    ensurePdfPageArrays(p);
    const pRot = curPdf.rotations[p] || 0;
    curPdf.signatures[p].push({
      ...sourceSig,
      id: 'psig_' + Date.now() + '_' + p + '_' + Math.random().toString(36).slice(2, 5),
      xPct: normBox.xPct,
      yPct: normBox.yPct,
      wPct: normBox.wPct,
      hPct: normBox.hPct,
      baseRot: pRot
    });
  }
  await renderPdfEditor();
  toast(tr('تم تعميم التوقيع على جميع الصفحات'));
  return true;
}

async function openPdfSignatureModal() {
  if (!curPdf.bytes) return;
  $('.pdf-sig-backdrop')?.remove();
  const sha256Hex = await computePdfBytesSha256Hex(curPdf.bytes);

  let activeTab = 'draw'; // 'draw' | 'type' | 'upload'
  let inkColor = '#0f172a';
  let strokeWidth = 3.5;
  let typeFontIdx = 0;
  let uploadedRawFile = null;
  let uploadedProcessed = null;

  const fontPresets = [
    { key: 'cursive', label: 'توقيع انسيابي', spec: 'italic 68px "Brush Script MT","Segoe Script","Snell Roundhand",cursive' },
    { key: 'serif', label: 'كلاسيكي رسمي', spec: 'italic 700 58px Georgia,"Times New Roman",serif' },
    { key: 'bold', label: 'عصري عريض', spec: '800 56px "Cairo","Segoe UI",Tahoma,sans-serif' },
    { key: 'mono', label: 'خط رفيع مائل', spec: 'italic 54px "Palatino Linotype","Book Antiqua",serif' }
  ];

  const close = () => {
    document.removeEventListener('keydown', onEsc, true);
    backdrop.remove();
  };
  const onEsc = (ev) => {
    if (ev.key === 'Escape') { ev.stopPropagation(); close(); }
  };
  document.addEventListener('keydown', onEsc, true);

  const backdrop = h('div', {
    class: 'md-modal-backdrop pdf-sig-backdrop',
    onclick: (e) => { if (e.target === backdrop) close(); }
  });

  const modal = h('div', { class: 'pdf-sig-modal', role: 'dialog', 'aria-modal': 'true' });

  // Saved signatures gallery
  const savedWrap = h('div', { style: { display: 'none', flexDirection: 'column', gap: '4px' } });
  const savedList = h('div', { class: 'pdf-sig-saved-list' });
  savedWrap.append(h('small', { class: 'muted', style: { fontWeight: '700' } }, tr('التواقيع المحفوظة محلياً:')), savedList);

  // Seal controls (declared early so saved-signature click can read them)
  const sealChk = h('input', { type: 'checkbox' });
  const signerNameIn = h('input', {
    type: 'text',
    placeholder: currentLanguage === 'en' ? 'e.g. Full Name' : 'مثال: الاسم الكامل',
    style: { width: '100%', border: '1px solid var(--line)', borderRadius: '8px', padding: '6px 9px', background: 'var(--card)', color: 'var(--ink)' }
  });
  const signerReasonIn = h('input', {
    type: 'text',
    placeholder: currentLanguage === 'en' ? 'e.g. Approved / General Manager' : 'مثال: معتمد / المدير العام',
    style: { width: '100%', border: '1px solid var(--line)', borderRadius: '8px', padding: '6px 9px', background: 'var(--card)', color: 'var(--ink)' }
  });

  const buildStampMeta = () => {
    const now = new Date();
    const pad2 = (n) => String(n).padStart(2, '0');
    const signedAt = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())} ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
    return {
      showStamp: !!sealChk.checked,
      signerName: signerNameIn.value.trim(),
      signerReason: signerReasonIn.value.trim(),
      signedAt,
      sha256Hex
    };
  };

  const renderSavedGallery = () => {
    const items = loadSavedPdfSignatures();
    savedList.replaceChildren();
    if (!items.length) {
      savedWrap.style.display = 'none';
      return;
    }
    savedWrap.style.display = 'flex';
    items.forEach((it) => {
      const card = h('div', {
        class: 'pdf-sig-saved-card',
        title: it.label || tr('إدراج التوقيع في الصفحة'),
        onclick: async () => {
          close();
          await insertPdfSignatureOnCurrentPage({
            sigImgDataUrl: it.sigImgDataUrl,
            ...buildStampMeta()
          });
        }
      },
        h('img', { src: it.sigImgDataUrl, alt: 'Saved signature' }),
        h('button', {
          type: 'button',
          class: 'pdf-sig-saved-del',
          title: tr('حذف التوقيع المحفوظ'),
          onclick: (ev) => {
            ev.stopPropagation();
            removeSavedPdfSignature(it.id);
            renderSavedGallery();
          }
        }, '✕')
      );
      savedList.append(card);
    });
  };
  renderSavedGallery();

  // Tabs
  const tabDrawBtn = h('button', { type: 'button', class: 'pdf-sig-tab active' }, tr('✍️ رسم التوقيع'));
  const tabTypeBtn = h('button', { type: 'button', class: 'pdf-sig-tab' }, tr('⌨️ كتابة التوقيع'));
  const tabUploadBtn = h('button', { type: 'button', class: 'pdf-sig-tab' }, tr('📤 رفع صورة / ختم'));

  // Draw panel
  const padCanvas = h('canvas', { class: 'pdf-sig-pad', width: 1040, height: 350 });
  const padCtx = padCanvas.getContext('2d', { willReadFrequently: true });
  let hasDrawnStrokes = false;
  let drawing = false;
  let lastPt = null;
  let midPt = null;

  const getPadPt = (ev) => {
    const r = padCanvas.getBoundingClientRect();
    return {
      x: ((ev.clientX - r.left) / (r.width || 1)) * padCanvas.width,
      y: ((ev.clientY - r.top) / (r.height || 1)) * padCanvas.height
    };
  };

  padCanvas.addEventListener('pointerdown', (ev) => {
    ev.preventDefault();
    drawing = true;
    hasDrawnStrokes = true;
    lastPt = getPadPt(ev);
    midPt = lastPt;
    padCtx.save();
    padCtx.fillStyle = inkColor;
    padCtx.beginPath();
    padCtx.arc(lastPt.x, lastPt.y, strokeWidth, 0, Math.PI * 2);
    padCtx.fill();
    padCtx.restore();
    padCanvas.setPointerCapture(ev.pointerId);
  });

  padCanvas.addEventListener('pointermove', (ev) => {
    if (!drawing || !lastPt) return;
    const pt = getPadPt(ev);
    const nextMid = { x: (lastPt.x + pt.x) / 2, y: (lastPt.y + pt.y) / 2 };
    padCtx.save();
    padCtx.strokeStyle = inkColor;
    padCtx.lineWidth = strokeWidth * 2;
    padCtx.lineCap = 'round';
    padCtx.lineJoin = 'round';
    padCtx.beginPath();
    padCtx.moveTo(midPt.x, midPt.y);
    padCtx.quadraticCurveTo(lastPt.x, lastPt.y, nextMid.x, nextMid.y);
    padCtx.stroke();
    padCtx.restore();
    lastPt = pt;
    midPt = nextMid;
  });

  const endPadDraw = () => { drawing = false; lastPt = null; midPt = null; };
  padCanvas.addEventListener('pointerup', endPadDraw);
  padCanvas.addEventListener('pointercancel', endPadDraw);

  const clearPad = () => {
    padCtx.clearRect(0, 0, padCanvas.width, padCanvas.height);
    hasDrawnStrokes = false;
  };

  // Shared color picker + stroke width
  const presetColors = ['#0f172a', '#1d4ed8', '#b91c1c', '#047857'];
  const swatchBtns = [];
  const customColIn = h('input', {
    type: 'color',
    value: inkColor,
    style: { width: '28px', height: '26px', border: '1px solid var(--line)', borderRadius: '6px', padding: '1px', cursor: 'pointer' }
  });
  const updateInkColor = (col) => {
    inkColor = col;
    customColIn.value = col;
    swatchBtns.forEach((b, idx) => b.classList.toggle('active', presetColors[idx].toLowerCase() === col.toLowerCase()));
    renderTypedPreview();
  };
  presetColors.forEach((c, idx) => {
    const sb = h('button', {
      type: 'button',
      class: 'pdf-sig-swatch' + (idx === 0 ? ' active' : ''),
      style: { background: c },
      onclick: () => updateInkColor(c)
    });
    swatchBtns.push(sb);
  });
  customColIn.oninput = () => updateInkColor(customColIn.value);

  const widthSel = h('select', { class: 'bar-select' },
    h('option', { value: '2.2' }, tr('رفيع')),
    h('option', { value: '3.5', selected: true }, tr('متوسط')),
    h('option', { value: '5.5' }, tr('عريض'))
  );
  widthSel.onchange = () => { strokeWidth = parseFloat(widthSel.value) || 3.5; };

  const drawPanel = h('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    h('div', { class: 'pdf-sig-pad-wrap' },
      padCanvas,
      h('div', { class: 'pdf-sig-pad-hint' }, tr('ارسم توقيعك بالماوس أو اللمس هنا'))
    ),
    h('div', { class: 'row', style: { justifyContent: 'space-between' } },
      h('div', { class: 'row' },
        h('span', { class: 'bar-label' }, tr('لون الحبر:')),
        h('div', { class: 'pdf-sig-swatches' }, ...swatchBtns, customColIn),
        h('span', { class: 'bar-label', style: { marginInlineStart: '6px' } }, tr('سمك الحبر:')),
        widthSel
      ),
      btn('🧹 مسح التوقيع', clearPad)
    )
  );

  // Type panel
  const typeInput = h('input', {
    type: 'text',
    placeholder: tr('اكتب اسمك للتوقيع…'),
    style: { width: '100%', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px 11px', fontSize: '1rem', background: 'var(--soft)', color: 'var(--ink)' }
  });
  const typeCanvas = h('canvas', { class: 'pdf-sig-pad', width: 1040, height: 300, style: { cursor: 'default' } });
  const renderTypedPreview = () => {
    const tCtx = typeCanvas.getContext('2d');
    tCtx.clearRect(0, 0, typeCanvas.width, typeCanvas.height);
    const val = typeInput.value.trim();
    if (!val) return;
    const isRtl = /[\u0590-\u08FF]/.test(val);
    const preset = fontPresets[typeFontIdx] || fontPresets[0];
    tCtx.save();
    tCtx.font = isRtl && preset.key === 'cursive' ? 'italic 700 64px "Cairo","Amiri","Traditional Arabic",serif' : preset.spec;
    tCtx.fillStyle = inkColor;
    tCtx.direction = isRtl ? 'rtl' : 'ltr';
    tCtx.textAlign = 'center';
    tCtx.textBaseline = 'middle';
    tCtx.fillText(val, typeCanvas.width / 2, typeCanvas.height / 2, typeCanvas.width - 60);
    tCtx.restore();
  };
  typeInput.addEventListener('input', () => {
    renderTypedPreview();
    if (!signerNameIn.value.trim()) signerNameIn.placeholder = typeInput.value.trim() || signerNameIn.placeholder;
  });

  const styleBtnsWrap = h('div', { class: 'row' });
  const rebuildStyleBtns = () => {
    styleBtnsWrap.replaceChildren();
    fontPresets.forEach((fp, idx) => {
      styleBtnsWrap.append(btn(fp.label, () => {
        typeFontIdx = idx;
        rebuildStyleBtns();
        renderTypedPreview();
      }, idx === typeFontIdx ? 'on' : ''));
    });
  };
  rebuildStyleBtns();

  const typeColorRow = h('div', { class: 'row' },
    h('span', { class: 'bar-label' }, tr('نمط الخط:')),
    styleBtnsWrap
  );

  const typePanel = h('div', { hidden: true, style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    typeInput,
    typeColorRow,
    h('div', { class: 'pdf-sig-pad-wrap', style: { height: '140px' } }, typeCanvas)
  );

  // Upload panel
  const upFileIn = h('input', { type: 'file', accept: 'image/*', hidden: true });
  const removeBgChk = h('input', { type: 'checkbox', checked: true });
  const uploadPreviewImg = h('img', { style: { maxWidth: '92%', maxHeight: '145px', objectFit: 'contain', display: 'none' }, alt: 'Uploaded signature preview' });
  const uploadEmptyHint = h('span', { class: 'muted', style: { fontSize: '.88rem' } }, tr('اختر صورة توقيع أو ختم رسمي'));

  const refreshUploadPreview = async () => {
    if (!uploadedRawFile) return;
    try {
      uploadedProcessed = await processUploadedSignatureImage(uploadedRawFile, removeBgChk.checked);
      if (uploadedProcessed?.dataUrl) {
        uploadPreviewImg.src = uploadedProcessed.dataUrl;
        uploadPreviewImg.style.display = 'block';
        uploadEmptyHint.style.display = 'none';
      }
    } catch (err) {
      toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة');
    }
  };
  upFileIn.onchange = () => {
    const f = upFileIn.files?.[0];
    if (!f) return;
    uploadedRawFile = f;
    refreshUploadPreview();
  };
  removeBgChk.onchange = () => refreshUploadPreview();

  const uploadPanel = h('div', { hidden: true, style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    h('div', { class: 'row', style: { justifyContent: 'space-between' } },
      btn('اختر صورة توقيع أو ختم رسمي', () => upFileIn.click(), 'pri'),
      h('label', { style: { display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '.86rem' } },
        removeBgChk, tr('تفريغ الخلفية البيضاء تلقائياً')
      )
    ),
    upFileIn,
    h('div', {
      class: 'pdf-sig-pad-wrap',
      style: { height: '155px', cursor: 'pointer', backgroundImage: 'repeating-conic-gradient(#f1f5f9 0% 25%, #ffffff 0% 50%)', backgroundSize: '16px 16px' },
      onclick: () => upFileIn.click()
    }, uploadPreviewImg, uploadEmptyHint)
  );

  const switchTab = (tab) => {
    activeTab = tab;
    tabDrawBtn.classList.toggle('active', tab === 'draw');
    tabTypeBtn.classList.toggle('active', tab === 'type');
    tabUploadBtn.classList.toggle('active', tab === 'upload');
    drawPanel.hidden = tab !== 'draw';
    typePanel.hidden = tab !== 'type';
    uploadPanel.hidden = tab !== 'upload';
    if (tab === 'type') setTimeout(() => typeInput.focus(), 20);
  };
  tabDrawBtn.onclick = () => switchTab('draw');
  tabTypeBtn.onclick = () => switchTab('type');
  tabUploadBtn.onclick = () => switchTab('upload');

  // Verification seal section
  const sealFieldsWrap = h('div', { hidden: true, style: { display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '4px' } },
    h('div', { class: 'fld-grid' },
      h('div', null, h('label', { style: { display: 'block', fontSize: '.78rem', color: 'var(--mut)', marginBottom: '2px' } }, tr('اسم الموقّع:')), signerNameIn),
      h('div', null, h('label', { style: { display: 'block', fontSize: '.78rem', color: 'var(--mut)', marginBottom: '2px' } }, tr('الصفة أو السبب (اختياري):')), signerReasonIn)
    ),
    h('div', { class: 'pdf-sig-hash', title: 'SHA-256: ' + sha256Hex }, `🔒 SHA-256: ${sha256Hex}`)
  );
  sealChk.onchange = () => {
    sealFieldsWrap.hidden = !sealChk.checked;
    if (sealChk.checked && !signerNameIn.value.trim() && typeInput.value.trim()) {
      signerNameIn.value = typeInput.value.trim();
    }
  };

  const sealBox = h('div', { class: 'pdf-sig-seal-box' },
    h('label', { style: { display: 'inline-flex', alignItems: 'center', gap: '7px', fontWeight: '700', cursor: 'pointer' } },
      sealChk, tr('إرفاق ختم التوثيق الرقمي (الاسم + التاريخ + SHA-256)')
    ),
    sealFieldsWrap
  );

  const saveForLaterChk = h('input', { type: 'checkbox', checked: true });

  const extractCurrentRawSignature = () => {
    if (activeTab === 'draw') {
      if (!hasDrawnStrokes) return null;
      return autoCropSignatureCanvas(padCanvas, 10);
    }
    if (activeTab === 'type') {
      if (!typeInput.value.trim()) return null;
      renderTypedPreview();
      return autoCropSignatureCanvas(typeCanvas, 10);
    }
    if (activeTab === 'upload') {
      return uploadedProcessed;
    }
    return null;
  };

  const confirmBtn = btn('إدراج التوقيع في الصفحة', async () => {
    const raw = extractCurrentRawSignature();
    if (!raw || !raw.dataUrl) {
      toast(tr('يرجى رسم التوقيع أو كتابته أو رفع صورة أولاً'));
      return;
    }
    if (saveForLaterChk.checked) {
      savePdfSignatureLocally(raw.dataUrl, signerNameIn.value.trim() || typeInput.value.trim() || '');
    }
    const meta = buildStampMeta();
    if (meta.showStamp && !meta.signerName && typeInput.value.trim()) {
      meta.signerName = typeInput.value.trim();
    }
    close();
    await insertPdfSignatureOnCurrentPage({
      sigImgDataUrl: raw.dataUrl,
      ...meta
    });
  }, 'pri');

  modal.append(
    h('div', { class: 'row', style: { justifyContent: 'space-between' } },
      h('strong', { style: { fontSize: '1.02rem' } }, '✒️ ' + tr('إضافة توقيع رقمي / إلكتروني')),
      h('button', { type: 'button', class: 'btn', style: { padding: '2px 8px' }, onclick: close }, '✕')
    ),
    savedWrap,
    h('div', { class: 'pdf-sig-tabs' }, tabDrawBtn, tabTypeBtn, tabUploadBtn),
    drawPanel,
    typePanel,
    uploadPanel,
    sealBox,
    h('div', { class: 'row', style: { justifyContent: 'space-between', marginTop: '2px' } },
      h('label', { style: { display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '.84rem', cursor: 'pointer' } },
        saveForLaterChk, tr('حفظ هذا التوقيع للاستخدام لاحقاً')
      ),
      h('div', { class: 'row' },
        btn('إلغاء', close),
        confirmBtn
      )
    )
  );

  backdrop.append(modal);
  document.body.append(backdrop);
}

function renderPdfSignatureLayer(sigLayer, pageIdx, curRot, pw, ph, onAnyChange) {
  if (!sigLayer) return;
  sigLayer.innerHTML = '';
  const items = curPdf.signatures[pageIdx] || [];
  if (!items.length) return;

  items.forEach((item) => {
    const proj = projectPdfBoxFromRot(item, item.baseRot ?? curRot, curRot);
    const box = h('div', {
      class: 'pdf-sig-item',
      style: {
        left: proj.xPct + '%',
        top: proj.yPct + '%',
        width: proj.wPct + '%',
        height: proj.hPct + '%'
      }
    });

    const imgEl = h('img', {
      src: item.stampDataUrl || item.sigImgDataUrl,
      alt: 'Digital Signature',
      draggable: false
    });

    const actions = h('div', { class: 'pdf-sig-actions' });
    if ((curPdf.pageCount || 1) > 1) {
      actions.append(h('button', {
        type: 'button',
        class: 'pdf-sig-act-btn',
        title: tr('📄 تعميم على كل الصفحات'),
        onclick: (ev) => {
          ev.stopPropagation();
          applyPdfSignatureToAllPages(item, pageIdx);
        }
      }, tr('📄 تعميم على كل الصفحات')));
    }
    actions.append(h('button', {
      type: 'button',
      class: 'pdf-sig-act-btn danger',
      title: tr('تم حذف التوقيع'),
      onclick: (ev) => {
        ev.stopPropagation();
        curPdf.signatures[pageIdx] = (curPdf.signatures[pageIdx] || []).filter(s => s.id !== item.id);
        renderPdfSignatureLayer(sigLayer, pageIdx, curRot, pw, ph, onAnyChange);
        if (onAnyChange) onAnyChange();
        toast(tr('تم حذف التوقيع'));
      }
    }, '✕'));

    const resizeHandle = h('div', { class: 'pdf-sig-resize' });

    resizeHandle.addEventListener('pointerdown', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      const startX = ev.clientX;
      const startWPct = proj.wPct;
      const layerRect = sigLayer.getBoundingClientRect();
      const aspect = Math.max(0.25, item.aspect || ((proj.wPct * (pw || 600)) / Math.max(1, proj.hPct * (ph || 800))));
      let resized = false;

      const onMove = (mv) => {
        const dx = mv.clientX - startX;
        if (!resized && Math.abs(dx) < 2) return;
        resized = true;
        const nextWPct = Math.max(6, Math.min(88, startWPct + (dx / Math.max(1, layerRect.width)) * 100));
        const nextWPx = (nextWPct / 100) * (layerRect.width || pw || 600);
        const nextHPx = nextWPx / aspect;
        const nextHPct = Math.max(3, Math.min(88, (nextHPx / Math.max(1, layerRect.height || ph || 800)) * 100));
        proj.wPct = nextWPct;
        proj.hPct = nextHPct;
        proj.xPct = Math.max(0, Math.min(100 - proj.wPct, proj.xPct));
        proj.yPct = Math.max(0, Math.min(100 - proj.hPct, proj.yPct));
        item.xPct = proj.xPct;
        item.yPct = proj.yPct;
        item.wPct = proj.wPct;
        item.hPct = proj.hPct;
        item.baseRot = curRot;
        box.style.left = proj.xPct + '%';
        box.style.top = proj.yPct + '%';
        box.style.width = proj.wPct + '%';
        box.style.height = proj.hPct + '%';
      };

      const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        if (resized && onAnyChange) onAnyChange();
      };
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    });

    box.addEventListener('pointerdown', (ev) => {
      if (ev.target.closest('.pdf-sig-actions') || ev.target === resizeHandle) return;
      if (curPdf.tool !== 'none') return;
      ev.preventDefault();
      ev.stopPropagation();
      $$('.pdf-sig-item.active', sigLayer).forEach(el => el.classList.remove('active'));
      box.classList.add('active');

      const startX = ev.clientX;
      const startY = ev.clientY;
      const origX = proj.xPct;
      const origY = proj.yPct;
      const layerRect = sigLayer.getBoundingClientRect();
      let dragged = false;

      const onMove = (mv) => {
        const dx = mv.clientX - startX;
        const dy = mv.clientY - startY;
        if (!dragged && Math.hypot(dx, dy) < 3) return;
        dragged = true;
        proj.xPct = Math.max(0, Math.min(100 - proj.wPct, origX + (dx / Math.max(1, layerRect.width)) * 100));
        proj.yPct = Math.max(0, Math.min(100 - proj.hPct, origY + (dy / Math.max(1, layerRect.height)) * 100));
        item.xPct = proj.xPct;
        item.yPct = proj.yPct;
        item.wPct = proj.wPct;
        item.hPct = proj.hPct;
        item.baseRot = curRot;
        box.style.left = proj.xPct + '%';
        box.style.top = proj.yPct + '%';
      };

      const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        if (dragged && onAnyChange) onAnyChange();
      };
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
    });

    box.append(imgEl, actions, resizeHandle);
    sigLayer.append(box);
  });
}

// يمنع عمليات العرض القديمة من استبدال واجهة أحدث عند التنقل السريع بين الصفحات.
let pdfRenderGeneration = 0;
// موضع القائمة محفوظ خارج DOM حتى لا نفقده عند إعادة إنشاء المحرر.
let pdfThumbScrollTop = 0;

// معرّف فريد لكل نسخة من بايتات الـ PDF، نستخدمه لنعرف متى نحتاج نعيد بناء المصغرات فعلاً.
const pdfBytesIds = new WeakMap();
let pdfBytesIdCounter = 0;
function getPdfBytesId(bytes) {
  if (!bytes || typeof bytes !== 'object') return 'none';
  if (!pdfBytesIds.has(bytes)) pdfBytesIds.set(bytes, ++pdfBytesIdCounter);
  return String(pdfBytesIds.get(bytes));
}

async function renderPdfEditor() {
  syncEditorTopBarForKind();
  const appEl = $('#app');
  if (!appEl || !curPdf.bytes) return;

  const oldThumbs = appEl.querySelector('.pdf-thumbs');
  if (oldThumbs) pdfThumbScrollTop = oldThumbs.scrollTop;
  let previousThumbScroll = pdfThumbScrollTop;
  const renderGeneration = ++pdfRenderGeneration;
  const previousWindowScroll = window.scrollY || window.pageYOffset || 0;

  const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
  if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;
  // Keep the currently displayed page on screen until the next page is ready.
  // This avoids the blank editor flash during navigation.
  const total = pdfDoc.numPages;
  curPdf.pageCount = total;
  if (curPdf.pageIdx >= total) curPdf.pageIdx = Math.max(0, total - 1);
  if (curPdf.pageIdx < 0) curPdf.pageIdx = 0;
  while (curPdf.rotations.length < total) curPdf.rotations.push(0);

  const curIdx = curPdf.pageIdx;
  const curRot = curPdf.rotations[curIdx] || 0;
  ensurePdfPageArrays(curIdx);

  // الحل الجذري: لو الصفحات نفسها ما تغيّرت (نفس الملف، نفس العدد، نفس التدوير)
  // ما نلمس عمود المصغرات أبداً، نغيّر بس الصفحة الرئيسية ونحرّك علامة "الصفحة الحالية".
  // أي إعادة إلحاق للعمود في الـ DOM تصفّر السكرول، وهذا كان سبب الرجوع لفوق.
  const thumbsSig = [getPdfBytesId(curPdf.bytes), total, currentLanguage, curPdf.rotations.slice(0, total).join(',')].join('|');
  const oldRoot = appEl.querySelector(':scope > .pdf-editor');
  const reuseThumbs = !!(oldRoot && oldThumbs && oldThumbs.parentNode === oldRoot &&
    oldThumbs.dataset.sig === thumbsSig && oldThumbs.querySelectorAll('.pdf-thumb').length === total);

  const root = h('div', { class: 'pdf-editor' });
  const thumbsCol = h('div', { class: 'pdf-thumbs' });
  const mainCol = h('div', { class: 'pdf-main' });
  const toolbar = h('div', { class: 'pdf-toolbar' });
  const stageWrap = h('div', { class: 'pdf-stage-wrap' });

  const mergeInput = $('#pdfMergeIn');
  const bMerge = btn(currentLanguage === 'en' ? '➕ Merge PDF' : '➕ دمج ملف PDF', () => {
    if (!mergeInput) return;
    mergeInput.value = '';
    mergeInput.onchange = () => {
      const picked = Array.from(mergeInput.files || []);
      if (!picked.length) return;
      if (picked.length === 1) { mergePdfFiles(picked); return; }
      openMergeOrderDialog(picked).then(ordered => { if (ordered && ordered.length) mergePdfFiles(ordered); });
    };
    mergeInput.click();
  }, 'pri');

  const bDelPage = btn(
    currentLanguage === 'en' ? '🗑 Delete Page' : '🗑 حذف الصفحة',
    () => deletePdfPageAt(curIdx),
    'danger',
    { disabled: total <= 1 }
  );
  const bBlankPage = btn(tr('📄 صفحة فارغة'), () => insertBlankPdfPageAfterCurrent());
  const bExtractPages = btn(tr('✂️ استخراج صفحات'), () => extractCurrentPdfPages());

  const bRotLeft = btn(currentLanguage === 'en' ? '↺ Rotate Left' : '↺ تدوير يسار', () => rotateCurrentPdfPage(-90));
  const bRotRight = btn(currentLanguage === 'en' ? '↻ Rotate Right' : '↻ تدوير يمين', () => rotateCurrentPdfPage(90));

  const bEditOrigTool = btn(
    tr('✏️ تعديل نص الملف'),
    () => {
      curPdf.tool = curPdf.tool === 'edit-text' ? 'none' : 'edit-text';
      curPdf.activeOrigEditRef = null;
      renderPdfEditor();
    },
    curPdf.tool === 'edit-text' ? 'on' : ''
  );

  const bTextTool = btn(
    currentLanguage === 'en' ? '🔤 Add Text' : '🔤 إضافة نص',
    () => {
      curPdf.tool = curPdf.tool === 'text' ? 'none' : 'text';
      curPdf.activeOrigEditRef = null;
      renderPdfEditor();
    },
    curPdf.tool === 'text' ? 'on' : ''
  );

  const applyActiveOrigStyle = (nextColor, nextSize) => {
    const ref = curPdf.activeOrigEditRef;
    if (curPdf.tool === 'edit-text' && ref && ref.pageIdx === curIdx && typeof ref.applyStyleUpdate === 'function') {
      ref.applyStyleUpdate(nextColor, nextSize);
    }
  };

  const txtColorIn = h('input', {
    type: 'color',
    value: curPdf.textColor,
    title: currentLanguage === 'en' ? 'Text color' : 'لون النص',
    oninput: () => { curPdf.textColor = txtColorIn.value; applyActiveOrigStyle(txtColorIn.value, null); },
    onchange: () => { curPdf.textColor = txtColorIn.value; applyActiveOrigStyle(txtColorIn.value, null); }
  });

  const txtSizeIn = h('input', {
    type: 'number',
    min: 8,
    max: 72,
    value: curPdf.textSize,
    title: currentLanguage === 'en' ? 'Text size (px)' : 'حجم النص (px)',
    onchange: () => {
      const sz = Math.max(8, Math.min(72, +txtSizeIn.value || 18));
      curPdf.textSize = sz;
      applyActiveOrigStyle(null, sz);
    }
  });

  const bPenTool = btn(
    currentLanguage === 'en' ? '✏️ Pen' : '✏️ قلم رسم',
    () => {
      curPdf.tool = curPdf.tool === 'pen' ? 'none' : 'pen';
      curPdf.activeOrigEditRef = null;
      renderPdfEditor();
    },
    curPdf.tool === 'pen' ? 'on' : ''
  );

  const bHighlightTool = btn(
    tr('🖍 تظليل'),
    () => {
      curPdf.tool = curPdf.tool === 'highlight' ? 'none' : 'highlight';
      curPdf.activeOrigEditRef = null;
      renderPdfEditor();
    },
    curPdf.tool === 'highlight' ? 'on' : ''
  );

  const bWhiteoutTool = btn(
    tr('▭ طمس/تغطية'),
    () => {
      curPdf.tool = curPdf.tool === 'whiteout' ? 'none' : 'whiteout';
      curPdf.activeOrigEditRef = null;
      renderPdfEditor();
    },
    curPdf.tool === 'whiteout' ? 'on' : ''
  );

  const bAddImgPdf = btn(
    tr('🖼 صورة PDF'),
    () => triggerPdfImageInsert()
  );

  const penColorIn = h('input', {
    type: 'color',
    value: curPdf.penColor,
    title: currentLanguage === 'en' ? 'Pen color' : 'لون القلم',
    onchange: () => { curPdf.penColor = penColorIn.value; }
  });

  const penSizeIn = h('input', {
    type: 'number',
    min: 1,
    max: 24,
    value: curPdf.penSize,
    title: currentLanguage === 'en' ? 'Pen width (px)' : 'سمك القلم (px)',
    onchange: () => { curPdf.penSize = Math.max(1, Math.min(24, +penSizeIn.value || 3)); }
  });

  const bClearDraw = btn(
    currentLanguage === 'en' ? '🧹 Clear Drawing' : '🧹 مسح الرسم',
    () => {
      curPdf.drawings[curIdx] = [];
      renderPdfEditor();
      toast(currentLanguage === 'en' ? 'Drawing cleared' : 'تم مسح الرسم');
    }
  );

  const bSignPdf = btn(
    tr('✒️ توقيع رقمي'),
    () => openPdfSignatureModal(),
    'pri'
  );

  const zoomSel = h('select', { class: 'zoom-select' },
    h('option', { value: 'fit', selected: curPdf.zoom === 'fit' }, tr('ملائمة العرض')),
    ...[0.75, 1, 1.15, 1.35, 1.6].map(z => h('option', { value: String(z), selected: typeof curPdf.zoom === 'number' && Math.abs(curPdf.zoom - z) < 0.02 }, Math.round(z * 100) + '%'))
  );
  zoomSel.onchange = () => {
    curPdf.zoom = zoomSel.value === 'fit' ? 'fit' : (parseFloat(zoomSel.value) || 1.15);
    renderPdfEditor();
  };

  const bExportPng = btn(
    currentLanguage === 'en' ? '🖼 Export PNG' : '🖼 تصدير PNG',
    () => exportCurrentPdfPagePng(curIdx)
  );

  const bToDocx = btn(
    tr('🔄 تحويل إلى Word'),
    () => convertCurrentPdfToDocx()
  );

  const bPrev = btn('◀', () => { if (curPdf.pageIdx > 0) { curPdf.pageIdx--; renderPdfEditor(); } }, '', { disabled: curIdx <= 0 });
  const bNext = btn('▶', () => { if (curPdf.pageIdx < total - 1) { curPdf.pageIdx++; renderPdfEditor(); } }, '', { disabled: curIdx >= total - 1 });
  const pageJumpBadge = h('button', {
    type: 'button',
    class: 'btn',
    title: tr('انقر للانتقال السريع إلى صفحة'),
    style: { fontWeight: '700', fontSize: '.84rem', padding: '4px 9px' },
    onclick: () => {
      const entered = prompt(tr('أدخل رقم الصفحة للانتقال إليها:') + ` (1-${total})`, String(curIdx + 1));
      if (entered == null) return;
      const target = parseInt(entered.trim(), 10);
      if (!isNaN(target) && target >= 1 && target <= total && target - 1 !== curPdf.pageIdx) {
        curPdf.pageIdx = target - 1;
        renderPdfEditor();
      }
    }
  }, `${curIdx + 1} / ${total}`);

  // Grouped cleanly by topic (Pages → Rotate → Text → Annotate → Insert/Sign → View/Export → Nav) without separators
  toolbar.append(
    bMerge, bBlankPage, bExtractPages, bDelPage,
    bRotLeft, bRotRight,
    bEditOrigTool, bTextTool, txtColorIn, txtSizeIn,
    bHighlightTool, bWhiteoutTool, bPenTool, penColorIn, penSizeIn, bClearDraw,
    bAddImgPdf, bSignPdf,
    zoomSel, bExportPng, bToDocx,
    h('span', { style: { flex: '1' } }),
    bPrev,
    pageJumpBadge,
    bNext
  );

  // Mount the editor shell immediately; page rasterization and thumbnails must not
  // hold the toolbar or the rest of the editor hostage.
  stageWrap.append(h('div', { class: 'pdf-loading-placeholder' }, currentLanguage === 'en' ? 'Loading page…' : 'جارٍ عرض الصفحة…'));
  mainCol.append(toolbar, stageWrap);
  root.append(thumbsCol, mainCol);

  const activePage = await pdfDoc.getPage(curIdx + 1);
  if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;
  let effectiveZoom = typeof curPdf.zoom === 'number' ? curPdf.zoom : 1.15;
  if (curPdf.zoom === 'fit') {
    const baseVp = activePage.getViewport({ scale: 1, rotation: curRot });
    const availW = Math.max(220, (window.innerWidth <= 960 ? window.innerWidth - 36 : (appEl.clientWidth || window.innerWidth) - 255));
    effectiveZoom = Math.max(0.35, Math.min(1.6, Math.round((availW / Math.max(1, baseVp.width)) * 100) / 100));
  }
  curPdf._effectiveZoom = effectiveZoom;
  const viewport = activePage.getViewport({ scale: effectiveZoom, rotation: curRot });
  const pw = Math.round(viewport.width);
  const ph = Math.round(viewport.height);

  const pageBox = h('div', {
    class: 'pdf-page-box' + (curPdf.tool !== 'none' ? ' mode-' + curPdf.tool : ''),
    style: { width: pw + 'px', height: ph + 'px' }
  });

  const renderCanvas = h('canvas', {
    class: 'pdf-render-canvas',
    width: pw,
    height: ph
  });
  await activePage.render({ canvasContext: renderCanvas.getContext('2d'), viewport }).promise;
  if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;

  const origLayer = h('div', {
    class: 'pdf-orig-layer' + (curPdf.tool === 'edit-text' ? ' active-edit-text' : '')
  });
  const hasSavedOrigEdits = Object.keys(curPdf.textEdits[curIdx] || {}).length > 0;
  if (curPdf.tool === 'edit-text' || hasSavedOrigEdits) {
    const blocks = curPdf.tool === 'edit-text'
      ? await extractPdfPageTextBlocks(activePage, curRot, renderCanvas)
      : [];
    if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;
    if (curPdf.tool === 'edit-text' && !blocks.length && !hasSavedOrigEdits) {
      toast(tr('هذه الصفحة مصوّرة ولا تحتوي على نص أصلي قابل للتعديل'));
    }
    renderPdfOrigTextLayer(origLayer, blocks, curIdx, curRot, {
      syncToolbarInputs: (col, sz) => {
        if (col) { txtColorIn.value = col; curPdf.textColor = col; }
        if (sz) { txtSizeIn.value = String(sz); curPdf.textSize = sz; }
      },
      onInputChange: () => {
        origLayer.dispatchEvent(new Event('origeditinput', { bubbles: true }));
      },
      onCommitChange: () => {
        origLayer.dispatchEvent(new Event('origeditcommit', { bubbles: true }));
      }
    });
  }

  const isDrawMode = ['pen', 'highlight', 'whiteout'].includes(curPdf.tool);
  const drawCanvas = h('canvas', {
    class: 'pdf-draw-canvas' + (isDrawMode ? ' active-' + curPdf.tool : ''),
    width: pw,
    height: ph
  });
  redrawPdfDrawCanvas(drawCanvas, curIdx);

  const textLayer = h('div', {
    class: 'pdf-text-layer' + (curPdf.tool === 'text' ? ' active-text' : '')
  });
  renderPdfTextOverlays(textLayer, curIdx);

  const sigLayer = h('div', { class: 'pdf-sig-layer' });
  renderPdfSignatureLayer(sigLayer, curIdx, curRot, pw, ph, () => {
    sigLayer.dispatchEvent(new Event('sigmutate', { bubbles: true }));
  });

  // رسم حر بالقلم / التظليل الشفاف / مستطيل التغطية على الصفحة
  let activeStroke = null;
  drawCanvas.addEventListener('pointerdown', (ev) => {
    if (!['pen', 'highlight', 'whiteout'].includes(curPdf.tool)) return;
    ev.preventDefault();
    const r = drawCanvas.getBoundingClientRect();
    const xNorm = Math.max(0, Math.min(1, (ev.clientX - r.left) / (r.width || 1)));
    const yNorm = Math.max(0, Math.min(1, (ev.clientY - r.top) / (r.height || 1)));
    if (curPdf.tool === 'whiteout') {
      activeStroke = {
        mode: 'whiteout',
        pts: [{ x: xNorm, y: yNorm }, { x: xNorm, y: yNorm }],
        color: '#ffffff',
        size: 1
      };
    } else if (curPdf.tool === 'highlight') {
      activeStroke = {
        mode: 'highlight',
        pts: [{ x: xNorm, y: yNorm }],
        color: curPdf.highlightColor || '#facc15',
        size: 18
      };
    } else {
      activeStroke = {
        mode: 'pen',
        pts: [{ x: xNorm, y: yNorm }],
        color: penColorIn.value || curPdf.penColor,
        size: Math.max(1, Math.min(24, +penSizeIn.value || curPdf.penSize))
      };
    }
    ensurePdfPageArrays(curIdx);
    curPdf.drawings[curIdx].push(activeStroke);
    drawCanvas.setPointerCapture(ev.pointerId);
  });

  drawCanvas.addEventListener('pointermove', (ev) => {
    if (!activeStroke || !['pen', 'highlight', 'whiteout'].includes(curPdf.tool)) return;
    const r = drawCanvas.getBoundingClientRect();
    const xNorm = Math.max(0, Math.min(1, (ev.clientX - r.left) / (r.width || 1)));
    let yNorm = Math.max(0, Math.min(1, (ev.clientY - r.top) / (r.height || 1)));
    if (activeStroke.mode === 'whiteout') {
      activeStroke.pts[1] = { x: xNorm, y: yNorm };
    } else {
      if (activeStroke.mode === 'highlight' && Math.abs(yNorm - activeStroke.pts[0].y) < 0.012) {
        yNorm = activeStroke.pts[0].y;
      }
      activeStroke.pts.push({ x: xNorm, y: yNorm });
    }
    redrawPdfDrawCanvas(drawCanvas, curIdx);
  });

  const finishStroke = () => {
    activeStroke = null;
  };
  drawCanvas.addEventListener('pointerup', finishStroke);
  drawCanvas.addEventListener('pointercancel', finishStroke);

  // إضافة نص جديد عند النقر في طبقة النص
  textLayer.addEventListener('pointerdown', (ev) => {
    if (curPdf.tool !== 'text') return;
    if (ev.target !== textLayer) return;
    ev.preventDefault();
    const r = textLayer.getBoundingClientRect();
    const xPct = Math.max(0, Math.min(85, ((ev.clientX - r.left) / (r.width || 1)) * 100));
    const yPct = Math.max(0, Math.min(90, ((ev.clientY - r.top) / (r.height || 1)) * 100));
    const defaultTxt = currentLanguage === 'en' ? 'New text' : 'نص جديد';
    ensurePdfPageArrays(curIdx);
    const newItem = {
      id: 'ptxt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
      xPct,
      yPct,
      text: defaultTxt,
      color: txtColorIn.value || curPdf.textColor,
      size: Math.max(10, Math.min(72, +txtSizeIn.value || curPdf.textSize))
    };
    curPdf.texts[curIdx].push(newItem);
    renderPdfTextOverlays(textLayer, curIdx);
    const created = textLayer.lastElementChild;
    if (created && created.contentEditable === 'true') {
      created.focus();
      const sel = window.getSelection();
      if (sel && created.firstChild) {
        const range = document.createRange();
        range.selectNodeContents(created.firstChild);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    }
  });

  pageBox.addEventListener('pointerdown', (ev) => {
    if (!ev.target.closest('.pdf-sig-item')) {
      $$('.pdf-sig-item.active', sigLayer).forEach(el => el.classList.remove('active'));
    }
  });
  pageBox.append(renderCanvas, origLayer, drawCanvas, textLayer, sigLayer);
  if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;
  stageWrap.innerHTML = '';
  stageWrap.append(pageBox);
  if (renderGeneration !== pdfRenderGeneration || !appEl.isConnected) return;

  if (reuseThumbs) {
    // نبدّل الجزء الرئيسي فقط ونترك عمود المصغرات في مكانه بسكرولّه.
    const oldMain = oldRoot.querySelector(':scope > .pdf-main');
    if (oldMain) oldRoot.replaceChild(mainCol, oldMain); else oldRoot.append(mainCol);
    oldThumbs.querySelectorAll('.pdf-thumb').forEach((c, idx) => c.classList.toggle('on', idx === curIdx));
    if (Math.abs((window.scrollY || window.pageYOffset || 0) - previousWindowScroll) > 1) {
      window.scrollTo(0, previousWindowScroll);
    }
    return;
  }
  // نقرأ موضع السكرول من العمود القديم قبل ما نستبدله (بعد الاستبدال يصير العمود الجديد بصفر).
  previousThumbScroll = oldThumbs ? oldThumbs.scrollTop : previousThumbScroll;
  const focusedBefore = document.activeElement?.closest?.('.pdf-thumb');
  if (focusedBefore && oldThumbs?.contains(focusedBefore)) document.activeElement.blur();
  appEl.replaceChildren(root);

  // Build every thumbnail card immediately. This keeps the full page list
  // scrollable from the first frame, even while preview canvases render lazily.
  const thumbCards = [];
  for (let i = 0; i < total; i++) {
    const card = h('div', {
      class: 'pdf-thumb' + (i === curIdx ? ' on' : ''),
      draggable: true,
      'data-page-index': String(i),
      onclick: () => {
        if (curPdf.pageIdx === i) return;
        curPdf.pageIdx = i;
        renderPdfEditor();
      },
      ondragstart: (e) => {
        e.dataTransfer.setData('text/plain', String(i));
        e.dataTransfer.effectAllowed = 'move';
        card.classList.add('dragging');
      },
      ondragend: () => card.classList.remove('dragging'),
      ondragover: (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; },
      ondrop: (e) => {
        e.preventDefault();
        e.stopPropagation();
        const from = Number(e.dataTransfer.getData('text/plain'));
        if (Number.isInteger(from) && from >= 0 && from < total && from !== i) {
          movePdfPage(from, i);
        }
      }
    });
    const rot = curPdf.rotations[i] || 0;
    const preview = h('div', {
      class: 'pdf-thumb-preview',
      style: { width: '100%', minHeight: '150px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg, #f3f4f6)', borderRadius: '5px', overflow: 'hidden' }
    }, h('span', { class: 'pdf-thumb-loading' }, currentLanguage === 'en' ? 'Loading…' : 'جارٍ التحميل…'));
    const moveControls = h('span', { class: 'pdf-thumb-actions' },
      h('button', {
        type: 'button', class: 'pdf-thumb-move', disabled: i === 0,
        title: currentLanguage === 'en' ? 'Move page up' : 'نقل الصفحة للأعلى',
        onclick: (ev) => { ev.stopPropagation(); if (i > 0) movePdfPage(i, i - 1); }
      }, '↑'),
      h('button', {
        type: 'button', class: 'pdf-thumb-move', disabled: i === total - 1,
        title: currentLanguage === 'en' ? 'Move page down' : 'نقل الصفحة للأسفل',
        onclick: (ev) => { ev.stopPropagation(); if (i < total - 1) movePdfPage(i, i + 1); }
      }, '↓'),
      total > 1 ? h('button', {
        type: 'button',
        class: 'pdf-thumb-del',
        title: currentLanguage === 'en' ? 'Delete page' : 'حذف الصفحة',
        onclick: (ev) => {
          ev.stopPropagation();
          deletePdfPageAt(i);
        }
      }, '🗑') : null
    );
    const bar = h('div', { class: 'pdf-thumb-bar' },
      h('span', null, (currentLanguage === 'en' ? 'Page ' : 'صفحة ') + (i + 1)),
      moveControls
    );
    card.append(preview, bar);
    thumbsCol.append(card);
    thumbCards.push({ card, preview, pageIndex: i, rotation: rot });
  }

  // Read the latest position immediately before replacing the old sidebar;
  // the user may continue scrolling while the selected page is rasterizing.

  // إزالة التركيز من بطاقة الصفحة القديمة قبل استبدال DOM يمنع المتصفح
  // من تحريك المستند تلقائياً إلى الأعلى عند إزالة العنصر الذي تم النقر عليه.
  const focusedThumb = document.activeElement?.closest?.('.pdf-thumb');
  if (focusedThumb && oldThumbs?.contains(focusedThumb)) document.activeElement.blur();

  // تتبع موضع القائمة باستمرار، ثم استعده بعد تركيب القائمة وحساب تخطيطها.
  thumbsCol.addEventListener('scroll', () => {
    pdfThumbScrollTop = thumbsCol.scrollTop;
  }, { passive: true });

  thumbsCol.dataset.sig = thumbsSig;
  const restoreThumbScroll = () => {
    if (renderGeneration !== pdfRenderGeneration || !thumbsCol.isConnected) return;
    thumbsCol.scrollTop = previousThumbScroll;
    pdfThumbScrollTop = previousThumbScroll;
  };
  restoreThumbScroll();
  requestAnimationFrame(() => {
    restoreThumbScroll();
    requestAnimationFrame(restoreThumbScroll);
  });
  setTimeout(restoreThumbScroll, 80);

  // Preview rendering is deliberately independent from page navigation.
  // A navigation render must not cancel the thumbnail queue or leave holes.
  const thumbQueueToken = (curPdf._thumbQueueToken || 0) + 1;
  curPdf._thumbQueueToken = thumbQueueToken;
  (async () => {
    for (const item of thumbCards) {
      if (curPdf._thumbQueueToken !== thumbQueueToken || !thumbsCol.isConnected) return;
      try {
        const page = await pdfDoc.getPage(item.pageIndex + 1);
        const vp = page.getViewport({ scale: 0.24, rotation: item.rotation });
        const tCvs = h('canvas', { width: Math.max(1, Math.round(vp.width)), height: Math.max(1, Math.round(vp.height)) });
        tCvs.style.width = '100%';
        tCvs.style.height = 'auto';
        await page.render({ canvasContext: tCvs.getContext('2d'), viewport: vp }).promise;
        if (curPdf._thumbQueueToken !== thumbQueueToken || !thumbsCol.isConnected) return;
        item.preview.replaceChildren(tCvs);
      } catch (err) {
        console.warn('PDF thumbnail render failed for page', item.pageIndex + 1, err);
        if (item.preview.isConnected) item.preview.textContent = currentLanguage === 'en' ? 'Preview unavailable' : 'تعذر عرض المعاينة';
      }
    }
  })();

  // استعادة تمرير الصفحة الرئيسية بعد إعادة بناء الواجهة.
  if (Math.abs((window.scrollY || window.pageYOffset || 0) - previousWindowScroll) > 1) {
    window.scrollTo(0, previousWindowScroll);
  }
}

function hexToPdfRgb(hexStr) {
  const { rgb } = getPdfLib();
  const clean = (hexStr || '#1f1c17').replace('#', '');
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return rgb(
    isNaN(r) ? 0 : r,
    isNaN(g) ? 0 : g,
    isNaN(b) ? 0 : b
  );
}

function pdfViewportPointToPage(xPct, yPct, width, height, rotation) {
  const r = ((rotation % 360) + 360) % 360;
  if (r === 90) return { x: yPct * width, y: xPct * height };
  if (r === 180) return { x: (1 - xPct) * width, y: yPct * height };
  if (r === 270) return { x: (1 - yPct) * width, y: (1 - xPct) * height };
  return { x: xPct * width, y: (1 - yPct) * height };
}

async function buildMergedPdfBytesWithAnnotations() {
  const { PDFDocument, degrees } = getPdfLib();
  const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true });
  const pages = pdfDoc.getPages();
  const embeddedSigCache = new Map();
  let hasDigitalSignature = false;

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    const rot = curPdf.rotations[i] || 0;
    if (rot) {
      page.setRotation(degrees(rot));
    }

    const { width, height } = page.getSize();
    const viewW = rot % 180 ? height : width;
    const viewH = rot % 180 ? width : height;

    // 1. Render transparent high-DPI shaped glyphs for any non-WinAnsi (Arabic/Unicode) edits
    // (Original PDF stream glyphs were already surgically removed from curPdf.bytes — NO whiteout box!)
    const origEdits = Object.values(curPdf.textEdits[i] || {});
    for (const edit of origEdits) {
      const rawText = String(edit.text ?? '');
      if (!rawText.trim()) continue;
      const proj = projectPdfBoxFromRot(edit, edit.baseRot ?? rot, rot);
      const lines = rawText.split('\n');
      const szPt = Math.max(6, Math.min(96, edit.fontSizePt || 12));
      const dpiScale = 4;
      const fontPx = szPt * dpiScale;
      const lineH = Math.ceil(fontPx * 1.18);
      const fontSpec = `${edit.bold ? '700 ' : ''}${fontPx}px "Cairo","Segoe UI",Tahoma,Arial,sans-serif`;
      const isRtl = edit.dir === 'rtl' || /[\u0590-\u08FF]/.test(rawText);

      const tCvs = document.createElement('canvas');
      const tCtx = tCvs.getContext('2d');
      tCtx.font = fontSpec;
      const maxLinePx = Math.max(...lines.map(ln => Math.ceil(tCtx.measureText(ln || ' ').width)), 8);
      const padCanvasPx = Math.ceil(3 * dpiScale);
      tCvs.width = maxLinePx + padCanvasPx * 2;
      tCvs.height = Math.max(lineH * lines.length + Math.ceil(3 * dpiScale), Math.ceil(((proj.hPct / 100) * viewH) * dpiScale));

      tCtx.clearRect(0, 0, tCvs.width, tCvs.height);
      tCtx.font = fontSpec;
      tCtx.direction = isRtl ? 'rtl' : 'ltr';
      tCtx.textAlign = isRtl ? 'right' : 'left';
      tCtx.textBaseline = 'top';
      tCtx.fillStyle = edit.color || '#1f1c17';
      lines.forEach((ln, li) => {
        tCtx.fillText(ln, isRtl ? (tCvs.width - padCanvasPx) : padCanvasPx, Math.ceil(dpiScale) + li * lineH);
      });

      const png = await pdfDoc.embedPng(tCvs.toDataURL('image/png'));
      const imgW = tCvs.width / dpiScale;
      const imgH = tCvs.height / dpiScale;
      const startXView = isRtl
        ? ((proj.xPct + proj.wPct) / 100) * viewW - imgW + (padCanvasPx / dpiScale)
        : (proj.xPct / 100) * viewW - (padCanvasPx / dpiScale);
      const startYView = (proj.yPct / 100) * viewH - 1;
      const centerX = startXView + imgW / 2;
      const centerY = startYView + imgH / 2;
      const center = pdfViewportPointToPage(centerX / viewW, centerY / viewH, width, height, rot);
      const thetaDeg = (360 - (rot % 360)) % 360;
      const thetaRad = (thetaDeg * Math.PI) / 180;
      const drawX = center.x - (imgW / 2) * Math.cos(thetaRad) + (imgH / 2) * Math.sin(thetaRad);
      const drawY = center.y - (imgW / 2) * Math.sin(thetaRad) - (imgH / 2) * Math.cos(thetaRad);
      page.drawImage(png, {
        x: drawX,
        y: drawY,
        width: imgW,
        height: imgH,
        rotate: degrees(thetaDeg)
      });
    }

    const strokes = curPdf.drawings[i] || [];
    for (const st of strokes) {
      if (!st.pts || st.pts.length < 2) continue;
      if (st.mode === 'whiteout') {
        const a = st.pts[0];
        const b = st.pts[st.pts.length - 1];
        const c1 = pdfViewportPointToPage(a.x, a.y, width, height, rot);
        const c2 = pdfViewportPointToPage(b.x, b.y, width, height, rot);
        page.drawRectangle({
          x: Math.min(c1.x, c2.x),
          y: Math.min(c1.y, c2.y),
          width: Math.max(1, Math.abs(c2.x - c1.x)),
          height: Math.max(1, Math.abs(c2.y - c1.y)),
          color: hexToPdfRgb(st.color || '#ffffff')
        });
        continue;
      }
      const isHi = st.mode === 'highlight';
      const col = hexToPdfRgb(st.color || (isHi ? '#facc15' : '#e11d48'));
      const thick = isHi ? Math.max(8, (st.size || 18) * 0.85) : Math.max(1, (st.size || 3) * 0.75);
      for (let p = 1; p < st.pts.length; p++) {
        const a = st.pts[p - 1];
        const b = st.pts[p];
        page.drawLine({
          start: pdfViewportPointToPage(a.x, a.y, width, height, rot),
          end: pdfViewportPointToPage(b.x, b.y, width, height, rot),
          thickness: thick,
          color: col,
          opacity: isHi ? 0.35 : 1
        });
      }
    }

    const txtList = curPdf.texts[i] || [];
    for (const item of txtList) {
      const textValue = String(item.text || '').trim();
      if (!textValue) continue;
      const sz = Math.max(8, Math.min(72, item.size || 18));
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const scale = 2;
      ctx.font = `${sz * scale}px Arial, sans-serif`;
      const isRtl = /[\u0590-\u08FF]/.test(textValue);
      ctx.direction = isRtl ? 'rtl' : 'ltr';
      ctx.textAlign = isRtl ? 'right' : 'left';
      const measured = Math.ceil(ctx.measureText(textValue).width) + 12 * scale;
      canvas.width = Math.max(32, measured); canvas.height = Math.ceil(sz * scale * 1.8);
      ctx.font = `${sz * scale}px Arial, sans-serif`; ctx.direction = isRtl ? 'rtl' : 'ltr'; ctx.textAlign = isRtl ? 'right' : 'left';
      ctx.textBaseline = 'middle'; ctx.fillStyle = item.color || '#000000';
      ctx.fillText(textValue, isRtl ? canvas.width - 6 : 6, canvas.height / 2);
      const png = await pdfDoc.embedPng(canvas.toDataURL('image/png'));
      const imgW = canvas.width * 0.375, imgH = canvas.height * 0.375;
      const viewW = rot % 180 ? height : width, viewH = rot % 180 ? width : height;
      const centerX = (item.xPct / 100) * viewW + imgW / 2;
      const centerY = (item.yPct / 100) * viewH + imgH / 2;
      const center = pdfViewportPointToPage(centerX / viewW, centerY / viewH, width, height, rot);
      const imageRotation = degrees((360 - (rot % 360)) % 360);
      page.drawImage(png, { x: Math.max(-imgW, Math.min(width, center.x - imgW / 2)), y: Math.max(-imgH, Math.min(height, center.y - imgH / 2)), width: imgW, height: imgH, rotate: imageRotation });
    }

    const sigList = curPdf.signatures[i] || [];
    for (const sig of sigList) {
      const srcUrl = sig.stampDataUrl || sig.sigImgDataUrl;
      if (!srcUrl) continue;
      hasDigitalSignature = true;
      let png = embeddedSigCache.get(srcUrl);
      if (!png) {
        png = await pdfDoc.embedPng(srcUrl);
        embeddedSigCache.set(srcUrl, png);
      }
      const proj = projectPdfBoxFromRot(sig, sig.baseRot ?? rot, rot);
      const imgW = (proj.wPct / 100) * viewW;
      const imgH = (proj.hPct / 100) * viewH;
      const startXView = (proj.xPct / 100) * viewW;
      const startYView = (proj.yPct / 100) * viewH;
      const centerX = startXView + imgW / 2;
      const centerY = startYView + imgH / 2;
      const center = pdfViewportPointToPage(centerX / viewW, centerY / viewH, width, height, rot);
      const thetaDeg = (360 - (rot % 360)) % 360;
      const thetaRad = (thetaDeg * Math.PI) / 180;
      const drawX = center.x - (imgW / 2) * Math.cos(thetaRad) + (imgH / 2) * Math.sin(thetaRad);
      const drawY = center.y - (imgW / 2) * Math.sin(thetaRad) - (imgH / 2) * Math.cos(thetaRad);
      page.drawImage(png, {
        x: drawX,
        y: drawY,
        width: imgW,
        height: imgH,
        rotate: degrees(thetaDeg)
      });
    }
  }

  if (hasDigitalSignature) {
    try {
      pdfDoc.setProducer('brivatedocs Local Digital Signature Engine');
      pdfDoc.setModificationDate(new Date());
    } catch (_) {}
  }

  return pdfDoc.save();
}

async function exportCurrentPdfPagePng(pageIdx) {
  if (!curPdf.bytes) return;
  busy(true, currentLanguage === 'en' ? 'Exporting page as PNG…' : 'جارٍ تصدير الصفحة كصورة PNG…');
  try {
    const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
    const page = await pdfDoc.getPage(pageIdx + 1);
    const vp = page.getViewport({ scale: 1.8, rotation: curPdf.rotations[pageIdx] || 0 });
    const cvs = h('canvas', { width: Math.round(vp.width), height: Math.round(vp.height) });
    await page.render({ canvasContext: cvs.getContext('2d'), viewport: vp }).promise;

    const drawCvs = h('canvas', { width: cvs.width, height: cvs.height });
    redrawPdfDrawCanvas(drawCvs, pageIdx);
    const outCvs = h('canvas', { width: cvs.width, height: cvs.height });
    const oCtx = outCvs.getContext('2d');
    oCtx.drawImage(cvs, 0, 0);

    // Render transparent non-WinAnsi live edits onto exported PNG (WinAnsi edits are already in curPdf.bytes!)
    const curRotPng = curPdf.rotations[pageIdx] || 0;
    for (const edit of Object.values(curPdf.textEdits[pageIdx] || {})) {
      const rawText = String(edit.text ?? '');
      if (!rawText.trim()) continue;
      const proj = projectPdfBoxFromRot(edit, edit.baseRot ?? curRotPng, curRotPng);
      oCtx.save();
      const fontPx = Math.round((edit.fontSizePt || 12) * 1.8);
      const isRtl = edit.dir === 'rtl' || /[\u0590-\u08FF]/.test(rawText);
      oCtx.font = `${edit.bold ? '700 ' : ''}${fontPx}px "Cairo","Segoe UI",Tahoma,Arial,sans-serif`;
      oCtx.direction = isRtl ? 'rtl' : 'ltr';
      oCtx.textAlign = isRtl ? 'right' : 'left';
      oCtx.textBaseline = 'top';
      oCtx.fillStyle = edit.color || '#1f1c17';
      const tx = isRtl ? ((proj.xPct + proj.wPct) / 100) * outCvs.width : (proj.xPct / 100) * outCvs.width;
      const ty = (proj.yPct / 100) * outCvs.height;
      rawText.split('\n').forEach((ln, li) => oCtx.fillText(ln, tx, ty + li * Math.round(fontPx * 1.18)));
      oCtx.restore();
    }

    oCtx.drawImage(drawCvs, 0, 0);

    for (const item of (curPdf.texts[pageIdx] || [])) {
      if (!item.text) continue;
      oCtx.save();
      oCtx.fillStyle = item.color || '#1f1c17';
      oCtx.font = `${Math.round((item.size || 18) * (4 / 3) * 1.8)}px "Cairo","Segoe UI",sans-serif`;
      oCtx.fillText(item.text, (item.xPct / 100) * outCvs.width, ((item.yPct / 100) * outCvs.height) + 18);
      oCtx.restore();
    }

    for (const sig of (curPdf.signatures[pageIdx] || [])) {
      const srcUrl = sig.stampDataUrl || sig.sigImgDataUrl;
      if (!srcUrl) continue;
      const sImg = new Image();
      sImg.src = srcUrl;
      await sImg.decode();
      const proj = projectPdfBoxFromRot(sig, sig.baseRot ?? curRotPng, curRotPng);
      oCtx.drawImage(
        sImg,
        (proj.xPct / 100) * outCvs.width,
        (proj.yPct / 100) * outCvs.height,
        (proj.wPct / 100) * outCvs.width,
        (proj.hPct / 100) * outCvs.height
      );
    }

    const a = document.createElement('a');
    a.href = outCvs.toDataURL('image/png');
    a.download = `pdf-page-${pageIdx + 1}.png`;
    a.click();
    toast(currentLanguage === 'en' ? 'Page exported as PNG' : 'تم تصدير الصفحة كصورة PNG');
  } catch (err) {
    console.error('PDF page image export failed:', err);
    toast(currentLanguage === 'en' ? 'Failed to export image' : 'تعذر تصدير الصورة');
  } finally {
    busy(false);
  }
}

async function downloadEditedPdfFile() {
  if (!curPdf.bytes) return;
  busy(true, currentLanguage === 'en' ? 'Preparing PDF for download…' : 'جارٍ تجهيز ملف PDF للتحميل…');
  try {
    const outBytes = await buildMergedPdfBytesWithAnnotations();
    const blob = new Blob([outBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const base = (fileName || 'document.pdf').replace(/\.pdf$/i, '');
    a.href = url;
    a.download = `${base}-edited.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast(currentLanguage === 'en' ? 'File downloaded' : 'تم تحميل الملف');
    return true;
  } catch (err) {
    console.error('Edited PDF export failed:', err);
    toast(currentLanguage === 'en' ? 'An error occurred while exporting the file' : 'حدث خطأ أثناء تصدير الملف');
    return false;
  } finally {
    busy(false);
  }
}

/* ── إصلاح نهائي لـ Undo و Redo في محرر PDF (بما في ذلك القلم) ── */
let resetPdfUndoHistory = () => {};
(function fixPdfUndoRedoFinal() {
  const pdfStack = [];
  let pdfPtr = -1;
  let pdfLock = false;
  const PDF_HISTORY_MAX_STATES = 60;
  const PDF_HISTORY_MAX_BYTES = 256 * 1024 * 1024;

  // PDF byte arrays are replaced, never edited in place, by the page operations below.
  // Sharing these immutable references avoids an extra full-file copy per history entry.
  const trimPdfHistory = () => {
    const retainedBytes = () => {
      const seen = new Set();
      let total = 0;
      for (const state of pdfStack) {
        if (state.bytes && !seen.has(state.bytes)) {
          seen.add(state.bytes);
          total += state.bytes.byteLength || state.bytes.length || 0;
        }
      }
      return total;
    };
    while (pdfStack.length > 1 &&
      (pdfStack.length > PDF_HISTORY_MAX_STATES || retainedBytes() > PDF_HISTORY_MAX_BYTES)) {
      pdfStack.shift();
      pdfPtr = Math.max(0, pdfPtr - 1);
    }
  };

  const clonePdfState = () => {
    const drawCopy = {};
    for (const k of Object.keys(curPdf.drawings || {})) {
      drawCopy[k] = (curPdf.drawings[k] || []).map(stroke => ({
        mode: stroke.mode || 'pen',
        pts: (stroke.pts || []).map(pt => ({ x: pt.x, y: pt.y })),
        color: stroke.color,
        size: stroke.size
      }));
    }
    const textCopy = {};
    for (const k of Object.keys(curPdf.texts || {})) {
      textCopy[k] = (curPdf.texts[k] || []).map(t => ({ ...t }));
    }
    const sigCopy = {};
    for (const k of Object.keys(curPdf.signatures || {})) {
      sigCopy[k] = (curPdf.signatures[k] || []).map(s => ({ ...s }));
    }
    const editsCopy = {};
    for (const k of Object.keys(curPdf.textEdits || {})) {
      editsCopy[k] = {};
      for (const ek of Object.keys(curPdf.textEdits[k] || {})) {
        editsCopy[k][ek] = { ...curPdf.textEdits[k][ek] };
      }
    }
    return {
      bytes: curPdf.bytes || null,
      pageCount: curPdf.pageCount,
      pageIdx: curPdf.pageIdx,
      rotations: Array.isArray(curPdf.rotations) ? [...curPdf.rotations] : [],
      drawings: drawCopy,
      texts: textCopy,
      signatures: sigCopy,
      textEdits: editsCopy
    };
  };

  const isSameSnap = (a, b) => {
    if (!a || !b) return false;
    if (a.pageIdx !== b.pageIdx || a.pageCount !== b.pageCount) return false;
    if (JSON.stringify(a.rotations) !== JSON.stringify(b.rotations)) return false;
    if (JSON.stringify(a.drawings) !== JSON.stringify(b.drawings)) return false;
    if (JSON.stringify(a.texts) !== JSON.stringify(b.texts)) return false;
    if (JSON.stringify(a.signatures) !== JSON.stringify(b.signatures)) return false;
    if (JSON.stringify(a.textEdits) !== JSON.stringify(b.textEdits)) return false;
    // All PDF content-changing operations assign a newly-created Uint8Array.
    // Reference equality is therefore exact and avoids scanning large documents.
    return a.bytes === b.bytes;
  };

  const syncPdfButtons = () => {
    if (kind !== 'pdf') return;
    const u = $('#bUndo');
    const r = $('#bRedo');
    if (u) u.disabled = pdfPtr <= 0;
    if (r) r.disabled = pdfPtr < 0 || pdfPtr >= pdfStack.length - 1;
  };

  resetPdfUndoHistory = () => {
    pdfStack.length = 0;
    pdfPtr = -1;
    pdfLock = false;
    syncPdfButtons();
  };

  const pushPdfSnap = (force = false) => {
    if (pdfLock || kind !== 'pdf' || !curPdf.bytes) return;
    const snap = clonePdfState();
    if (!force && pdfPtr >= 0 && isSameSnap(pdfStack[pdfPtr], snap)) {
      syncPdfButtons();
      return;
    }
    pdfStack.splice(pdfPtr + 1);
    pdfStack.push(snap);
    pdfPtr = pdfStack.length - 1;
    trimPdfHistory();
    syncPdfButtons();
  };

  const applyPdfSnap = async (targetPtr) => {
    if (targetPtr < 0 || targetPtr >= pdfStack.length || pdfLock) return;
    pdfLock = true;
    pdfPtr = targetPtr;
    const snap = pdfStack[pdfPtr];
    curPdf.bytes = snap.bytes || null;
    curPdf.pageCount = snap.pageCount;
    curPdf.pageIdx = snap.pageIdx;
    curPdf.rotations = [...snap.rotations];
    curPdf.drawings = JSON.parse(JSON.stringify(snap.drawings || {}));
    curPdf.texts = JSON.parse(JSON.stringify(snap.texts || {}));
    curPdf.signatures = JSON.parse(JSON.stringify(snap.signatures || {}));
    curPdf.textEdits = JSON.parse(JSON.stringify(snap.textEdits || {}));
    curPdf.activeOrigEditRef = null;
    try {
      await renderPdfEditor();
    } finally {
      pdfLock = false;
      syncPdfButtons();
    }
  };

  const origInitPdf = initPdfFromFile;
  initPdfFromFile = async function (file) {
    pdfStack.length = 0;
    pdfPtr = -1;
    const res = await origInitPdf.call(this, file);
    pushPdfSnap(true);
    return res;
  };

  const origMergePdf = mergePdfFiles;
  mergePdfFiles = async function (filesList) {
    const res = await origMergePdf.call(this, filesList);
    pushPdfSnap(true);
    return res;
  };

  const origDelPdfPage = deletePdfPageAt;
  deletePdfPageAt = async function (targetIdx) {
    const res = await origDelPdfPage.call(this, targetIdx);
    pushPdfSnap(true);
    return res;
  };

  const origMovePdfPage = movePdfPage;
  movePdfPage = async function (fromIdx, toIdx) {
    const res = await origMovePdfPage.call(this, fromIdx, toIdx);
    pushPdfSnap(true);
    return res;
  };

  const origRotatePdf = rotateCurrentPdfPage;
  rotateCurrentPdfPage = function (degDelta) {
    const res = origRotatePdf.call(this, degDelta);
    pushPdfSnap(true);
    return res;
  };

  const origSurgicalEdit = applySurgicalPdfTextEdit;
  applySurgicalPdfTextEdit = async function (...args) {
    const res = await origSurgicalEdit.apply(this, args);
    if (res) pushPdfSnap(true);
    return res;
  };

  const origInsertSig = insertPdfSignatureOnCurrentPage;
  insertPdfSignatureOnCurrentPage = async function (...args) {
    const res = await origInsertSig.apply(this, args);
    if (res) pushPdfSnap(true);
    return res;
  };

  const origApplySigAll = applyPdfSignatureToAllPages;
  applyPdfSignatureToAllPages = async function (...args) {
    const res = await origApplySigAll.apply(this, args);
    if (res) pushPdfSnap(true);
    return res;
  };

  const origInsertBlank = insertBlankPdfPageAfterCurrent;
  insertBlankPdfPageAfterCurrent = async function (...args) {
    const res = await origInsertBlank.apply(this, args);
    if (res) pushPdfSnap(true);
    return res;
  };

  const origRenderPdf = renderPdfEditor;
  renderPdfEditor = async function () {
    const res = await origRenderPdf.call(this);
    if (!pdfLock && pdfPtr === -1 && curPdf.bytes) {
      pushPdfSnap(true);
    }

    const drawCvs = $('.pdf-draw-canvas');
    if (drawCvs && !drawCvs._pdfUndoReady) {
      drawCvs._pdfUndoReady = true;
      let strokeActive = false;
      drawCvs.addEventListener('pointerdown', () => {
        if (kind === 'pdf' && ['pen', 'highlight', 'whiteout'].includes(curPdf.tool)) strokeActive = true;
      });
      const finishStroke = () => {
        if (strokeActive) {
          strokeActive = false;
          setTimeout(() => pushPdfSnap(false), 15);
        }
      };
      drawCvs.addEventListener('pointerup', finishStroke);
      drawCvs.addEventListener('pointercancel', finishStroke);
      window.addEventListener('pointerup', finishStroke);
    }

    const textLayer = $('.pdf-text-layer');
    if (textLayer && !textLayer._pdfUndoReady) {
      textLayer._pdfUndoReady = true;
      textLayer.addEventListener('pointerup', () => {
        if (kind === 'pdf' && curPdf.tool === 'text') {
          setTimeout(() => pushPdfSnap(false), 25);
        }
      });
      textLayer.addEventListener('input', () => {
        if (kind === 'pdf') {
          clearTimeout(textLayer._pdfTxtTimer);
          textLayer._pdfTxtTimer = setTimeout(() => pushPdfSnap(false), 300);
        }
      });
    }

    const origLayer = $('.pdf-orig-layer');
    if (origLayer && !origLayer._pdfUndoReady) {
      origLayer._pdfUndoReady = true;
      origLayer.addEventListener('origeditinput', () => {
        if (kind === 'pdf') {
          clearTimeout(origLayer._pdfOrigTimer);
          origLayer._pdfOrigTimer = setTimeout(() => pushPdfSnap(false), 320);
        }
      });
      origLayer.addEventListener('origeditcommit', () => {
        if (kind === 'pdf') {
          clearTimeout(origLayer._pdfOrigTimer);
          setTimeout(() => pushPdfSnap(false), 20);
        }
      });
    }

    const sigLayer = $('.pdf-sig-layer');
    if (sigLayer && !sigLayer._pdfUndoReady) {
      sigLayer._pdfUndoReady = true;
      sigLayer.addEventListener('sigmutate', () => {
        if (kind === 'pdf') {
          setTimeout(() => pushPdfSnap(false), 20);
        }
      });
    }

    const toolbar = $('.pdf-toolbar');
    if (toolbar && !toolbar._pdfUndoReady) {
      toolbar._pdfUndoReady = true;
      toolbar.addEventListener('click', (e) => {
        if (kind === 'pdf' && e.target.closest('button')) {
          setTimeout(() => pushPdfSnap(false), 30);
        }
      });
      toolbar.addEventListener('change', (e) => {
        if (kind === 'pdf' && (e.target.type === 'color' || e.target.type === 'number')) {
          setTimeout(() => pushPdfSnap(false), 30);
        }
      });
    }

    syncPdfButtons();
    return res;
  };

  const origRenderEditor = renderEditor;
  renderEditor = function () {
    syncEditorTopBarForKind();
    if (kind === 'pdf') {
      const p = renderPdfEditor();
      if (p && typeof p.then === 'function') p.then(syncPdfButtons);
      else syncPdfButtons();
      return p;
    }
    return origRenderEditor();
  };

  const uBtn = $('#bUndo');
  const rBtn = $('#bRedo');

  if (uBtn && !uBtn._pdfFinalUndo) {
    uBtn._pdfFinalUndo = true;
    uBtn.addEventListener('click', (e) => {
      if (kind !== 'pdf') return;
      e.preventDefault();
      e.stopImmediatePropagation();
      if (pdfPtr > 0) applyPdfSnap(pdfPtr - 1);
    }, true);
  }

  if (rBtn && !rBtn._pdfFinalRedo) {
    rBtn._pdfFinalRedo = true;
    rBtn.addEventListener('click', (e) => {
      if (kind !== 'pdf') return;
      e.preventDefault();
      e.stopImmediatePropagation();
      if (pdfPtr < pdfStack.length - 1) applyPdfSnap(pdfPtr + 1);
    }, true);
  }

  window.pushPdfUndoSnapshot = () => pushPdfSnap(true);
})();

/* ── ربط محرر PDF بفتح الملفات والرسم والتحميل ── */
(function patchPdfIntegration() {
  const origHandleFile = handleFile;
  handleFile = async function (file) {
    if (!file) return;
    const isPdf = /\.pdf$/i.test(file.name || '') || file.type === 'application/pdf';
    if (!isPdf) return origHandleFile(file);

    showHomeError('');
    activeDraftId = null;
    pk = null;
    hist = []; histIdx = -1;
    curPdf.bytes = null; curPdf.pageCount = 0; curPdf.pageIdx = 0; curPdf.rotations = []; curPdf.drawings = {}; curPdf.texts = {}; curPdf.signatures = {}; curPdf.textEdits = {}; curPdf.activeOrigEditRef = null;
    busy(true, currentLanguage === 'en' ? 'Opening file…' : 'جارٍ فتح الملف…');
    try {
      await initPdfFromFile(file);
    } catch (err) {
      showView('home');
      showHomeError(err instanceof UserError ? err.message : ('حدث خطأ أثناء قراءة الملف: ' + (err && err.message || err)));
    } finally {
      busy(false);
    }
  };

  const fileInEl = $('#fileIn');
  if (fileInEl && !fileInEl._pdfIntercepted) {
    fileInEl._pdfIntercepted = true;
    fileInEl.addEventListener('change', (e) => {
      const f = e.target?.files?.[0];
      if (f && (/\.pdf$/i.test(f.name || '') || f.type === 'application/pdf')) {
        e.stopImmediatePropagation();
        handleFile(f);
        fileInEl.value = '';
      }
    }, true);
  }

  const origRenderEditor = renderEditor;
  renderEditor = function () {
    syncEditorTopBarForKind();
    if (kind === 'pdf') { renderPdfEditor(); return; }
    return origRenderEditor();
  };

  const origDownload = downloadEditedFile;
  downloadEditedFile = async function () {
    if (kind === 'pdf') return downloadEditedPdfFile();
    return origDownload();
  };
})();
/* ── ربط محرر PDF بقاعدة بيانات IndexedDB (DraftDB) ── */
(function connectPdfToDraftDb() {
  if (typeof DraftDB === 'undefined' || typeof curPdf === 'undefined' || typeof renderPdfEditor !== 'function') return;

  const packPdfForDb = () => {
    const drawCopy = {};
    for (const k of Object.keys(curPdf.drawings || {})) {
      drawCopy[k] = (curPdf.drawings[k] || []).map(st => ({
        mode: st.mode || 'pen',
        pts: (st.pts || []).map(pt => ({ x: pt.x, y: pt.y })),
        color: st.color,
        size: st.size
      }));
    }
    const textCopy = {};
    for (const k of Object.keys(curPdf.texts || {})) {
      textCopy[k] = (curPdf.texts[k] || []).map(t => ({ ...t }));
    }
    const sigCopy = {};
    for (const k of Object.keys(curPdf.signatures || {})) {
      sigCopy[k] = (curPdf.signatures[k] || []).map(s => ({ ...s }));
    }
    const editsCopy = {};
    for (const k of Object.keys(curPdf.textEdits || {})) {
      editsCopy[k] = {};
      for (const ek of Object.keys(curPdf.textEdits[k] || {})) {
        editsCopy[k][ek] = { ...curPdf.textEdits[k][ek] };
      }
    }
    return {
      pdfBytes: curPdf.bytes ? curPdf.bytes.slice() : new Uint8Array(),
      pageCount: curPdf.pageCount || 0,
      pageIdx: curPdf.pageIdx || 0,
      rotations: Array.isArray(curPdf.rotations) ? [...curPdf.rotations] : [],
      drawings: drawCopy,
      texts: textCopy,
      signatures: sigCopy,
      textEdits: editsCopy
    };
  };

  const savePdfDraftNow = () => {
    if (kind !== 'pdf' || !curPdf.bytes) return;
    DraftDB.save({
      files: [],
      pdfData: packPdfForDb(),
      fileName: fileName || 'document.pdf',
      kind: 'pdf',
      mainPart: '',
      updatedAt: Date.now()
    });
  };
  window.savePdfDraftNow = savePdfDraftNow;

  const withDraftSave = (fn) => async function (...args) {
    const res = await fn.apply(this, args);
    savePdfDraftNow();
    return res;
  };
  initPdfFromFile = withDraftSave(initPdfFromFile);
  mergePdfFiles = withDraftSave(mergePdfFiles);
  deletePdfPageAt = withDraftSave(deletePdfPageAt);
  movePdfPage = withDraftSave(movePdfPage);
  rotateCurrentPdfPage = withDraftSave(rotateCurrentPdfPage);
  applySurgicalPdfTextEdit = withDraftSave(applySurgicalPdfTextEdit);
  insertPdfSignatureOnCurrentPage = withDraftSave(insertPdfSignatureOnCurrentPage);
  applyPdfSignatureToAllPages = withDraftSave(applyPdfSignatureToAllPages);
  insertBlankPdfPageAfterCurrent = withDraftSave(insertBlankPdfPageAfterCurrent);

  const origRenderPdf = renderPdfEditor;
  renderPdfEditor = async function (...args) {
    const res = await origRenderPdf.apply(this, args);
    const stageWrap = $('.pdf-stage-wrap');
    if (stageWrap && !stageWrap._dbBound) {
      stageWrap._dbBound = true;
      stageWrap.addEventListener('pointerup', () => {
        if (kind === 'pdf' && ['pen', 'highlight', 'whiteout', 'text', 'edit-text'].includes(curPdf.tool)) {
          setTimeout(savePdfDraftNow, 40);
        }
      });
      stageWrap.addEventListener('origeditcommit', () => {
        if (kind === 'pdf') setTimeout(savePdfDraftNow, 40);
      });
      stageWrap.addEventListener('sigmutate', () => {
        if (kind === 'pdf') setTimeout(savePdfDraftNow, 40);
      });
      stageWrap.addEventListener('input', () => {
        if (kind === 'pdf') {
          clearTimeout(stageWrap._pdfDbTimer);
          stageWrap._pdfDbTimer = setTimeout(savePdfDraftNow, 400);
        }
      });
    }
    const toolbar = $('.pdf-toolbar');
    if (toolbar && !toolbar._dbBound) {
      toolbar._dbBound = true;
      toolbar.addEventListener('click', (e) => {
        if (kind === 'pdf' && e.target.closest('button')) {
          setTimeout(savePdfDraftNow, 60);
        }
      });
    }
    return res;
  };

  const origRestoreDraft = restoreSavedDraft;
  restoreSavedDraft = async function (targetDraft) {
    const draft = targetDraft || (await DraftDB.loadAll())[0];
    if (draft && draft.kind === 'pdf' && draft.pdfData) {
      busy(true, currentLanguage === 'en' ? 'Opening file…' : 'جارٍ فتح الملف…');
      try {
        activeDraftId = draft.id || ('draft_' + Date.now());
        kind = 'pdf';
        fileName = draft.fileName || 'document.pdf';
        if ($('#fname')) $('#fname').textContent = fileName;
        const pd = draft.pdfData;
        curPdf.bytes = pd.pdfBytes ? (pd.pdfBytes instanceof Uint8Array ? pd.pdfBytes.slice() : new Uint8Array(pd.pdfBytes)) : null;
        curPdf.pageCount = pd.pageCount || 0;
        curPdf.pageIdx = pd.pageIdx || 0;
        curPdf.zoom = window.innerWidth <= 680 ? 'fit' : 1.15;
        curPdf.rotations = Array.isArray(pd.rotations) ? [...pd.rotations] : [];
        curPdf.drawings = JSON.parse(JSON.stringify(pd.drawings || {}));
        curPdf.texts = JSON.parse(JSON.stringify(pd.texts || {}));
        curPdf.signatures = JSON.parse(JSON.stringify(pd.signatures || {}));
        curPdf.textEdits = JSON.parse(JSON.stringify(pd.textEdits || {}));
        curPdf.activeOrigEditRef = null;
        resetHistory();
        resetPdfUndoHistory();
        showView('editor');
        await renderPdfEditor();
        applyLanguage(currentLanguage);
        toast('تمت استعادة المسودة بنجاح');
      } catch (err) {
        if (typeof showHomeError === 'function') {
          showHomeError(err instanceof UserError ? err.message : ('حدث خطأ أثناء قراءة الملف: ' + (err && err.message || err)));
        }
      } finally {
        busy(false);
      }
      return;
    }
    return origRestoreDraft(targetDraft);
  };

  const origDownloadPdf = downloadEditedPdfFile;
  downloadEditedPdfFile = async function () {
    const succeeded = await origDownloadPdf.apply(this, arguments);
    if (succeeded === true && activeDraftId) {
      await DraftDB.removeById(activeDraftId);
      activeDraftId = null;
    }
    return succeeded;
  };
})();
/* ── تحسين ترتيب صفحات PDF (سحب وإفلات أدق + أزرار نقل فورية) ── */
(function enhancePdfPageOrdering() {
  if (typeof renderPdfEditor !== 'function' || typeof movePdfPage !== 'function') return;

  const origRender = renderPdfEditor;
  renderPdfEditor = async function (...args) {
    const res = await origRender.apply(this, args);
    if (kind !== 'pdf' || !curPdf.bytes) return res;

    const thumbs = $$('.pdf-thumb');
    const total = curPdf.pageCount || thumbs.length;
    const curIdx = curPdf.pageIdx || 0;

    // Thumbnail cards already receive drag/drop handlers from renderPdfEditor.
    // Rebinding here on every render created redundant listeners on reused cards.

    const toolbar = $('.pdf-toolbar');
    if (toolbar && !toolbar._pdfMoveBtnsAdded) {
      toolbar._pdfMoveBtnsAdded = true;
      const bUp = btn(
        currentLanguage === 'en' ? '⬆️ Move Up' : '⬆️ للأعلى',
        () => { if ((curPdf.pageIdx || 0) > 0) movePdfPage(curPdf.pageIdx, curPdf.pageIdx - 1); },
        '',
        { disabled: curIdx <= 0 }
      );
      const bDown = btn(
        currentLanguage === 'en' ? '⬇️ Move Down' : '⬇️ للأسفل',
        () => { if ((curPdf.pageIdx || 0) < total - 1) movePdfPage(curPdf.pageIdx, curPdf.pageIdx + 1); },
        '',
        { disabled: curIdx >= total - 1 }
      );
      const delBtn = $$('button', toolbar).find(b => /Delete Page|حذف الصفحة/i.test(b.textContent || ''));
      if (delBtn && delBtn.parentNode === toolbar) {
        toolbar.insertBefore(bUp, delBtn.nextSibling);
        toolbar.insertBefore(bDown, bUp.nextSibling);
      } else {
        toolbar.append(bUp, bDown);
      }
    }

    return res;
  };
})();
(function fixPdfMergeBusyLang() {
  if (typeof mergePdfFiles !== 'function') return;
  const origMerge = mergePdfFiles;
  mergePdfFiles = async function (filesList) {
    const origBusy = busy;
    busy = function (on, msg) {
      if (on && msg === 'جارٍ دمج ملفات PDF…') {
        msg = currentLanguage === 'en' ? 'Merging PDF files…' : 'جارٍ دمج ملفات PDF…';
      }
      return origBusy.call(this, on, msg);
    };
    try {
      return await origMerge.apply(this, arguments);
    } finally {
      busy = origBusy;
    }
  };
})();
