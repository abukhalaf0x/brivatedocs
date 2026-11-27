'use strict';
/* ══════════════════════════════════════════════════════════════
   brivatedocs — Phase 1 Power Features (100% Client-Side)
   1. File System Access API & PWA File Handlers (Direct Save)
   2. Metadata Scrubber (PDF, DOCX, PPTX)
   3. PDF Watermark & Page Numbers
   4. True PDF Redaction (Surgical Stream Strip + Optional Flatten)
   5. Command Palette (Ctrl+K)
   ══════════════════════════════════════════════════════════════ */
(() => {
  const isAr = () => currentLanguage === 'ar';
  const say = (ar, en) => (isAr() ? ar : en);
  const $id = (id) => document.getElementById(id);

  function openPhaseModal(title, opts = {}) {
    const body = h('div', { class: 'pt-body' });
    let closed = false;
    const api = {
      body,
      render(...nodes) { body.replaceChildren(...nodes.flat().filter(Boolean)); },
      close() {
        if (closed) return;
        closed = true;
        backdrop.remove();
        document.removeEventListener('keydown', onKey);
        if (opts.onClose) opts.onClose();
      }
    };
    const onKey = (e) => { if (e.key === 'Escape') api.close(); };
    const box = h('div', { class: 'md-modal pt-modal p1-modal' + (opts.wide ? ' p1-modal-wide' : '') },
      h('div', { class: 'row', style: { justifyContent: 'space-between', alignItems: 'center' } },
        h('strong', null, title),
        h('button', { type: 'button', class: 'btn', style: { padding: '2px 8px' }, onclick: () => api.close() }, '✕')
      ),
      body
    );
    const backdrop = h('div', { class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) api.close(); } }, box);
    document.body.append(backdrop);
    document.addEventListener('keydown', onKey);
    return api;
  }

  /* ═══════════════════════════════════════════════════════════
     1. File System Access API & PWA File Handlers (Direct Save)
     ═══════════════════════════════════════════════════════════ */
  let activeFileHandle = null;
  let suppressHandleReset = false;

  const FILE_PICKER_TYPES = [
    {
      description: 'Documents (DOCX, PPTX, PDF)',
      accept: {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx', '.docm'],
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx', '.pptm'],
        'application/pdf': ['.pdf']
      }
    }
  ];

  function updateDirectSaveUi() {
    const saveBtn = $id('bDirectSave');
    if (saveBtn) {
      saveBtn.hidden = !activeFileHandle;
      if (activeFileHandle) {
        saveBtn.title = say('حفظ التعديلات مباشرة على الملف الأصلي (Ctrl+S)', 'Save changes directly to original file (Ctrl+S)');
      }
    }
    const statusEl = document.querySelector('#editor .editor-status');
    if (statusEl) {
      statusEl.classList.toggle('direct-linked', !!activeFileHandle);
      statusEl.title = activeFileHandle
        ? say('مرتبط بالملف المحلي — Ctrl+S يحفظ مباشرة', 'Linked to local file — Ctrl+S saves directly')
        : '';
    }
  }

  function clearActiveHandle() {
    if (suppressHandleReset) return;
    activeFileHandle = null;
    updateDirectSaveUi();
  }

  async function openWithFileHandle(handle) {
    if (!handle || typeof handle.getFile !== 'function') return;
    const file = await handle.getFile();
    if (!file) return;
    suppressHandleReset = true;
    try {
      await handleFile(file);
      activeFileHandle = handle;
      updateDirectSaveUi();
    } finally {
      suppressHandleReset = false;
    }
  }

  async function openLocalFileDirect() {
    if (typeof window.showOpenFilePicker !== 'function') {
      $id('fileIn')?.click();
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: FILE_PICKER_TYPES
      });
      if (handle) await openWithFileHandle(handle);
    } catch (err) {
      if (err && err.name === 'AbortError') return;
      // Fallback to standard file input if picker is restricted
      $id('fileIn')?.click();
    }
  }

  async function buildActiveFileBlob() {
    if (kind === 'pdf') {
      if (!curPdf?.bytes || typeof buildMergedPdfBytesWithAnnotations !== 'function') return null;
      const outBytes = await buildMergedPdfBytesWithAnnotations();
      return new Blob([outBytes], { type: 'application/pdf' });
    }
    if (!pk) return null;
    if (kind === 'docx') {
      const paper = $('#app .paper');
      if (paper && typeof syncPaperToDocxXml === 'function') syncPaperToDocxXml(paper);
      if (typeof normalizeWordXmlPackage === 'function') normalizeWordXmlPackage(pk);
    }
    const mime = kind === 'pptx'
      ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    return pk.toBlob(mime);
  }

  async function writeBlobToHandle(handle, blob) {
    if (handle.queryPermission && handle.requestPermission) {
      let perm = await handle.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') perm = await handle.requestPermission({ mode: 'readwrite' });
      if (perm !== 'granted') throw new UserError(say('لم يتم منح صلاحية الكتابة على الملف.', 'Write permission to the file was not granted.'));
    }
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  async function saveDirectToDisk() {
    if (!activeFileHandle) return downloadEditedFile();
    busy(true, say('جارٍ الحفظ المباشر على الملف…', 'Saving directly to file…'));
    try {
      const blob = await buildActiveFileBlob();
      if (!blob) return false;
      await writeBlobToHandle(activeFileHandle, blob);
      if (activeDraftId && typeof DraftDB !== 'undefined') {
        await DraftDB.removeById(activeDraftId);
        activeDraftId = null;
      }
      toast(say('تم الحفظ مباشرة على الملف الأصلي ✓', 'Saved directly to original file ✓'));
      return true;
    } catch (err) {
      if (err && err.name === 'AbortError') return false;
      console.error('Direct save failed:', err);
      toast(err instanceof UserError ? err.message : say('تعذر الحفظ المباشر — سيتم التنزيل كنسخة', 'Direct save failed — downloading copy instead'));
      return downloadEditedFile();
    } finally {
      busy(false);
    }
  }

  async function saveAsLocalFile() {
    if (typeof window.showSaveFilePicker !== 'function') return downloadEditedFile();
    const ext = kind === 'pdf' ? '.pdf' : (kind === 'pptx' ? '.pptx' : '.docx');
    const mime = kind === 'pdf'
      ? 'application/pdf'
      : (kind === 'pptx'
          ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    const suggested = (fileName || ('document' + ext)).replace(/\.(docx|pptx|docm|pptm|pdf)$/i, '') + ext;
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: suggested,
        types: [{ description: ext.toUpperCase().slice(1), accept: { [mime]: [ext] } }]
      });
      busy(true, say('جارٍ حفظ الملف…', 'Saving file…'));
      const blob = await buildActiveFileBlob();
      if (!blob) return false;
      await writeBlobToHandle(handle, blob);
      activeFileHandle = handle;
      if (handle.name) {
        fileName = handle.name;
        if ($id('fname')) $id('fname').textContent = fileName;
      }
      updateDirectSaveUi();
      toast(say('تم حفظ الملف وربطه للحفظ المباشر ✓', 'File saved & linked for direct save ✓'));
      return true;
    } catch (err) {
      if (err && err.name === 'AbortError') return false;
      console.error('Save As failed:', err);
      toast(say('تعذر حفظ الملف', 'Failed to save file'));
      return false;
    } finally {
      busy(false);
    }
  }

  // Intercept Home dropzone click to use File System Access API when supported
  const dropEl = $id('drop');
  if (dropEl && typeof window.showOpenFilePicker === 'function') {
    dropEl.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopImmediatePropagation();
      openLocalFileDirect();
    }, true);
  }

  // Register PWA File Handler (OS "Open With brivatedocs")
  if ('launchQueue' in window && typeof window.launchQueue.setConsumer === 'function') {
    window.launchQueue.setConsumer(async (launchParams) => {
      if (!launchParams?.files?.length) return;
      try {
        await openWithFileHandle(launchParams.files[0]);
      } catch (err) {
        console.error('PWA launchQueue failed:', err);
      }
    });
  }

  // Reset file handle when switching files or creating new docs
  if (typeof handleFile === 'function') {
    const origHandleFile = handleFile;
    handleFile = async function (...args) {
      clearActiveHandle();
      pendingRedactions = {};
      return origHandleFile.apply(this, args);
    };
  }
  if (typeof startNewDeck === 'function') {
    const origStartDeck = startNewDeck;
    startNewDeck = function (...args) {
      clearActiveHandle();
      return origStartDeck.apply(this, args);
    };
  }
  if (typeof startNewDocx === 'function') {
    const origStartDocx = startNewDocx;
    startNewDocx = function (...args) {
      clearActiveHandle();
      return origStartDocx.apply(this, args);
    };
  }
  if (typeof restoreSavedDraft === 'function') {
    const origRestore = restoreSavedDraft;
    restoreSavedDraft = async function (...args) {
      clearActiveHandle();
      pendingRedactions = {};
      return origRestore.apply(this, args);
    };
  }

  // Ctrl+S uses direct save when file handle is linked
  document.addEventListener('keydown', (e) => {
    if (!$id('editor')?.hidden && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      if (activeFileHandle) {
        e.preventDefault();
        e.stopImmediatePropagation();
        saveDirectToDisk();
      }
    }
  }, true);

  /* ═══════════════════════════════════════════════════════════
     2. Metadata Scrubber (PDF, DOCX, PPTX)
     ═══════════════════════════════════════════════════════════ */
  async function inspectActiveMetadata() {
    const items = [];
    if (kind === 'pdf' && curPdf?.bytes) {
      const { PDFDocument, PDFName } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const check = (ar, en, val) => {
        if (val == null) return;
        const s = val instanceof Date ? val.toISOString() : String(val).trim();
        if (s) items.push({ label: say(ar, en), value: s });
      };
      check('العنوان', 'Title', pdfDoc.getTitle());
      check('المؤلف', 'Author', pdfDoc.getAuthor());
      check('الموضوع', 'Subject', pdfDoc.getSubject());
      check('الكلمات المفتاحية', 'Keywords', pdfDoc.getKeywords());
      check('البرنامج المنشئ', 'Creator', pdfDoc.getCreator());
      check('المنتج (Producer)', 'Producer', pdfDoc.getProducer());
      try { check('تاريخ الإنشاء', 'Creation Date', pdfDoc.getCreationDate()); } catch (_) {}
      try { check('تاريخ التعديل', 'Modification Date', pdfDoc.getModificationDate()); } catch (_) {}
      if (pdfDoc.catalog && pdfDoc.catalog.has(PDFName.of('Metadata'))) {
        items.push({ label: say('بيانات XMP المخفية', 'Embedded XMP Metadata'), value: say('موجودة (XML Stream)', 'Present (XML Stream)') });
      }
      return items;
    }

    if (pk && (kind === 'docx' || kind === 'pptx')) {
      const readTag = (xmlDoc, localName) => {
        if (!xmlDoc?.documentElement) return '';
        const els = xmlDoc.getElementsByTagName('*');
        for (const el of els) {
          if (el.localName === localName && el.textContent?.trim()) return el.textContent.trim();
        }
        return '';
      };
      if (pk.has('docProps/core.xml')) {
        const core = pk.doc('docProps/core.xml');
        const coreMap = [
          ['creator', 'المؤلف', 'Author'],
          ['lastModifiedBy', 'آخر تعديل بواسطة', 'Last Modified By'],
          ['title', 'العنوان', 'Title'],
          ['subject', 'الموضوع', 'Subject'],
          ['description', 'الوصف', 'Description'],
          ['keywords', 'الكلمات المفتاحية', 'Keywords'],
          ['category', 'التصنيف', 'Category'],
          ['revision', 'رقم المراجعة', 'Revision'],
          ['created', 'تاريخ الإنشاء', 'Created'],
          ['modified', 'تاريخ التعديل', 'Modified']
        ];
        for (const [tag, ar, en] of coreMap) {
          const v = readTag(core, tag);
          if (v) items.push({ label: say(ar, en), value: v });
        }
      }
      if (pk.has('docProps/app.xml')) {
        const app = pk.doc('docProps/app.xml');
        const appMap = [
          ['Company', 'الشركة / الجهة', 'Company'],
          ['Manager', 'المدير', 'Manager'],
          ['Application', 'التطبيق', 'Application'],
          ['AppVersion', 'إصدار التطبيق', 'App Version'],
          ['Template', 'القالب', 'Template'],
          ['TotalTime', 'مدة التحرير (دقائق)', 'Total Edit Time (min)']
        ];
        for (const [tag, ar, en] of appMap) {
          const v = readTag(app, tag);
          if (v && v !== '0') items.push({ label: say(ar, en), value: v });
        }
      }
      if (pk.has('docProps/custom.xml')) {
        const custom = pk.doc('docProps/custom.xml');
        const props = custom?.documentElement ? Array.from(custom.documentElement.children) : [];
        if (props.length) {
          items.push({
            label: say('خصائص مخصصة', 'Custom Properties'),
            value: say(`${props.length} خاصية`, `${props.length} propert${props.length === 1 ? 'y' : 'ies'}`)
          });
        }
      }
    }
    return items;
  }

  async function scrubActiveMetadata() {
    if (kind === 'pdf' && curPdf?.bytes) {
      const { PDFDocument, PDFName } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      pdfDoc.setTitle('');
      pdfDoc.setAuthor('');
      pdfDoc.setSubject('');
      pdfDoc.setKeywords([]);
      pdfDoc.setCreator('');
      pdfDoc.setProducer('');
      if (pdfDoc.catalog && pdfDoc.catalog.has(PDFName.of('Metadata'))) {
        pdfDoc.catalog.delete(PDFName.of('Metadata'));
      }
      if (pdfDoc.context?.trailerInfo?.Info) {
        const infoDict = pdfDoc.context.lookup(pdfDoc.context.trailerInfo.Info);
        if (infoDict && typeof infoDict.keys === 'function') {
          for (const k of Array.from(infoDict.keys())) infoDict.delete(k);
        }
        delete pdfDoc.context.trailerInfo.Info;
      }
      const cleaned = await pdfDoc.save({ useObjectStreams: false, updateMetadata: false });
      curPdf.bytes = new Uint8Array(cleaned);
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      if (typeof renderPdfEditor === 'function') await renderPdfEditor();
      return true;
    }

    if (pk && (kind === 'docx' || kind === 'pptx')) {
      if (kind === 'docx') {
        const paper = $('#app .paper');
        if (paper && typeof syncPaperToDocxXml === 'function') syncPaperToDocxXml(paper);
      }
      const clearTags = (xmlDoc, tagNames) => {
        if (!xmlDoc?.documentElement) return;
        const set = new Set(tagNames);
        const els = Array.from(xmlDoc.getElementsByTagName('*'));
        for (const el of els) {
          if (set.has(el.localName)) el.textContent = '';
        }
      };
      if (pk.has('docProps/core.xml')) {
        clearTags(pk.doc('docProps/core.xml'), [
          'creator', 'lastModifiedBy', 'title', 'subject', 'description',
          'keywords', 'category', 'revision', 'created', 'modified', 'lastPrinted'
        ]);
      }
      if (pk.has('docProps/app.xml')) {
        clearTags(pk.doc('docProps/app.xml'), [
          'Company', 'Manager', 'Application', 'AppVersion', 'Template', 'TotalTime', 'HyperlinkBase'
        ]);
      }
      if (pk.has('docProps/custom.xml')) {
        const custom = pk.doc('docProps/custom.xml');
        if (custom?.documentElement) {
          while (custom.documentElement.firstChild) custom.documentElement.removeChild(custom.documentElement.firstChild);
        }
      }
      pk.commit();
      if (typeof pushHistory === 'function') pushHistory('scrubMetadata');
      return true;
    }
    return false;
  }

  async function openMetadataScrubberModal() {
    if ($id('editor')?.hidden || (!pk && !curPdf?.bytes)) {
      toast(say('افتح ملفاً أولاً لتنظيف بياناته المخفية', 'Open a file first to clean its metadata'));
      return;
    }
    const modal = openPhaseModal('🛡️ ' + say('تنظيف البيانات المخفية (Metadata)', 'Clean Hidden Metadata'));
    modal.render(h('div', { class: 'pt-note' }, say('جارٍ فحص البيانات الوصفية…', 'Scanning document metadata…')));
    try {
      const items = await inspectActiveMetadata();
      const list = items.length
        ? h('div', { class: 'p1-meta-list' },
            ...items.map((it) =>
              h('div', { class: 'p1-meta-row' },
                h('span', { class: 'p1-meta-k' }, it.label),
                h('span', { class: 'p1-meta-v', title: it.value }, it.value)
              )
            )
          )
        : h('div', { class: 'p1-clean-badge' }, '✓ ' + say('الملف نظيف! لا توجد بيانات وصفية مخفية.', 'File is clean! No hidden metadata detected.'));

      const scrubBtn = h('button', {
        type: 'button',
        class: 'btn pri',
        disabled: !items.length,
        onclick: async () => {
          scrubBtn.disabled = true;
          scrubBtn.textContent = say('جارٍ التنظيف…', 'Cleaning…');
          try {
            await scrubActiveMetadata();
            modal.close();
            toast(say('تم مسح جميع البيانات المخفية بنجاح ✓', 'All hidden metadata scrubbed successfully ✓'));
          } catch (err) {
            console.error('Metadata scrub failed:', err);
            toast(say('تعذر تنظيف البيانات المخفية', 'Failed to clean metadata'));
            scrubBtn.disabled = false;
            scrubBtn.textContent = say('تنظيف البيانات الآن', 'Clean All Metadata');
          }
        }
      }, say('تنظيف البيانات الآن', 'Clean All Metadata'));

      modal.render(
        h('div', { class: 'pt-note' },
          say(
            'يزيل اسم المؤلف، الجهة، البرنامج، التواريخ، وبيانات XMP المخفية داخل الملف لحماية خصوصيتك قبل المشاركة.',
            'Removes hidden author, company, software, timestamps, and XMP streams from your file before sharing.'
          )
        ),
        list,
        h('div', { class: 'pt-actions' },
          h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')),
          scrubBtn
        )
      );
    } catch (err) {
      console.error('Metadata inspect failed:', err);
      modal.render(
        h('div', { class: 'pt-err' }, say('تعذر قراءة البيانات الوصفية للملف.', 'Could not read document metadata.')),
        h('div', { class: 'pt-actions' }, h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')))
      );
    }
  }

  /* ═══════════════════════════════════════════════════════════
     3. PDF Watermark & Page Numbers (Unicode / Arabic Ready)
     ═══════════════════════════════════════════════════════════ */
  function parsePageRanges(spec, total) {
    const clean = String(spec || '').trim();
    if (!clean) return Array.from({ length: total }, (_, i) => i);
    const out = new Set();
    for (const part of clean.split(',')) {
      const seg = part.trim();
      if (!seg) continue;
      const m = /^(\d+)\s*-\s*(\d+)$/.exec(seg);
      if (m) {
        const a = Math.max(1, Math.min(total, parseInt(m[1], 10)));
        const b = Math.max(1, Math.min(total, parseInt(m[2], 10)));
        for (let p = Math.min(a, b); p <= Math.max(a, b); p++) out.add(p - 1);
      } else {
        const n = parseInt(seg, 10);
        if (!isNaN(n) && n >= 1 && n <= total) out.add(n - 1);
      }
    }
    return Array.from(out).sort((a, b) => a - b);
  }

  async function renderStampTextToPng(text, fontSizePx, colorHex, opacity = 1, bold = true) {
    const scale = 3;
    const fontStr = `${bold ? '700 ' : '500 '}${fontSizePx * scale}px "Segoe UI", Tahoma, "Noto Sans Arabic", Arial, sans-serif`;
    const probe = document.createElement('canvas').getContext('2d');
    probe.font = fontStr;
    const metrics = probe.measureText(text);
    const padX = Math.ceil(fontSizePx * scale * 0.45);
    const padY = Math.ceil(fontSizePx * scale * 0.35);
    const w = Math.max(10, Math.ceil(metrics.width + padX * 2));
    const hPx = Math.max(10, Math.ceil(fontSizePx * scale * 1.35 + padY * 2));

    const cvs = document.createElement('canvas');
    cvs.width = w;
    cvs.height = hPx;
    const ctx = cvs.getContext('2d');
    ctx.globalAlpha = Math.max(0.05, Math.min(1, opacity));
    ctx.font = fontStr;
    ctx.fillStyle = colorHex || '#1f1c17';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    if (/[\u0590-\u08FF]/.test(text)) ctx.direction = 'rtl';
    ctx.fillText(text, w / 2, hPx / 2);

    const blob = await new Promise((res) => cvs.toBlob(res, 'image/png'));
    const bytes = new Uint8Array(await blob.arrayBuffer());
    return { bytes, widthPt: w / scale, heightPt: hPx / scale };
  }

  function drawCenteredRotatedStamp(page, pngImg, cxView, cyView, wPt, hPt, stampDeg, viewW, viewH, pageW, pageH, curRot) {
    const { degrees } = window.PDFLib;
    const totalDeg = (stampDeg + (360 - (curRot % 360))) % 360;
    const rad = (totalDeg * Math.PI) / 180;
    const centerPage = pdfViewportPointToPage(cxView / viewW, cyView / viewH, pageW, pageH, curRot);
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    // Offset from bottom-left of image to its center after rotation
    const dx = (wPt / 2) * cos - (hPt / 2) * sin;
    const dy = (wPt / 2) * sin + (hPt / 2) * cos;
    page.drawImage(pngImg, {
      x: centerPage.x - dx,
      y: centerPage.y - dy,
      width: wPt,
      height: hPt,
      rotate: degrees(totalDeg)
    });
  }

  async function applyPdfWatermark({ text, fontSize, color, opacity, angle, tiled, rangeMode, customRange }) {
    if (!curPdf?.bytes || !text.trim()) return;
    busy(true, say('جارٍ إضافة العلامة المائية…', 'Applying watermark…'));
    try {
      const { PDFDocument } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const pages = pdfDoc.getPages();
      const total = pages.length;
      const targetIndices = rangeMode === 'current'
        ? [curPdf.pageIdx || 0]
        : (rangeMode === 'custom' ? parsePageRanges(customRange, total) : Array.from({ length: total }, (_, i) => i));

      if (!targetIndices.length) {
        toast(say('نطاق الصفحات غير صالح', 'Invalid page range'));
        return;
      }

      const stamp = await renderStampTextToPng(text.trim(), fontSize, color, opacity, true);
      const pngImg = await pdfDoc.embedPng(stamp.bytes);

      for (const idx of targetIndices) {
        const page = pages[idx];
        if (!page) continue;
        const { width, height } = page.getSize();
        const rot = curPdf.rotations[idx] || 0;
        const viewW = rot % 180 ? height : width;
        const viewH = rot % 180 ? width : height;

        if (tiled) {
          const stepX = Math.max(stamp.widthPt * 1.45, viewW / 2.6);
          const stepY = Math.max(stamp.heightPt * 2.8, viewH / 3.4);
          for (let y = stepY * 0.45; y < viewH; y += stepY) {
            for (let x = stepX * 0.45; x < viewW; x += stepX) {
              drawCenteredRotatedStamp(page, pngImg, x, y, stamp.widthPt, stamp.heightPt, angle, viewW, viewH, width, height, rot);
            }
          }
        } else {
          drawCenteredRotatedStamp(page, pngImg, viewW / 2, viewH / 2, stamp.widthPt, stamp.heightPt, angle, viewW, viewH, width, height, rot);
        }
      }

      curPdf.bytes = new Uint8Array(await pdfDoc.save({ useObjectStreams: false, updateMetadata: false }));
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      await renderPdfEditor();
      toast(say('تمت إضافة العلامة المائية بنجاح ✓', 'Watermark applied successfully ✓'));
    } catch (err) {
      console.error('Watermark failed:', err);
      toast(say('تعذر إضافة العلامة المائية', 'Failed to apply watermark'));
    } finally {
      busy(false);
    }
  }

  async function applyPdfPageNumbers({ format, position, startNum, skipFirst, fontSize, color }) {
    if (!curPdf?.bytes) return;
    busy(true, say('جارٍ ترقيم الصفحات…', 'Adding page numbers…'));
    try {
      const { PDFDocument } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const pages = pdfDoc.getPages();
      const total = pages.length;
      const start = Math.max(1, parseInt(startNum, 10) || 1);
      const totalDisplayed = start + total - 1 - (skipFirst ? 1 : 0);

      for (let i = 0; i < total; i++) {
        if (skipFirst && i === 0) continue;
        const page = pages[i];
        const num = start + i - (skipFirst ? 1 : 0);
        let label = String(num);
        if (format === 'n_of_total') label = `${num} / ${totalDisplayed}`;
        else if (format === 'page_en') label = `Page ${num} of ${totalDisplayed}`;
        else if (format === 'page_ar') label = `صفحة ${num} من ${totalDisplayed}`;

        const stamp = await renderStampTextToPng(label, fontSize, color, 0.92, false);
        const pngImg = await pdfDoc.embedPng(stamp.bytes);
        const { width, height } = page.getSize();
        const rot = curPdf.rotations[i] || 0;
        const viewW = rot % 180 ? height : width;
        const viewH = rot % 180 ? width : height;
        const marginX = 36 + stamp.widthPt / 2;
        const marginY = 26;

        let cx = viewW / 2;
        let cy = viewH - marginY;
        if (position.startsWith('top')) cy = marginY;
        if (position.endsWith('left')) cx = marginX;
        else if (position.endsWith('right')) cx = viewW - marginX;

        drawCenteredRotatedStamp(page, pngImg, cx, cy, stamp.widthPt, stamp.heightPt, 0, viewW, viewH, width, height, rot);
      }

      curPdf.bytes = new Uint8Array(await pdfDoc.save({ useObjectStreams: false, updateMetadata: false }));
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      await renderPdfEditor();
      toast(say('تم ترقيم صفحات PDF بنجاح ✓', 'Page numbers added successfully ✓'));
    } catch (err) {
      console.error('Page numbering failed:', err);
      toast(say('تعذر ترقيم الصفحات', 'Failed to add page numbers'));
    } finally {
      busy(false);
    }
  }

  function openPdfWatermarkModal(initialTab = 'watermark') {
    if (kind !== 'pdf' || !curPdf?.bytes) {
      toast(say('افتح ملف PDF أولاً', 'Open a PDF file first'));
      return;
    }
    let activeTab = initialTab;
    const modal = openPhaseModal('💧 ' + say('العلامة المائية وترقيم الصفحات', 'Watermark & Page Numbers'));

    const renderBody = () => {
      const tabWm = h('button', {
        type: 'button',
        class: 'btn ' + (activeTab === 'watermark' ? 'pri' : ''),
        onclick: () => { activeTab = 'watermark'; renderBody(); }
      }, '💧 ' + say('علامة مائية', 'Watermark'));

      const tabNum = h('button', {
        type: 'button',
        class: 'btn ' + (activeTab === 'numbers' ? 'pri' : ''),
        onclick: () => { activeTab = 'numbers'; renderBody(); }
      }, '🔢 ' + say('ترقيم الصفحات', 'Page Numbers'));

      const tabs = h('div', { class: 'p1-tabs' }, tabWm, tabNum);

      if (activeTab === 'watermark') {
        const textIn = h('input', { type: 'text', value: say('سري / مسودة', 'CONFIDENTIAL') });
        const sizeIn = h('input', { type: 'number', min: 14, max: 140, value: '48' });
        const colorIn = h('input', { type: 'color', value: '#b4531a' });
        const opIn = h('input', { type: 'range', min: 8, max: 75, value: '22' });
        const opVal = h('span', { class: 'p1-range-val' }, '22%');
        opIn.oninput = () => { opVal.textContent = opIn.value + '%'; };
        const angleSel = h('select', null,
          h('option', { value: '-45', selected: true }, '-45° (' + say('مائل', 'Diagonal') + ')'),
          h('option', { value: '0' }, '0° (' + say('أفقي', 'Horizontal') + ')'),
          h('option', { value: '45' }, '45° (' + say('مائل عكسي', 'Reverse Diagonal') + ')')
        );
        const layoutSel = h('select', null,
          h('option', { value: 'center', selected: true }, say('وسط الصفحة', 'Centered')),
          h('option', { value: 'tiled' }, say('شبكة متكررة (Tiled)', 'Tiled Grid'))
        );
        const rangeSel = h('select', null,
          h('option', { value: 'all', selected: true }, say('كل الصفحات', 'All pages')),
          h('option', { value: 'current' }, say('الصفحة الحالية فقط', 'Current page only')),
          h('option', { value: 'custom' }, say('نطاق مخصص…', 'Custom range…'))
        );
        const customIn = h('input', { type: 'text', placeholder: '1-3, 5', hidden: true });
        rangeSel.onchange = () => { customIn.hidden = rangeSel.value !== 'custom'; };

        const applyBtn = h('button', {
          type: 'button',
          class: 'btn pri',
          onclick: async () => {
            if (!textIn.value.trim()) { textIn.focus(); return; }
            modal.close();
            await applyPdfWatermark({
              text: textIn.value,
              fontSize: Math.max(14, Math.min(140, +sizeIn.value || 48)),
              color: colorIn.value,
              opacity: (+opIn.value || 22) / 100,
              angle: parseInt(angleSel.value, 10) || 0,
              tiled: layoutSel.value === 'tiled',
              rangeMode: rangeSel.value,
              customRange: customIn.value
            });
          }
        }, say('تطبيق العلامة المائية', 'Apply Watermark'));

        modal.render(
          tabs,
          h('div', { class: 'pt-field' }, h('label', null, say('نص العلامة المائية', 'Watermark text')), textIn),
          h('div', { class: 'p1-grid-2' },
            h('div', { class: 'pt-field' }, h('label', null, say('حجم الخط', 'Font size')), sizeIn),
            h('div', { class: 'pt-field' }, h('label', null, say('اللون', 'Color')), colorIn)
          ),
          h('div', { class: 'pt-field' },
            h('label', null, say('الشفافية', 'Opacity'), ' ', opVal),
            opIn
          ),
          h('div', { class: 'p1-grid-2' },
            h('div', { class: 'pt-field' }, h('label', null, say('الزاوية', 'Angle')), angleSel),
            h('div', { class: 'pt-field' }, h('label', null, say('التوزيع', 'Layout')), layoutSel)
          ),
          h('div', { class: 'pt-field' }, h('label', null, say('الصفحات', 'Pages')), rangeSel, customIn),
          h('div', { class: 'pt-actions' },
            h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إلغاء', 'Cancel')),
            applyBtn
          )
        );
        textIn.focus();
      } else {
        const fmtSel = h('select', null,
          h('option', { value: 'n_of_total', selected: true }, '1 / ' + (curPdf.pageCount || 1)),
          h('option', { value: 'n' }, '1, 2, 3…'),
          h('option', { value: 'page_en' }, 'Page 1 of ' + (curPdf.pageCount || 1)),
          h('option', { value: 'page_ar' }, 'صفحة 1 من ' + (curPdf.pageCount || 1))
        );
        const posSel = h('select', null,
          h('option', { value: 'bottom-center', selected: true }, say('أسفل الوسط', 'Bottom Center')),
          h('option', { value: 'bottom-right' }, say('أسفل اليمين', 'Bottom Right')),
          h('option', { value: 'bottom-left' }, say('أسفل اليسار', 'Bottom Left')),
          h('option', { value: 'top-center' }, say('أعلى الوسط', 'Top Center')),
          h('option', { value: 'top-right' }, say('أعلى اليمين', 'Top Right')),
          h('option', { value: 'top-left' }, say('أعلى اليسار', 'Top Left'))
        );
        const startIn = h('input', { type: 'number', min: 1, max: 9999, value: '1' });
        const sizeIn = h('input', { type: 'number', min: 8, max: 32, value: '11' });
        const colorIn = h('input', { type: 'color', value: '#3f3b35' });
        const skipChk = h('input', { type: 'checkbox' });

        const applyBtn = h('button', {
          type: 'button',
          class: 'btn pri',
          onclick: async () => {
            modal.close();
            await applyPdfPageNumbers({
              format: fmtSel.value,
              position: posSel.value,
              startNum: startIn.value,
              skipFirst: skipChk.checked,
              fontSize: Math.max(8, Math.min(32, +sizeIn.value || 11)),
              color: colorIn.value
            });
          }
        }, say('تطبيق الترقيم', 'Add Page Numbers'));

        modal.render(
          tabs,
          h('div', { class: 'p1-grid-2' },
            h('div', { class: 'pt-field' }, h('label', null, say('صيغة الترقيم', 'Format')), fmtSel),
            h('div', { class: 'pt-field' }, h('label', null, say('الموضع', 'Position')), posSel)
          ),
          h('div', { class: 'p1-grid-3' },
            h('div', { class: 'pt-field' }, h('label', null, say('يبدأ من', 'Start at')), startIn),
            h('div', { class: 'pt-field' }, h('label', null, say('الحجم', 'Size')), sizeIn),
            h('div', { class: 'pt-field' }, h('label', null, say('اللون', 'Color')), colorIn)
          ),
          h('label', { class: 'pt-check' }, skipChk, h('span', null, say('تخطي الصفحة الأولى (صفحة الغلاف)', 'Skip first page (cover page)'))),
          h('div', { class: 'pt-actions' },
            h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إلغاء', 'Cancel')),
            applyBtn
          )
        );
      }
    };
    renderBody();
  }

  /* ═══════════════════════════════════════════════════════════
     4. True PDF Redaction (Surgical Strip + Optional Flatten)
     ═══════════════════════════════════════════════════════════ */
  let pendingRedactions = {}; // { [pageIdx]: [{ id, xPct, yPct, wPct, hPct, baseRot }] }
  let strictRedactMode = false;

  function countPendingRedactions() {
    let n = 0;
    for (const arr of Object.values(pendingRedactions)) n += (arr?.length || 0);
    return n;
  }

  function boxesOverlap(a, b) {
    return a.xPct < (b.xPct + b.wPct) && (a.xPct + a.wPct) > b.xPct &&
           a.yPct < (b.yPct + b.hPct) && (a.yPct + a.hPct) > b.yPct;
  }

  async function markRedactionsByText() {
    if (!curPdf?.bytes) return;
    const query = prompt(say('أدخل الكلمة أو العبارة المطلوب تعتيمها في كل الصفحات:', 'Enter word or phrase to redact across all pages:'));
    if (!query || !query.trim()) return;
    const needle = query.trim().toLowerCase();
    busy(true, say('جارٍ البحث عن النص وتحديده للتعتيم…', 'Searching text to mark for redaction…'));
    try {
      const pdfDoc = await loadPdfDocumentForView(curPdf.bytes);
      let found = 0;
      for (let pIdx = 0; pIdx < pdfDoc.numPages; pIdx++) {
        const page = await pdfDoc.getPage(pIdx + 1);
        const rot = curPdf.rotations[pIdx] || 0;
        const blocks = await extractPdfPageTextBlocks(page, rot, null);
        for (const blk of blocks) {
          if ((blk.origText || '').toLowerCase().includes(needle)) {
            if (!pendingRedactions[pIdx]) pendingRedactions[pIdx] = [];
            pendingRedactions[pIdx].push({
              id: 'rd_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
              xPct: Math.max(0, blk.xPct - 0.3),
              yPct: Math.max(0, blk.yPct - 0.3),
              wPct: Math.min(100, blk.wPct + 0.6),
              hPct: Math.min(100, blk.hPct + 0.6),
              baseRot: rot
            });
            found++;
          }
        }
      }
      await renderPdfEditor();
      if (found > 0) {
        toast(say(`تم تحديد ${found} موضع للتعتيم — اضغط "تطبيق التعتيم نهائياً"`, `Marked ${found} occurrence(s) — click "Apply Redactions"`));
      } else {
        toast(say('لم يتم العثور على النص المطلوب', 'Text not found in PDF'));
      }
    } catch (err) {
      console.error('Redact by text failed:', err);
      toast(say('تعذر البحث داخل ملف PDF', 'Failed to search PDF'));
    } finally {
      busy(false);
    }
  }

  async function applyTruePdfRedactions() {
    const totalMarks = countPendingRedactions();
    if (!curPdf?.bytes || !totalMarks) {
      toast(say('ارسم منطقة تعتيم واحدة على الأقل أولاً', 'Draw at least one redaction box first'));
      return;
    }
    const ok = window.confirm(
      say(
        `سيتم تطبيق ${totalMarks} منطقة تعتيم حقيقي وحذف النص الأساسي نهائياً من بيانات الـ PDF. هل تريد المتابعة؟`,
        `Apply ${totalMarks} permanent redaction(s) and strip underlying text from the PDF stream?`
      )
    );
    if (!ok) return;

    busy(true, say('جارٍ تطبيق التعتيم الحقيقي وحذف النص من الملف…', 'Applying true redaction and stripping text…'));
    try {
      const { PDFDocument, rgb } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const pdfJsDoc = await loadPdfDocumentForView(curPdf.bytes);
      const pages = pdfDoc.getPages();

      for (const [pKey, boxes] of Object.entries(pendingRedactions)) {
        const pIdx = parseInt(pKey, 10);
        if (!boxes?.length || isNaN(pIdx) || pIdx < 0 || pIdx >= pages.length) continue;
        const page = pages[pIdx];
        const rot = curPdf.rotations[pIdx] || 0;
        const projectedBoxes = boxes.map((b) => projectPdfBoxFromRot(b, b.baseRot ?? rot, rot));

        // Remove overlapping in-memory overlays (live unicode edits & floating text)
        if (curPdf.textEdits?.[pIdx]) {
          for (const [k, item] of Object.entries(curPdf.textEdits[pIdx])) {
            const pItem = projectPdfBoxFromRot(item, item.baseRot ?? rot, rot);
            if (projectedBoxes.some((rb) => boxesOverlap(rb, pItem))) delete curPdf.textEdits[pIdx][k];
          }
        }
        if (curPdf.texts?.[pIdx]) {
          curPdf.texts[pIdx] = curPdf.texts[pIdx].filter((t) => {
            const tb = { xPct: t.xPct, yPct: t.yPct, wPct: 14, hPct: 4 };
            return !projectedBoxes.some((rb) => boxesOverlap(rb, tb));
          });
        }

        if (strictRedactMode) {
          // Strict mode: rasterize page with black redaction boxes so even underlying images/paths are destroyed
          const jsPage = await pdfJsDoc.getPage(pIdx + 1);
          const vp = jsPage.getViewport({ scale: 2.0, rotation: rot });
          const cvs = document.createElement('canvas');
          cvs.width = Math.max(1, Math.round(vp.width));
          cvs.height = Math.max(1, Math.round(vp.height));
          const ctx = cvs.getContext('2d');
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, cvs.width, cvs.height);
          await jsPage.render({ canvasContext: ctx, viewport: vp }).promise;
          ctx.fillStyle = '#000000';
          for (const rb of projectedBoxes) {
            ctx.fillRect(
              (rb.xPct / 100) * cvs.width,
              (rb.yPct / 100) * cvs.height,
              (rb.wPct / 100) * cvs.width,
              (rb.hPct / 100) * cvs.height
            );
          }
          const pngBlob = await new Promise((res) => cvs.toBlob(res, 'image/png'));
          const pngImg = await pdfDoc.embedPng(new Uint8Array(await pngBlob.arrayBuffer()));
          const { width, height } = page.getSize();
          const viewW = rot % 180 ? height : width;
          const viewH = rot % 180 ? width : height;
          // Replace page with clean rasterized page
          const newPage = pdfDoc.insertPage(pIdx, [viewW, viewH]);
          newPage.drawImage(pngImg, { x: 0, y: 0, width: viewW, height: viewH });
          pdfDoc.removePage(pIdx + 1);
          curPdf.rotations[pIdx] = 0;
        } else {
          // Fast vector-preserving mode: surgically strip intersecting PDF text operators + burn solid black rect
          const jsPage = await pdfJsDoc.getPage(pIdx + 1);
          const blocks = await extractPdfPageTextBlocks(jsPage, rot, null);
          const matchedSegs = [];
          for (const blk of blocks) {
            if (projectedBoxes.some((rb) => boxesOverlap(rb, blk))) {
              if (blk.rawSegments?.length) matchedSegs.push(...blk.rawSegments);
            }
          }
          if (matchedSegs.length && typeof surgicallyStripPdfPageText === 'function') {
            surgicallyStripPdfPageText(pdfDoc, page, matchedSegs);
          }
          const { width, height } = page.getSize();
          for (const rb of projectedBoxes) {
            const c1 = pdfViewportPointToPage(rb.xPct / 100, rb.yPct / 100, width, height, rot);
            const c2 = pdfViewportPointToPage((rb.xPct + rb.wPct) / 100, (rb.yPct + rb.hPct) / 100, width, height, rot);
            page.drawRectangle({
              x: Math.min(c1.x, c2.x),
              y: Math.min(c1.y, c2.y),
              width: Math.abs(c2.x - c1.x),
              height: Math.abs(c2.y - c1.y),
              color: rgb(0, 0, 0)
            });
          }
        }
      }

      curPdf.bytes = new Uint8Array(await pdfDoc.save({ useObjectStreams: false, updateMetadata: false }));
      pendingRedactions = {};
      curPdf.tool = 'none';
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      await renderPdfEditor();
      toast(say('تم تطبيق التعتيم الحقيقي وحذف النص نهائياً ✓', 'True redaction applied & underlying text stripped ✓'));
    } catch (err) {
      console.error('True redaction failed:', err);
      toast(say('حدث خطأ أثناء تطبيق التعتيم الحقيقي', 'Failed to apply true redaction'));
    } finally {
      busy(false);
    }
  }

  function mountRedactionUi() {
    if (kind !== 'pdf' || !curPdf?.bytes) return;
    const curIdx = curPdf.pageIdx || 0;
    const curRot = curPdf.rotations[curIdx] || 0;
    const isRedact = curPdf.tool === 'redact';
    const totalMarks = countPendingRedactions();

    // 1. Redaction action bar above stage
    const mainCol = document.querySelector('.pdf-main');
    const stageWrap = document.querySelector('.pdf-stage-wrap');
    document.querySelector('.p1-redact-bar')?.remove();
    if (mainCol && stageWrap && (isRedact || totalMarks > 0)) {
      const strictChk = h('input', {
        type: 'checkbox',
        checked: strictRedactMode,
        onchange: () => { strictRedactMode = strictChk.checked; }
      });
      const bar = h('div', { class: 'p1-redact-bar' },
        h('span', { class: 'p1-redact-info' },
          '⬛ ' + (isRedact
            ? say('اسحب لتحديد مناطق التعتيم الحقيقي', 'Drag to mark true redaction areas')
            : say('مناطق تعتيم معلّقة', 'Pending redaction areas')),
          totalMarks ? h('strong', { class: 'p1-redact-count' }, ` (${totalMarks})`) : null
        ),
        h('button', { type: 'button', class: 'btn', onclick: () => markRedactionsByText() }, '🔎 ' + say('تعتيم نص محدد…', 'Redact Text…')),
        h('label', { class: 'p1-redact-strict', title: say('يحوّل الصفحات المعتمة إلى صور لإزالة الصور والرسوم المخفية أيضاً', 'Flattens redacted pages to images so underlying graphics/images are also destroyed') },
          strictChk,
          h('span', null, say('وضع صارم (تسطيح)', 'Strict flatten'))
        ),
        h('button', {
          type: 'button',
          class: 'btn danger',
          disabled: !totalMarks,
          onclick: () => applyTruePdfRedactions()
        }, say('تطبيق التعتيم نهائياً', 'Apply Redactions')),
        totalMarks ? h('button', {
          type: 'button',
          class: 'btn',
          onclick: () => { pendingRedactions = {}; renderPdfEditor(); }
        }, say('مسح التحديدات', 'Clear')) : null,
        isRedact ? h('button', {
          type: 'button',
          class: 'btn',
          onclick: () => { curPdf.tool = 'none'; renderPdfEditor(); }
        }, '✕') : null
      );
      mainCol.insertBefore(bar, stageWrap);
    }

    // 2. Redaction overlay layer on the page box
    const pageBox = document.querySelector('.pdf-page-box');
    if (!pageBox) return;
    pageBox.querySelector('.pdf-redact-layer')?.remove();
    const pageMarks = pendingRedactions[curIdx] || [];
    if (!isRedact && !pageMarks.length) return;

    const layer = h('div', { class: 'pdf-redact-layer' + (isRedact ? ' active-redact' : '') });
    const drawBoxes = () => {
      layer.innerHTML = '';
      for (const mark of (pendingRedactions[curIdx] || [])) {
        const proj = projectPdfBoxFromRot(mark, mark.baseRot ?? curRot, curRot);
        const box = h('div', {
          class: 'pdf-redact-box',
          style: {
            left: proj.xPct + '%',
            top: proj.yPct + '%',
            width: proj.wPct + '%',
            height: proj.hPct + '%'
          }
        },
          h('button', {
            type: 'button',
            class: 'pdf-redact-del',
            title: say('إزالة التحديد', 'Remove mark'),
            onclick: (ev) => {
              ev.preventDefault();
              ev.stopPropagation();
              pendingRedactions[curIdx] = (pendingRedactions[curIdx] || []).filter((m) => m.id !== mark.id);
              mountRedactionUi();
            }
          }, '✕')
        );
        layer.append(box);
      }
    };
    drawBoxes();

    if (isRedact) {
      let dragStart = null;
      let previewBox = null;
      layer.addEventListener('pointerdown', (ev) => {
        if (ev.target !== layer) return;
        ev.preventDefault();
        const r = layer.getBoundingClientRect();
        dragStart = {
          xPct: Math.max(0, Math.min(100, ((ev.clientX - r.left) / (r.width || 1)) * 100)),
          yPct: Math.max(0, Math.min(100, ((ev.clientY - r.top) / (r.height || 1)) * 100))
        };
        previewBox = h('div', { class: 'pdf-redact-box preview' });
        layer.append(previewBox);
        layer.setPointerCapture(ev.pointerId);
      });
      layer.addEventListener('pointermove', (ev) => {
        if (!dragStart || !previewBox) return;
        const r = layer.getBoundingClientRect();
        const cx = Math.max(0, Math.min(100, ((ev.clientX - r.left) / (r.width || 1)) * 100));
        const cy = Math.max(0, Math.min(100, ((ev.clientY - r.top) / (r.height || 1)) * 100));
        previewBox.style.left = Math.min(dragStart.xPct, cx) + '%';
        previewBox.style.top = Math.min(dragStart.yPct, cy) + '%';
        previewBox.style.width = Math.abs(cx - dragStart.xPct) + '%';
        previewBox.style.height = Math.abs(cy - dragStart.yPct) + '%';
      });
      const finishDrag = (ev) => {
        if (!dragStart) return;
        const r = layer.getBoundingClientRect();
        const cx = Math.max(0, Math.min(100, ((ev.clientX - r.left) / (r.width || 1)) * 100));
        const cy = Math.max(0, Math.min(100, ((ev.clientY - r.top) / (r.height || 1)) * 100));
        const xPct = Math.min(dragStart.xPct, cx);
        const yPct = Math.min(dragStart.yPct, cy);
        const wPct = Math.abs(cx - dragStart.xPct);
        const hPct = Math.abs(cy - dragStart.yPct);
        dragStart = null;
        if (previewBox) { previewBox.remove(); previewBox = null; }
        if (wPct >= 0.8 && hPct >= 0.6) {
          if (!pendingRedactions[curIdx]) pendingRedactions[curIdx] = [];
          pendingRedactions[curIdx].push({
            id: 'rd_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
            xPct, yPct, wPct, hPct, baseRot: curRot
          });
          mountRedactionUi();
        }
      };
      layer.addEventListener('pointerup', finishDrag);
      layer.addEventListener('pointercancel', finishDrag);
    }
    pageBox.append(layer);
  }

  // Hook into renderPdfEditor to inject Redact, Watermark, and Page Numbers buttons + Redaction UI
  if (typeof renderPdfEditor === 'function') {
    const origRenderPdf = renderPdfEditor;
    renderPdfEditor = async function (...args) {
      const res = await origRenderPdf.apply(this, args);
      try {
        if (kind !== 'pdf' || !curPdf?.bytes) return res;
        const toolbar = document.querySelector('.pdf-toolbar');
        if (toolbar && !toolbar._phase1Added) {
          toolbar._phase1Added = true;
          const bRedact = btn(
            say('⬛ تعتيم حقيقي', '⬛ Redact'),
            () => {
              curPdf.tool = curPdf.tool === 'redact' ? 'none' : 'redact';
              curPdf.activeOrigEditRef = null;
              renderPdfEditor();
            },
            curPdf.tool === 'redact' ? 'on' : ''
          );
          const bWatermark = btn(say('💧 علامة مائية', '💧 Watermark'), () => openPdfWatermarkModal('watermark'));
          const bPageNums = btn(say('🔢 ترقيم الصفحات', '🔢 Page #'), () => openPdfWatermarkModal('numbers'));

          const btns = Array.from(toolbar.querySelectorAll('button'));
          const whiteoutBtn = btns.find((b) => /Whiteout|طمس|تغطية/i.test(b.textContent || ''));
          const signBtn = btns.find((b) => /Sign|توقيع/i.test(b.textContent || ''));
          const spacer = Array.from(toolbar.children).find((c) => c.tagName === 'SPAN' && c.style.flex === '1');

          if (whiteoutBtn && whiteoutBtn.parentNode === toolbar) {
            toolbar.insertBefore(bRedact, whiteoutBtn.nextSibling);
          } else if (spacer) {
            toolbar.insertBefore(bRedact, spacer);
          } else {
            toolbar.append(bRedact);
          }

          if (signBtn && signBtn.parentNode === toolbar) {
            toolbar.insertBefore(bWatermark, signBtn.nextSibling);
            toolbar.insertBefore(bPageNums, bWatermark.nextSibling);
          } else if (spacer) {
            toolbar.insertBefore(bWatermark, spacer);
            toolbar.insertBefore(bPageNums, spacer);
          } else {
            toolbar.append(bWatermark, bPageNums);
          }
        }
        mountRedactionUi();
      } catch (err) {
        console.warn('Phase 1 PDF toolbar integration failed:', err);
      }
      return res;
    };
  }

  /* ═══════════════════════════════════════════════════════════
     5. Command Palette (Ctrl+K / ⌘K)
     ═══════════════════════════════════════════════════════════ */
  function buildCommandList() {
    const inEditor = !$id('editor')?.hidden;
    const cmds = [];
    const add = (group, ar, en, keywords, shortcut, run) => {
      cmds.push({ group, ar, en, keywords: (keywords + ' ' + ar + ' ' + en).toLowerCase(), shortcut: shortcut || '', run });
    };

    if (inEditor) {
      if (activeFileHandle) {
        add('file', 'حفظ مباشر على الملف الأصلي', 'Save Directly to File', 'save disk direct حفظ مباشر', 'Ctrl+S', () => saveDirectToDisk());
      }
      if (typeof window.showSaveFilePicker === 'function') {
        add('file', 'حفظ باسم… (ملف محلي)', 'Save As… (Local File)', 'save as export حفظ باسم', '', () => saveAsLocalFile());
      }
      add('file', 'تحميل نسخة من الملف', 'Download File Copy', 'download export تحميل تنزيل', 'Ctrl+S', () => downloadEditedFile());
      add('file', 'تنظيف البيانات المخفية (Metadata)', 'Clean Hidden Metadata', 'privacy scrub clean metadata author خصوصية بيانات مخفية تنظيف', '', () => openMetadataScrubberModal());
      if (window.Phase2Tools) {
        add('file', 'استخراج كل الصور من الملف الحالي', 'Extract All Images from Current File', 'extract images media photos zip استخراج صور', '', () => window.Phase2Tools.openExtractImagesModal(true));
        if (kind === 'pdf' || kind === 'docx') {
          add('file', 'مقارنة الملف الحالي مع ملف آخر', 'Compare Current File with Another', 'compare diff contract مقارنة فروقات عقد', '', () => window.Phase2Tools.openCompareDocsModal(true));
        }
      }
      add('file', 'البحث والاستبدال / البحث في النص', 'Find & Replace / Search', 'find search replace بحث استبدال', 'Ctrl+F', () => findAndReplace());
      add('file', 'تراجع (Undo)', 'Undo', 'undo تراجع', 'Ctrl+Z', () => $id('bUndo')?.click());
      add('file', 'إعادة (Redo)', 'Redo', 'redo إعادة', 'Ctrl+Y', () => $id('bRedo')?.click());
      add('file', 'وضع التركيز', 'Toggle Focus Mode', 'focus zen تركيز', '', () => $id('bFocusMode')?.click());

      if (kind === 'pdf') {
        add('editor', 'تعتيم حقيقي (حذف النص نهائياً)', 'True PDF Redaction', 'redact censor black hide تعتيم حقيقي حذف نص', '', () => {
          curPdf.tool = 'redact';
          renderPdfEditor();
        });
        add('editor', 'إضافة علامة مائية للـ PDF', 'Add PDF Watermark', 'watermark stamp confidential علامة مائية سري', '', () => openPdfWatermarkModal('watermark'));
        add('editor', 'إضافة ترقيم صفحات للـ PDF', 'Add PDF Page Numbers', 'page numbers footer header ترقيم صفحات', '', () => openPdfWatermarkModal('numbers'));
        if (window.Phase2Tools) {
          add('editor', 'توحيد مقاس صفحات PDF إلى A4', 'Normalize PDF Pages to A4', 'normalize a4 letter page size توحيد مقاس', '', () => window.Phase2Tools.openPdfPageSizeCropModal('normalize'));
          add('editor', 'قص هوامش صفحات PDF', 'Crop PDF Page Margins', 'crop margins trim قص هوامش', '', () => window.Phase2Tools.openPdfPageSizeCropModal('crop'));
        }
        add('editor', 'توقيع رقمي / إلكتروني', 'Digital Signature', 'sign signature توقيع رقمي', '', () => typeof openPdfSignatureModal === 'function' && openPdfSignatureModal());
        add('editor', 'تعديل النص الأصلي في PDF', 'Edit Original PDF Text', 'edit text stream تعديل نص', '', () => { curPdf.tool = 'edit-text'; renderPdfEditor(); });
        add('editor', 'دمج ملفات PDF', 'Merge PDF Files', 'merge combine دمج', '', () => $id('pdfMergeIn')?.click());
        add('editor', 'تحويل PDF إلى Word (.docx)', 'Convert PDF to Word', 'word docx convert تحويل وورد', '', () => typeof convertCurrentPdfToDocx === 'function' && convertCurrentPdfToDocx());
        if (window.PdfTools) {
          add('editor', 'ضغط حجم ملف PDF', 'Compress Current PDF', 'compress shrink size ضغط حجم', '', async () => {
            const bytes = await buildMergedPdfBytesWithAnnotations();
            window.PdfTools.compress(bytes, fileName || 'document.pdf', {});
          });
          add('editor', 'حماية PDF بكلمة مرور (AES-256)', 'Protect PDF with Password', 'protect encrypt password lock حماية كلمة مرور تشفير', '', async () => {
            const bytes = await buildMergedPdfBytesWithAnnotations();
            window.PdfTools.protect(bytes, fileName || 'document.pdf', {});
          });
          add('editor', 'تعبئة حقول نموذج PDF', 'Fill PDF Form Fields', 'fill form acroform تعبئة نموذج', '', async () => {
            const bytes = await buildMergedPdfBytesWithAnnotations();
            window.PdfTools.fill(bytes, fileName || 'document.pdf', {});
          });
        }
      } else if (kind === 'docx') {
        add('editor', 'طباعة المستند', 'Print Document', 'print طباعة', 'Ctrl+P', () => window.print());
        if (typeof exportDocxToPdf === 'function') {
          add('editor', 'تصدير Word إلى PDF', 'Export Word to PDF', 'export pdf تصدير', '', () => exportDocxToPdf());
        }
        if (typeof convertDocxToPdfEditor === 'function') {
          add('editor', 'فتح المستند في محرر PDF', 'Open in PDF Editor', 'convert pdf محرر', '', () => convertDocxToPdfEditor());
        }
      } else if (kind === 'pptx') {
        if (typeof startSlideshow === 'function') {
          add('editor', 'بدء العرض التقديمي (ملء الشاشة)', 'Start Slideshow', 'present slideshow play عرض تقديمي', '', () => startSlideshow(curPptx?.slideIdx || 0));
        }
        if (typeof openPresenterView === 'function') {
          add('editor', 'وضع مقدم العرض', 'Open Presenter View', 'presenter notes مقدم العرض', '', () => openPresenterView());
        }
        if (typeof addNewSlide === 'function' && curPptx?.deck) {
          add('editor', 'إضافة شريحة جديدة', 'Add New Slide', 'new slide شريحة جديدة', '', () => addNewSlide(curPptx.deck, curPptx.slideIdx));
        }
      }
      add('nav', 'العودة إلى الصفحة الرئيسية', 'Go to Home', 'home back رئيسية', '', () => showView('home'));
    }

    // Global actions (always available)
    add('tools', 'فتح ملف من الجهاز (مع الحفظ المباشر)', 'Open Local File (Direct Save)', 'open file browse فتح ملف', '', () => openLocalFileDirect());
    add('tools', 'إنشاء مستند Word جديد', 'New Word Document', 'new docx word مستند جديد', '', () => startNewDocx('blank'));
    add('tools', 'إنشاء عرض PowerPoint جديد', 'New Presentation (PPTX)', 'new pptx slide deck عرض جديد', '', () => startNewDeck('light'));
    add('tools', 'تحويل الصور إلى PDF', 'Images to PDF', 'images jpg png pdf صور', '', () => $id('bImagesToPdf')?.click());
    add('tools', 'تقسيم ملف PDF', 'Split PDF', 'split extract تقسيم', '', () => $id('bSplitPdf')?.click());
    add('tools', 'ضغط ملف PDF', 'Compress PDF', 'compress ضغط', '', () => $id('bCompressPdf')?.click());
    add('tools', 'حماية PDF بكلمة مرور', 'Protect PDF', 'protect password lock حماية', '', () => $id('bProtectPdf')?.click());
    add('tools', 'فك حماية PDF', 'Unlock PDF', 'unlock decrypt فك حماية', '', () => $id('bUnlockPdf')?.click());
    add('tools', 'تعبئة نموذج PDF', 'Fill PDF Form', 'fill form تعبئة نموذج', '', () => $id('bFillPdf')?.click());
    if (window.Phase2Tools) {
      add('tools', 'مقارنة مستندين (PDF / Word)', 'Compare Two Documents (Diff)', 'compare diff pdf word مقارنة مستندين فروقات', '', () => window.Phase2Tools.openCompareDocsModal(false));
      add('tools', 'استخراج الصور من أي ملف (DOCX / PPTX / PDF)', 'Extract Images from File', 'extract images media zip استخراج صور', '', () => window.Phase2Tools.openExtractImagesModal(false));
    }

    add('prefs', 'تبديل الوضع الليلي / النهاري', 'Toggle Light / Dark Theme', 'theme dark light mode ليلي نهاري مظهر', '', () => {
      applyTheme(currentTheme === 'dark' ? 'light' : 'dark');
    });
    add('prefs', 'تبديل اللغة (العربية / English)', 'Switch Language (EN / AR)', 'language arabic english لغة عربي إنجليزي', '', () => {
      const next = currentLanguage === 'ar' ? 'en' : 'ar';
      applyLanguage(next);
      try { localStorage.setItem('doc-studio-language', next); } catch (_) {}
    });

    return cmds;
  }

  let paletteOpen = false;
  function openCommandPalette() {
    const existing = $id('cmdPaletteBackdrop');
    if (existing) {
      existing.remove();
      paletteOpen = false;
      return;
    }
    paletteOpen = true;
    const allCmds = buildCommandList();
    let filtered = allCmds;
    let activeIdx = 0;

    const searchIn = h('input', {
      type: 'text',
      class: 'p1-cmd-input',
      placeholder: say('ابحث عن أداة أو أمر… (مثال: علامة مائية، تعتيم، بيانات، ضغط)', 'Search commands… (e.g. watermark, redact, metadata, compress)'),
      autocomplete: 'off',
      spellcheck: 'false'
    });
    const listEl = h('div', { class: 'p1-cmd-list', role: 'listbox' });

    const close = () => {
      backdrop.remove();
      paletteOpen = false;
    };

    const runSelected = (idx) => {
      const item = filtered[idx];
      if (!item) return;
      close();
      try { item.run(); } catch (err) { console.error('Command failed:', err); }
    };

    const groupNames = {
      editor: say('أدوات المحرر الحالي', 'Current Editor'),
      file: say('الملف والخصوصية', 'File & Privacy'),
      tools: say('أدوات سريعة', 'Quick Tools'),
      prefs: say('المظهر واللغة', 'Preferences'),
      nav: say('التنقل', 'Navigation')
    };

    const renderList = () => {
      const q = searchIn.value.trim().toLowerCase();
      filtered = q
        ? allCmds.filter((c) => q.split(/\s+/).every((tok) => c.keywords.includes(tok)))
        : allCmds;
      if (activeIdx >= filtered.length) activeIdx = Math.max(0, filtered.length - 1);

      listEl.innerHTML = '';
      if (!filtered.length) {
        listEl.append(h('div', { class: 'p1-cmd-empty' }, say('لا توجد أوامر مطابقة', 'No matching commands')));
        return;
      }

      let lastGroup = null;
      filtered.forEach((cmd, idx) => {
        if (cmd.group !== lastGroup) {
          lastGroup = cmd.group;
          listEl.append(h('div', { class: 'p1-cmd-group' }, groupNames[cmd.group] || cmd.group));
        }
        const row = h('button', {
          type: 'button',
          class: 'p1-cmd-item' + (idx === activeIdx ? ' active' : ''),
          role: 'option',
          onmouseenter: () => {
            if (activeIdx !== idx) {
              listEl.querySelector('.p1-cmd-item.active')?.classList.remove('active');
              activeIdx = idx;
              row.classList.add('active');
            }
          },
          onclick: () => runSelected(idx)
        },
          h('span', { class: 'p1-cmd-label' }, isAr() ? cmd.ar : cmd.en),
          cmd.shortcut ? h('kbd', { class: 'p1-cmd-kbd' }, cmd.shortcut) : null
        );
        listEl.append(row);
      });
      listEl.querySelector('.p1-cmd-item.active')?.scrollIntoView({ block: 'nearest' });
    };

    searchIn.addEventListener('input', () => { activeIdx = 0; renderList(); });
    searchIn.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        if (filtered.length) { activeIdx = (activeIdx + 1) % filtered.length; renderList(); }
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        if (filtered.length) { activeIdx = (activeIdx - 1 + filtered.length) % filtered.length; renderList(); }
      } else if (e.key === 'Enter') {
        e.preventDefault();
        runSelected(activeIdx);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        close();
      }
    });

    const box = h('div', { class: 'p1-cmd-box' },
      h('div', { class: 'p1-cmd-head' },
        h('span', { class: 'p1-cmd-search-ic', 'aria-hidden': 'true' }, '⌘'),
        searchIn,
        h('kbd', { class: 'p1-cmd-esc' }, 'ESC')
      ),
      listEl
    );

    const backdrop = h('div', {
      id: 'cmdPaletteBackdrop',
      class: 'p1-cmd-backdrop',
      onclick: (e) => { if (e.target === backdrop) close(); }
    }, box);

    document.body.append(backdrop);
    renderList();
    searchIn.focus();
  }

  // Global Ctrl+K / Cmd+K shortcut
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
      if (!$id('slideshowOverlay')?.hidden) return;
      e.preventDefault();
      openCommandPalette();
    }
  });

  /* ═══════════════════════════════════════════════════════════
     6. Mount Topbar & Home Trigger Buttons
     ═══════════════════════════════════════════════════════════ */
  function mountShellButtons() {
    // Home nav Ctrl+K button
    const navRight = document.querySelector('.v2-nav-right');
    if (navRight && !$id('bCmdHome')) {
      const cmdHome = h('button', {
        type: 'button',
        id: 'bCmdHome',
        class: 'v2-cmd-pill',
        title: say('لوحة الأوامر السريعة (Ctrl+K)', 'Quick Command Palette (Ctrl+K)'),
        onclick: () => openCommandPalette()
      },
        h('span', null, say('أوامر سريعة', 'Commands')),
        h('kbd', null, 'Ctrl+K')
      );
      navRight.insertBefore(cmdHome, navRight.firstChild);
    }

    // Editor topbar buttons: Ctrl+K, Metadata Scrubber, Direct Save
    const topBar = document.querySelector('#editor .top');
    const downBtn = $id('bDown');
    if (topBar) {
      if (!$id('bCmdEditor')) {
        const cmdBtn = btn('⌘ Ctrl+K', () => openCommandPalette());
        cmdBtn.id = 'bCmdEditor';
        cmdBtn.title = say('لوحة الأوامر السريعة (Ctrl+K)', 'Quick Command Palette (Ctrl+K)');
        const findBtn = $id('bFind');
        if (findBtn) topBar.insertBefore(cmdBtn, findBtn);
        else if (downBtn) topBar.insertBefore(cmdBtn, downBtn);
        else topBar.append(cmdBtn);
      }

      if (!$id('bScrubMeta')) {
        const metaBtn = btn('🛡️ تنظيف البيانات', () => openMetadataScrubberModal());
        metaBtn.id = 'bScrubMeta';
        metaBtn.title = say('فحص وتنظيف البيانات المخفية (Metadata)', 'Inspect & clean hidden metadata');
        const focusBtn = $id('bFocusMode');
        if (focusBtn) topBar.insertBefore(metaBtn, focusBtn);
        else if (downBtn) topBar.insertBefore(metaBtn, downBtn);
        else topBar.append(metaBtn);
      }

      if (!$id('bDirectSave')) {
        const saveBtn = btn('💾 حفظ مباشر', () => saveDirectToDisk(), 'pri');
        saveBtn.id = 'bDirectSave';
        saveBtn.hidden = !activeFileHandle;
        if (downBtn) topBar.insertBefore(saveBtn, downBtn);
        else topBar.append(saveBtn);
      }
    }
    updateDirectSaveUi();
  }

  mountShellButtons();

  // Expose clean API
  window.Phase1Features = {
    openLocalFileDirect,
    saveDirectToDisk,
    saveAsLocalFile,
    openMetadataScrubberModal,
    openPdfWatermarkModal,
    applyTruePdfRedactions,
    openCommandPalette
  };
})();
