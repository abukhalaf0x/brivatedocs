'use strict';
/* ══════════════════════════════════════════════════════════════
   brivatedocs — Practical Document Tools (100% Client-Side)
   1. Document Diff (PDF & DOCX side-by-side / inline comparison)
   2. Extract All Images (DOCX, PPTX, PDF -> Gallery + ZIP download)
   3. PDF Normalize to A4 / Letter & Crop Margins (lossless vector)
   ══════════════════════════════════════════════════════════════ */
(() => {
  const isAr = () => currentLanguage === 'ar';
  const say = (ar, en) => (isAr() ? ar : en);
  const $id = (id) => document.getElementById(id);

  function openToolModal(title, opts = {}) {
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
    const box = h('div', { class: 'md-modal pt-modal' + (opts.wide ? ' p2-modal-wide' : '') },
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

  function pickFile(accept) {
    return new Promise((resolve) => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = accept;
      inp.hidden = true;
      document.body.append(inp);
      const done = (f) => { inp.remove(); resolve(f); };
      inp.onchange = () => done((inp.files && inp.files[0]) || null);
      inp.oncancel = () => done(null);
      inp.click();
    });
  }

  const fmtSize = (n) => n >= 1048576 ? (n / 1048576).toFixed(2) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';

  /* ═══════════════════════════════════════════════════════════
     1. Document Diff (PDF & DOCX)
     ═══════════════════════════════════════════════════════════ */
  async function extractLinesFromPdfBytes(bytes) {
    const pdfDoc = await loadPdfDocumentForView(bytes);
    const lines = [];
    for (let i = 1; i <= pdfDoc.numPages; i++) {
      const page = await pdfDoc.getPage(i);
      const tc = await page.getTextContent();
      let curLine = '';
      let lastY = null;
      for (const item of (tc.items || [])) {
        const str = item.str || '';
        const y = item.transform ? Math.round(item.transform[5] / 3) * 3 : null;
        if (lastY !== null && y !== null && Math.abs(y - lastY) > 5) {
          if (curLine.trim()) lines.push(curLine.trim());
          curLine = str;
        } else {
          if (curLine && str && !curLine.endsWith(' ') && !str.startsWith(' ')) curLine += ' ';
          curLine += str;
        }
        if (y !== null) lastY = y;
      }
      if (curLine.trim()) lines.push(curLine.trim());
    }
    return lines;
  }

  async function extractLinesFromDocxBytes(bytes) {
    const zip = await window.JSZip.loadAsync(bytes);
    const docEntry = zip.file('word/document.xml');
    if (!docEntry) throw new UserError(say('ملف Word غير صالح', 'Invalid Word (.docx) file'));
    const xmlStr = await docEntry.async('string');
    const dom = new DOMParser().parseFromString(xmlStr, 'application/xml');
    const paras = Array.from(dom.getElementsByTagName('*')).filter((el) => el.localName === 'p');
    const lines = [];
    for (const p of paras) {
      let text = '';
      const nodes = Array.from(p.getElementsByTagName('*'));
      for (const n of nodes) {
        if (n.localName === 't') text += n.textContent || '';
        else if (n.localName === 'tab') text += '\t';
        else if (n.localName === 'br') text += '\n';
      }
      for (const sub of text.split('\n')) {
        if (sub.trim()) lines.push(sub.trim());
      }
    }
    return lines;
  }

  async function extractLinesFromFile(file) {
    assertFileSize(file);
    const bytes = new Uint8Array(await file.arrayBuffer());
    const name = (file.name || '').toLowerCase();
    if (name.endsWith('.pdf') || file.type === 'application/pdf') {
      return extractLinesFromPdfBytes(bytes);
    }
    if (name.endsWith('.docx') || name.endsWith('.docm')) {
      return extractLinesFromDocxBytes(bytes);
    }
    throw new UserError(say('يدعم المقارنة لملفات PDF و Word (.docx) فقط', 'Comparison supports PDF and Word (.docx) files only'));
  }

  async function extractLinesFromActiveDoc() {
    if (kind === 'pdf' && curPdf?.bytes) {
      return extractLinesFromPdfBytes(curPdf.bytes);
    }
    if (kind === 'docx' && pk) {
      const paper = $('#app .paper');
      if (paper && typeof syncPaperToDocxXml === 'function') syncPaperToDocxXml(paper);
      pk.commit();
      const xmlStr = pk.files.get('word/document.xml');
      if (typeof xmlStr === 'string') {
        const dom = new DOMParser().parseFromString(xmlStr, 'application/xml');
        const paras = Array.from(dom.getElementsByTagName('*')).filter((el) => el.localName === 'p');
        return paras.map((p) => Array.from(p.getElementsByTagName('*')).filter((n) => n.localName === 't').map((n) => n.textContent || '').join('').trim()).filter(Boolean);
      }
    }
    return null;
  }

  // LCS-based line diff (caps DP table gracefully for huge documents by chunking)
  function computeLineDiff(aLines, bLines) {
    const MAX_DP = 1400;
    if (aLines.length * bLines.length > MAX_DP * MAX_DP) {
      // Fast greedy sync fallback for very large documents
      const out = [];
      let i = 0, j = 0;
      const bMap = new Map();
      bLines.forEach((ln, idx) => {
        if (!bMap.has(ln)) bMap.set(ln, []);
        bMap.get(ln).push(idx);
      });
      while (i < aLines.length && j < bLines.length) {
        if (aLines[i] === bLines[j]) {
          out.push({ type: 'eq', text: aLines[i] });
          i++; j++;
        } else {
          const nextInB = (bMap.get(aLines[i]) || []).find((idx) => idx >= j && idx < j + 40);
          if (nextInB !== undefined) {
            while (j < nextInB) { out.push({ type: 'add', text: bLines[j++] }); }
          } else {
            out.push({ type: 'del', text: aLines[i++] });
          }
        }
      }
      while (i < aLines.length) out.push({ type: 'del', text: aLines[i++] });
      while (j < bLines.length) out.push({ type: 'add', text: bLines[j++] });
      return out;
    }

    const m = aLines.length, n = bLines.length;
    const dp = Array.from({ length: m + 1 }, () => new Uint32Array(n + 1));
    for (let i = m - 1; i >= 0; i--) {
      for (let j = n - 1; j >= 0; j--) {
        if (aLines[i] === bLines[j]) dp[i][j] = dp[i + 1][j + 1] + 1;
        else dp[i][j] = dp[i + 1][j] >= dp[i][j + 1] ? dp[i + 1][j] : dp[i][j + 1];
      }
    }
    const out = [];
    let i = 0, j = 0;
    while (i < m && j < n) {
      if (aLines[i] === bLines[j]) {
        out.push({ type: 'eq', text: aLines[i] });
        i++; j++;
      } else if (dp[i + 1][j] >= dp[i][j + 1]) {
        out.push({ type: 'del', text: aLines[i++] });
      } else {
        out.push({ type: 'add', text: bLines[j++] });
      }
    }
    while (i < m) out.push({ type: 'del', text: aLines[i++] });
    while (j < n) out.push({ type: 'add', text: bLines[j++] });
    return out;
  }

  function openCompareDocsModal(useCurrentAsA = false) {
    const modal = openToolModal('⚖️ ' + say('مقارنة مستندين (PDF / Word)', 'Compare Two Documents (PDF / Word)'), { wide: true });
    let fileA = null;
    let fileB = null;
    let linesA = null;
    let linesB = null;
    let nameA = '';
    let nameB = '';
    let diffResult = null;
    let changesOnly = false;

    const canUseCurrent = !$id('editor')?.hidden && (kind === 'pdf' || kind === 'docx') && (curPdf?.bytes || pk);

    const renderUi = () => {
      const btnA = h('button', {
        type: 'button',
        class: 'btn ' + (nameA ? '' : 'pri'),
        onclick: async () => {
          const f = await pickFile('.pdf,.docx,.docm,application/pdf');
          if (!f) return;
          fileA = f;
          nameA = f.name;
          linesA = null;
          diffResult = null;
          renderUi();
        }
      }, nameA ? ('📄 ' + nameA) : say('اختر الملف الأصلي (1)', 'Choose Original File (1)'));

      const btnB = h('button', {
        type: 'button',
        class: 'btn ' + (nameB ? '' : 'pri'),
        onclick: async () => {
          const f = await pickFile('.pdf,.docx,.docm,application/pdf');
          if (!f) return;
          fileB = f;
          nameB = f.name;
          linesB = null;
          diffResult = null;
          renderUi();
        }
      }, nameB ? ('📄 ' + nameB) : say('اختر الملف المعدّل (2)', 'Choose Modified File (2)'));

      const useCurBtn = canUseCurrent && !useCurrentAsA ? h('button', {
        type: 'button',
        class: 'btn',
        onclick: () => {
          useCurrentAsA = true;
          fileA = null;
          nameA = (fileName || 'Current') + ' (' + say('الحالي', 'Current') + ')';
          linesA = null;
          diffResult = null;
          renderUi();
        }
      }, say('استخدم الملف المفتوح كملف (1)', 'Use open file as (1)')) : null;

      const runBtn = h('button', {
        type: 'button',
        class: 'btn pri',
        disabled: !(nameA && nameB),
        onclick: async () => {
          busy(true, say('جارٍ مقارنة المستندين…', 'Comparing documents…'));
          try {
            if (!linesA) linesA = useCurrentAsA && !fileA ? await extractLinesFromActiveDoc() : await extractLinesFromFile(fileA);
            if (!linesB) linesB = await extractLinesFromFile(fileB);
            diffResult = computeLineDiff(linesA || [], linesB || []);
            renderUi();
          } catch (err) {
            console.error('Compare failed:', err);
            toast(err instanceof UserError ? err.message : say('تعذر مقارنة الملفين', 'Failed to compare files'));
          } finally {
            busy(false);
          }
        }
      }, '⚖️ ' + say('مقارنة الآن', 'Compare Now'));

      let resultsBlock = null;
      if (diffResult) {
        const added = diffResult.filter((d) => d.type === 'add').length;
        const removed = diffResult.filter((d) => d.type === 'del').length;
        const same = diffResult.filter((d) => d.type === 'eq').length;
        const visible = changesOnly ? diffResult.filter((d) => d.type !== 'eq') : diffResult;

        const filterChk = h('input', {
          type: 'checkbox',
          checked: changesOnly,
          onchange: () => { changesOnly = filterChk.checked; renderUi(); }
        });

        resultsBlock = h('div', { class: 'p2-diff-wrap' },
          h('div', { class: 'p2-diff-stats' },
            h('span', { class: 'p2-stat-add' }, `+${added} ` + say('إضافة', 'added')),
            h('span', { class: 'p2-stat-del' }, `-${removed} ` + say('حذف', 'removed')),
            h('span', { class: 'p2-stat-eq' }, `${same} ` + say('متطابق', 'unchanged')),
            h('label', { class: 'p2-diff-toggle' }, filterChk, h('span', null, say('عرض الفروقات فقط', 'Changes only')))
          ),
          visible.length
            ? h('div', { class: 'p2-diff-list' },
                ...visible.map((row) =>
                  h('div', {
                    class: 'p2-diff-line ' + row.type,
                    dir: /[\u0590-\u08FF]/.test(row.text) ? 'rtl' : 'ltr'
                  },
                    h('span', { class: 'p2-diff-mark' }, row.type === 'add' ? '+' : (row.type === 'del' ? '−' : ' ')),
                    h('span', { class: 'p2-diff-txt' }, row.text)
                  )
                )
              )
            : h('div', { class: 'p1-clean-badge' }, '✓ ' + say('المستندان متطابقان تماماً في النص!', 'Both documents have identical text!'))
        );
      }

      modal.render(
        h('div', { class: 'p2-compare-pickers' }, btnA, btnB, useCurBtn, runBtn),
        resultsBlock,
        h('div', { class: 'pt-actions' },
          h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close'))
        )
      );
    };

    if (useCurrentAsA && canUseCurrent) {
      nameA = (fileName || 'Current') + ' (' + say('الحالي', 'Current') + ')';
    }
    renderUi();
  }

  /* ═══════════════════════════════════════════════════════════
     2. Extract All Images (DOCX, PPTX, PDF -> Gallery + ZIP)
     ═══════════════════════════════════════════════════════════ */
  async function extractImagesFromOoxmlBytes(bytes) {
    const zip = await window.JSZip.loadAsync(bytes);
    const out = [];
    const mediaRegex = /^(word|ppt)\/media\/(.+)\.(png|jpe?g|gif|webp|bmp|svg)$/i;
    const entries = Object.values(zip.files).filter((f) => !f.dir && mediaRegex.test(f.name));
    for (const entry of entries) {
      const m = mediaRegex.exec(entry.name);
      const ext = m[3].toLowerCase().replace('jpeg', 'jpg');
      const mime = ext === 'svg' ? 'image/svg+xml' : (ext === 'jpg' ? 'image/jpeg' : 'image/' + ext);
      const u8 = new Uint8Array(await entry.async('uint8array'));
      out.push({
        name: entry.name.split('/').pop(),
        bytes: u8,
        mime,
        url: URL.createObjectURL(new Blob([u8], { type: mime }))
      });
    }
    return out;
  }

  async function extractImagesFromPdfBytes(bytes) {
    const { PDFDocument, PDFName, PDFRawStream, PDFNumber, PDFArray, decodePDFRawStream } = window.PDFLib;
    const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    const ctx = pdfDoc.context;
    const nm = (s) => PDFName.of(s);
    const look = (dict, key) => ctx.lookup(dict.get(nm(key)));
    const isNamed = (obj, n) => !!obj && typeof obj.toString === 'function' && obj.toString() === '/' + n;

    const objs = ctx.enumerateIndirectObjects().filter(([, obj]) =>
      obj instanceof PDFRawStream && isNamed(obj.dict.lookup(nm('Subtype')), 'Image'));

    const out = [];
    let idx = 1;
    for (const [, obj] of objs) {
      try {
        const dict = obj.dict;
        const wObj = look(dict, 'Width');
        const hObj = look(dict, 'Height');
        const w = wObj instanceof PDFNumber ? wObj.asNumber() : 0;
        const hgt = hObj instanceof PDFNumber ? hObj.asNumber() : 0;
        if (w < 8 || hgt < 8 || w * hgt > 50e6) continue;

        const filterObj = look(dict, 'Filter');
        let filters = [];
        if (filterObj instanceof PDFArray) filters = filterObj.asArray().map((f) => String(ctx.lookup(f)));
        else if (filterObj) filters = [String(filterObj)];

        if (filters.length === 1 && filters[0] === '/DCTDecode') {
          const u8 = obj.contents.slice();
          out.push({
            name: `image-${idx++}.jpg`,
            bytes: u8,
            mime: 'image/jpeg',
            w, h: hgt,
            url: URL.createObjectURL(new Blob([u8], { type: 'image/jpeg' }))
          });
          continue;
        }

        // Decode FlateDecode or raw pixels into PNG via canvas
        if (typeof decodePDFRawStream === 'function' && (!filters.length || (filters.length === 1 && filters[0] === '/FlateDecode'))) {
          const cs = look(dict, 'ColorSpace');
          let comps = 0;
          if (isNamed(cs, 'DeviceRGB') || isNamed(cs, 'CalRGB')) comps = 3;
          else if (isNamed(cs, 'DeviceGray') || isNamed(cs, 'CalGray')) comps = 1;
          else if (cs instanceof PDFArray) {
            const head = ctx.lookup(cs.get(0));
            if (isNamed(head, 'ICCBased')) {
              const st = ctx.lookup(cs.get(1));
              const nVal = st?.dict ? ctx.lookup(st.dict.get(nm('N'))) : null;
              if (nVal instanceof PDFNumber) comps = nVal.asNumber();
            }
          }
          if (comps !== 1 && comps !== 3) continue;
          const raw = filters.length ? decodePDFRawStream(obj).decode() : obj.contents;
          if (raw.length < w * hgt * comps) continue;
          const cvs = document.createElement('canvas');
          cvs.width = w;
          cvs.height = hgt;
          const c2d = cvs.getContext('2d');
          const imgData = c2d.createImageData(w, hgt);
          const d = imgData.data;
          if (comps === 3) {
            for (let p = 0, s = 0; p < d.length; p += 4, s += 3) {
              d[p] = raw[s]; d[p + 1] = raw[s + 1]; d[p + 2] = raw[s + 2]; d[p + 3] = 255;
            }
          } else {
            for (let p = 0, s = 0; p < d.length; p += 4, s++) {
              d[p] = d[p + 1] = d[p + 2] = raw[s]; d[p + 3] = 255;
            }
          }
          c2d.putImageData(imgData, 0, 0);
          const blob = await new Promise((res) => cvs.toBlob(res, 'image/png'));
          if (blob) {
            const u8 = new Uint8Array(await blob.arrayBuffer());
            out.push({
              name: `image-${idx++}.png`,
              bytes: u8,
              mime: 'image/png',
              w, h: hgt,
              url: URL.createObjectURL(blob)
            });
          }
        }
      } catch (_) { /* skip unsupported mask/CCITT streams */ }
    }
    return out;
  }

  async function openExtractImagesModal(fromActiveEditor = false) {
    let sourceName = '';
    let images = [];

    if (fromActiveEditor && !$id('editor')?.hidden && (curPdf?.bytes || pk)) {
      sourceName = fileName || 'document';
      busy(true, say('جارٍ استخراج الصور من الملف…', 'Extracting images from document…'));
      try {
        if (kind === 'pdf' && curPdf?.bytes) {
          images = await extractImagesFromPdfBytes(curPdf.bytes);
        } else if (pk) {
          const blob = await pk.toBlob('application/zip');
          images = await extractImagesFromOoxmlBytes(new Uint8Array(await blob.arrayBuffer()));
        }
      } catch (err) {
        busy(false);
        toast(say('تعذر استخراج الصور من الملف', 'Failed to extract images'));
        return;
      } finally {
        busy(false);
      }
    } else {
      const f = await pickFile('.docx,.docm,.pptx,.pptm,.pdf,application/pdf');
      if (!f) return;
      assertFileSize(f);
      sourceName = f.name || 'document';
      busy(true, say('جارٍ استخراج الصور من الملف…', 'Extracting images from file…'));
      try {
        const bytes = new Uint8Array(await f.arrayBuffer());
        if (/\.pdf$/i.test(sourceName) || f.type === 'application/pdf') {
          images = await extractImagesFromPdfBytes(bytes);
        } else {
          images = await extractImagesFromOoxmlBytes(bytes);
        }
      } catch (err) {
        busy(false);
        toast(say('تعذر استخراج الصور من الملف', 'Failed to extract images'));
        return;
      } finally {
        busy(false);
      }
    }

    const modal = openToolModal('🖼️ ' + say('استخراج الصور', 'Extract Images') + ` — ${sourceName}`, {
      wide: true,
      onClose: () => images.forEach((im) => URL.revokeObjectURL(im.url))
    });

    if (!images.length) {
      modal.render(
        h('div', { class: 'pt-note' }, say('لم يتم العثور على صور مدمجة في هذا الملف.', 'No embedded images were found in this file.')),
        h('div', { class: 'pt-actions' }, h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')))
      );
      return;
    }

    const downloadOne = (im) => {
      const a = document.createElement('a');
      a.href = im.url;
      a.download = im.name;
      document.body.append(a);
      a.click();
      a.remove();
    };

    const downloadZip = async () => {
      busy(true, say('جارٍ ضغط الصور في ملف ZIP…', 'Creating ZIP archive…'));
      try {
        const z = new window.JSZip();
        images.forEach((im, i) => z.file(im.name || `image-${i + 1}.png`, im.bytes));
        const blob = await z.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = sourceName.replace(/\.[^.]+$/, '') + '-images.zip';
        document.body.append(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1500);
        toast(say('تم تحميل جميع الصور كملف ZIP ✓', 'All images downloaded as ZIP ✓'));
      } finally {
        busy(false);
      }
    };

    const grid = h('div', { class: 'p2-img-grid' },
      ...images.map((im) =>
        h('div', { class: 'p2-img-card' },
          h('div', { class: 'p2-img-thumb' }, h('img', { src: im.url, alt: im.name, loading: 'lazy' })),
          h('div', { class: 'p2-img-meta' },
            h('span', { class: 'p2-img-name', title: im.name }, im.name),
            h('small', null, fmtSize(im.bytes.length))
          ),
          h('button', { type: 'button', class: 'btn', onclick: () => downloadOne(im) }, '⬇ ' + say('تحميل', 'Download'))
        )
      )
    );

    modal.render(
      h('div', { class: 'pt-note' }, say(`تم العثور على ${images.length} صورة داخل الملف.`, `Found ${images.length} embedded image(s).`)),
      grid,
      h('div', { class: 'pt-actions' },
        h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')),
        h('button', { type: 'button', class: 'btn pri', onclick: () => downloadZip() }, '📦 ' + say(`تحميل الكل ZIP (${images.length})`, `Download All ZIP (${images.length})`))
      )
    );
  }

  /* ═══════════════════════════════════════════════════════════
     3. PDF Normalize to A4 / Letter & Crop Margins
     ═══════════════════════════════════════════════════════════ */
  async function normalizePdfPageSizes({ preset, autoOrient, marginPt }) {
    if (!curPdf?.bytes) return;
    busy(true, say('جارٍ توحيد مقاسات صفحات PDF…', 'Normalizing PDF page sizes…'));
    try {
      const { PDFDocument, rgb } = window.PDFLib;
      const srcDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const outDoc = await PDFDocument.create();
      const srcPages = srcDoc.getPages();
      const firstSz = srcPages[0]?.getSize() || { width: 595.28, height: 841.89 };

      let baseW = 595.28, baseH = 841.89; // A4 Portrait
      if (preset === 'a4-landscape') { baseW = 841.89; baseH = 595.28; }
      else if (preset === 'letter') { baseW = 612; baseH = 792; }
      else if (preset === 'match-first') { baseW = firstSz.width; baseH = firstSz.height; }

      const margin = Math.max(0, Math.min(72, +marginPt || 0));
      const newRotations = [];

      for (let i = 0; i < srcPages.length; i++) {
        const sp = srcPages[i];
        const { width: sw, height: sh } = sp.getSize();
        const rot = curPdf.rotations[i] || 0;
        const isLandscape = (rot % 180 ? sh > sw : sw > sh);

        let tw = baseW, th = baseH;
        if (autoOrient) {
          const longSide = Math.max(baseW, baseH);
          const shortSide = Math.min(baseW, baseH);
          tw = isLandscape ? longSide : shortSide;
          th = isLandscape ? shortSide : longSide;
        }

        const embedded = await outDoc.embedPage(sp);
        const page = outDoc.addPage([tw, th]);
        page.drawRectangle({ x: 0, y: 0, width: tw, height: th, color: rgb(1, 1, 1) });

        const availW = Math.max(10, tw - margin * 2);
        const availH = Math.max(10, th - margin * 2);
        const scale = Math.min(availW / sw, availH / sh);
        const drawW = sw * scale;
        const drawH = sh * scale;
        page.drawPage(embedded, {
          x: (tw - drawW) / 2,
          y: (th - drawH) / 2,
          width: drawW,
          height: drawH
        });
        newRotations.push(rot);
      }

      curPdf.bytes = new Uint8Array(await outDoc.save({ useObjectStreams: false, updateMetadata: false }));
      curPdf.rotations = newRotations;
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      await renderPdfEditor();
      toast(say('تم توحيد مقاسات جميع الصفحات بنجاح ✓', 'All pages normalized successfully ✓'));
    } catch (err) {
      console.error('Normalize pages failed:', err);
      toast(say('تعذر توحيد مقاسات الصفحات', 'Failed to normalize page sizes'));
    } finally {
      busy(false);
    }
  }

  async function cropPdfMargins({ topPct, bottomPct, leftPct, rightPct, scope }) {
    if (!curPdf?.bytes) return;
    busy(true, say('جارٍ قص هوامش الصفحات…', 'Cropping PDF margins…'));
    try {
      const { PDFDocument } = window.PDFLib;
      const pdfDoc = await PDFDocument.load(curPdf.bytes, { ignoreEncryption: true, updateMetadata: false });
      const pages = pdfDoc.getPages();
      const targets = scope === 'current' ? [curPdf.pageIdx || 0] : pages.map((_, i) => i);

      for (const idx of targets) {
        const page = pages[idx];
        if (!page) continue;
        const box = page.getCropBox() || page.getMediaBox();
        const rot = ((curPdf.rotations[idx] || 0) % 360 + 360) % 360;
        let t = topPct / 100, b = bottomPct / 100, l = leftPct / 100, r = rightPct / 100;
        // Map visual crop edges to unrotated PDF box edges
        let bxL = l, bxR = r, bxT = t, bxB = b;
        if (rot === 90) { bxL = t; bxR = b; bxT = r; bxB = l; }
        else if (rot === 180) { bxL = r; bxR = l; bxT = b; bxB = t; }
        else if (rot === 270) { bxL = b; bxR = t; bxT = l; bxB = r; }

        const nx = box.x + box.width * bxL;
        const ny = box.y + box.height * bxB;
        const nw = Math.max(40, box.width * (1 - bxL - bxR));
        const nh = Math.max(40, box.height * (1 - bxT - bxB));
        page.setCropBox(nx, ny, nw, nh);
        page.setMediaBox(nx, ny, nw, nh);
      }

      curPdf.bytes = new Uint8Array(await pdfDoc.save({ useObjectStreams: false, updateMetadata: false }));
      if (typeof window.pushPdfUndoSnapshot === 'function') window.pushPdfUndoSnapshot();
      if (typeof window.savePdfDraftNow === 'function') window.savePdfDraftNow();
      await renderPdfEditor();
      toast(say('تم قص هوامش الصفحات بنجاح ✓', 'PDF margins cropped successfully ✓'));
    } catch (err) {
      console.error('Crop PDF failed:', err);
      toast(say('تعذر قص هوامش الصفحات', 'Failed to crop margins'));
    } finally {
      busy(false);
    }
  }

  function openPdfPageSizeCropModal(initialTab = 'normalize') {
    if (kind !== 'pdf' || !curPdf?.bytes) {
      toast(say('افتح ملف PDF أولاً', 'Open a PDF file first'));
      return;
    }
    let activeTab = initialTab;
    const modal = openToolModal('📐 ' + say('توحيد مقاس A4 وقص الهوامش', 'Normalize to A4 & Crop Margins'));

    const renderBody = () => {
      const tabNorm = h('button', {
        type: 'button',
        class: 'btn ' + (activeTab === 'normalize' ? 'pri' : ''),
        onclick: () => { activeTab = 'normalize'; renderBody(); }
      }, '📐 ' + say('توحيد المقاس (A4)', 'Normalize Size'));

      const tabCrop = h('button', {
        type: 'button',
        class: 'btn ' + (activeTab === 'crop' ? 'pri' : ''),
        onclick: () => { activeTab = 'crop'; renderBody(); }
      }, '✂️ ' + say('قص الهوامش', 'Crop Margins'));

      const tabs = h('div', { class: 'p1-tabs' }, tabNorm, tabCrop);

      if (activeTab === 'normalize') {
        const presetSel = h('select', null,
          h('option', { value: 'a4-portrait', selected: true }, 'A4 ' + say('عمودي (210×297 mm)', 'Portrait (210×297 mm)')),
          h('option', { value: 'a4-landscape' }, 'A4 ' + say('أفقي (297×210 mm)', 'Landscape (297×210 mm)')),
          h('option', { value: 'letter' }, 'US Letter (8.5×11 in)'),
          h('option', { value: 'match-first' }, say('مطابقة مقاس الصفحة الأولى', 'Match Page 1 size'))
        );
        const marginIn = h('input', { type: 'number', min: 0, max: 72, value: '12' });
        const autoChk = h('input', { type: 'checkbox', checked: true });

        modal.render(
          tabs,
          h('div', { class: 'pt-note' }, say('يوحّد مقاسات جميع الصفحات بدون فقدان جودة النصوص (Vector).', 'Standardizes all page sizes losslessly while keeping vector text sharp.')),
          h('div', { class: 'p1-grid-2' },
            h('div', { class: 'pt-field' }, h('label', null, say('المقاس المستهدف', 'Target size')), presetSel),
            h('div', { class: 'pt-field' }, h('label', null, say('الهامش (pt)', 'Margin (pt)')), marginIn)
          ),
          h('label', { class: 'pt-check' }, autoChk, h('span', null, say('الحفاظ تلقائياً على اتجاه كل صفحة (عمودي / أفقي)', 'Auto-orient each page (keep landscape pages landscape)'))),
          h('div', { class: 'pt-actions' },
            h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إلغاء', 'Cancel')),
            h('button', {
              type: 'button',
              class: 'btn pri',
              onclick: async () => {
                modal.close();
                await normalizePdfPageSizes({ preset: presetSel.value, autoOrient: autoChk.checked, marginPt: marginIn.value });
              }
            }, say('توحيد الصفحات الآن', 'Normalize Pages'))
          )
        );
      } else {
        const topIn = h('input', { type: 'number', min: 0, max: 35, value: '5' });
        const botIn = h('input', { type: 'number', min: 0, max: 35, value: '5' });
        const leftIn = h('input', { type: 'number', min: 0, max: 35, value: '5' });
        const rightIn = h('input', { type: 'number', min: 0, max: 35, value: '5' });
        const scopeSel = h('select', null,
          h('option', { value: 'all', selected: true }, say('كل الصفحات', 'All pages')),
          h('option', { value: 'current' }, say('الصفحة الحالية فقط', 'Current page only'))
        );

        const setPreset = (v) => { topIn.value = botIn.value = leftIn.value = rightIn.value = String(v); };

        modal.render(
          tabs,
          h('div', { class: 'row', style: { gap: '6px', flexWrap: 'wrap' } },
            h('span', { class: 'pt-note', style: { margin: 0 } }, say('قص سريع:', 'Presets:')),
            h('button', { type: 'button', class: 'btn', onclick: () => setPreset(5) }, '5%'),
            h('button', { type: 'button', class: 'btn', onclick: () => setPreset(10) }, '10%'),
            h('button', { type: 'button', class: 'btn', onclick: () => setPreset(15) }, '15%')
          ),
          h('div', { class: 'p1-grid-2' },
            h('div', { class: 'pt-field' }, h('label', null, say('أعلى (%)', 'Top (%)')), topIn),
            h('div', { class: 'pt-field' }, h('label', null, say('أسفل (%)', 'Bottom (%)')), botIn),
            h('div', { class: 'pt-field' }, h('label', null, say('يسار (%)', 'Left (%)')), leftIn),
            h('div', { class: 'pt-field' }, h('label', null, say('يمين (%)', 'Right (%)')), rightIn)
          ),
          h('div', { class: 'pt-field' }, h('label', null, say('التطبيق على', 'Apply to')), scopeSel),
          h('div', { class: 'pt-actions' },
            h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إلغاء', 'Cancel')),
            h('button', {
              type: 'button',
              class: 'btn pri',
              onclick: async () => {
                const t = Math.max(0, Math.min(35, +topIn.value || 0));
                const b = Math.max(0, Math.min(35, +botIn.value || 0));
                const l = Math.max(0, Math.min(35, +leftIn.value || 0));
                const r = Math.max(0, Math.min(35, +rightIn.value || 0));
                if (t + b >= 75 || l + r >= 75) {
                  toast(say('نسبة القص كبيرة جداً', 'Crop percentage is too large'));
                  return;
                }
                modal.close();
                await cropPdfMargins({ topPct: t, bottomPct: b, leftPct: l, rightPct: r, scope: scopeSel.value });
              }
            }, say('قص الهوامش الآن', 'Crop Margins'))
          )
        );
      }
    };
    renderBody();
  }

  /* ═══════════════════════════════════════════════════════════
     4. Home Quick Tools: Merge PDFs & Standalone Metadata Scrubber
     ═══════════════════════════════════════════════════════════ */
  function pickMultipleFiles(accept) {
    return new Promise((resolve) => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = accept;
      inp.multiple = true;
      inp.hidden = true;
      document.body.append(inp);
      const done = (files) => { inp.remove(); resolve(Array.from(files || [])); };
      inp.onchange = () => done(inp.files);
      inp.oncancel = () => done([]);
      inp.click();
    });
  }

  function downloadBlobFile(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function startHomeMergePdfs() {
    const picked = await pickMultipleFiles('.pdf,application/pdf');
    if (!picked.length) return;
    let list = picked.map((f, i) => ({ f, id: i + '_' + Date.now() }));
    const modal = openToolModal('📑 ' + say('دمج ملفات PDF', 'Merge PDF Files'));

    const renderList = () => {
      const rows = list.map((it, idx) =>
        h('div', { class: 'p1-meta-row', style: { alignItems: 'center' } },
          h('span', { class: 'p1-meta-k' }, `${idx + 1}.`),
          h('span', { class: 'p1-meta-v', style: { flex: '1', maxWidth: 'none' }, title: it.f.name }, it.f.name),
          h('small', { style: { color: 'var(--v2-mut)' } }, fmtSize(it.f.size)),
          h('button', {
            type: 'button', class: 'btn', style: { padding: '2px 6px' }, disabled: idx === 0,
            onclick: () => { const [m] = list.splice(idx, 1); list.splice(idx - 1, 0, m); renderList(); }
          }, '↑'),
          h('button', {
            type: 'button', class: 'btn', style: { padding: '2px 6px' }, disabled: idx === list.length - 1,
            onclick: () => { const [m] = list.splice(idx, 1); list.splice(idx + 1, 0, m); renderList(); }
          }, '↓'),
          h('button', {
            type: 'button', class: 'btn danger', style: { padding: '2px 6px' },
            onclick: () => { list.splice(idx, 1); if (!list.length) modal.close(); else renderList(); }
          }, '✕')
        )
      );

      modal.render(
        h('div', { class: 'pt-note' }, say('رتّب الملفات ثم اضغط دمج لتنزيل ملف PDF واحد مدمج.', 'Arrange files in order, then click Merge to download a single combined PDF.')),
        h('div', { class: 'p1-meta-list' }, ...rows),
        h('div', { class: 'pt-actions' },
          h('button', {
            type: 'button', class: 'btn',
            onclick: async () => {
              const more = await pickMultipleFiles('.pdf,application/pdf');
              if (more.length) {
                more.forEach((f, i) => list.push({ f, id: 'm_' + i + '_' + Date.now() }));
                renderList();
              }
            }
          }, '➕ ' + say('إضافة ملفات', 'Add Files')),
          h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إلغاء', 'Cancel')),
          h('button', {
            type: 'button', class: 'btn pri', disabled: list.length < 2,
            onclick: async () => {
              modal.close();
              busy(true, say('جارٍ دمج ملفات PDF…', 'Merging PDF files…'));
              try {
                const { PDFDocument } = window.PDFLib;
                const merged = await PDFDocument.create();
                for (const it of list) {
                  assertFileSize(it.f);
                  let bytes = new Uint8Array(await it.f.arrayBuffer());
                  if (window.PdfTools?.ensureReadable) bytes = await window.PdfTools.ensureReadable(bytes, it.f.name);
                  const src = await PDFDocument.load(bytes, { ignoreEncryption: true });
                  const copied = await merged.copyPages(src, src.getPageIndices());
                  copied.forEach((p) => merged.addPage(p));
                }
                const out = await merged.save({ useObjectStreams: false });
                const base = (list[0].f.name || 'merged').replace(/\.pdf$/i, '');
                downloadBlobFile(new Blob([out], { type: 'application/pdf' }), `${base}-merged.pdf`);
                toast(say(`تم دمج ${list.length} ملفات PDF بنجاح ✓`, `Merged ${list.length} PDF files successfully ✓`));
              } catch (err) {
                console.error('Home merge failed:', err);
                toast(err instanceof UserError ? err.message : say('تعذر دمج ملفات PDF', 'Failed to merge PDFs'));
              } finally {
                busy(false);
              }
            }
          }, '📑 ' + say(`دمج (${list.length})`, `Merge (${list.length})`))
        )
      );
    };
    renderList();
  }

  async function startHomeCleanMetadata() {
    const f = await pickFile('.pdf,.docx,.docm,.pptx,.pptm,application/pdf');
    if (!f) return;
    assertFileSize(f);
    const name = f.name || 'document';
    const isPdf = /\.pdf$/i.test(name) || f.type === 'application/pdf';
    const modal = openToolModal('🛡️ ' + say('تنظيف البيانات المخفية', 'Clean Hidden Metadata') + ` — ${name}`);
    modal.render(h('div', { class: 'pt-note' }, say('جارٍ فحص البيانات الوصفية…', 'Scanning metadata…')));

    try {
      let bytes = new Uint8Array(await f.arrayBuffer());
      const items = [];

      if (isPdf) {
        if (window.PdfTools?.ensureReadable) bytes = await window.PdfTools.ensureReadable(bytes, name);
        const { PDFDocument, PDFName } = window.PDFLib;
        const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
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
        check('المنتج', 'Producer', pdfDoc.getProducer());
        try { check('تاريخ الإنشاء', 'Creation Date', pdfDoc.getCreationDate()); } catch (_) {}
        try { check('تاريخ التعديل', 'Modification Date', pdfDoc.getModificationDate()); } catch (_) {}
        if (pdfDoc.catalog?.has(PDFName.of('Metadata'))) {
          items.push({ label: say('بيانات XMP المخفية', 'Embedded XMP Metadata'), value: say('موجودة', 'Present') });
        }

        const cleanAndSave = async () => {
          pdfDoc.setTitle(''); pdfDoc.setAuthor(''); pdfDoc.setSubject('');
          pdfDoc.setKeywords([]); pdfDoc.setCreator(''); pdfDoc.setProducer('');
          if (pdfDoc.catalog?.has(PDFName.of('Metadata'))) pdfDoc.catalog.delete(PDFName.of('Metadata'));
          if (pdfDoc.context?.trailerInfo?.Info) {
            const infoDict = pdfDoc.context.lookup(pdfDoc.context.trailerInfo.Info);
            if (infoDict?.keys) for (const k of Array.from(infoDict.keys())) infoDict.delete(k);
            delete pdfDoc.context.trailerInfo.Info;
          }
          const out = await pdfDoc.save({ useObjectStreams: false, updateMetadata: false });
          downloadBlobFile(new Blob([out], { type: 'application/pdf' }), name.replace(/\.pdf$/i, '') + '-clean.pdf');
          modal.close();
          toast(say('تم تنظيف الملف وتنزيله بنجاح ✓', 'Cleaned file downloaded ✓'));
        };

        renderHomeMetaModal(modal, items, cleanAndSave);
      } else {
        const zip = await window.JSZip.loadAsync(bytes);
        const parser = new DOMParser();
        const ser = new XMLSerializer();
        const scanXml = async (path, map) => {
          const entry = zip.file(path);
          if (!entry) return null;
          const doc = parser.parseFromString(await entry.async('string'), 'application/xml');
          for (const [tag, ar, en] of map) {
            const el = Array.from(doc.getElementsByTagName('*')).find((n) => n.localName === tag && n.textContent?.trim());
            if (el && el.textContent.trim() !== '0') items.push({ label: say(ar, en), value: el.textContent.trim() });
          }
          return doc;
        };
        const coreDoc = await scanXml('docProps/core.xml', [
          ['creator', 'المؤلف', 'Author'], ['lastModifiedBy', 'آخر تعديل بواسطة', 'Last Modified By'],
          ['title', 'العنوان', 'Title'], ['subject', 'الموضوع', 'Subject'], ['keywords', 'الكلمات المفتاحية', 'Keywords'],
          ['created', 'تاريخ الإنشاء', 'Created'], ['modified', 'تاريخ التعديل', 'Modified']
        ]);
        const appDoc = await scanXml('docProps/app.xml', [
          ['Company', 'الشركة', 'Company'], ['Manager', 'المدير', 'Manager'],
          ['Application', 'التطبيق', 'Application'], ['Template', 'القالب', 'Template']
        ]);
        const customEntry = zip.file('docProps/custom.xml');
        let customDoc = null;
        if (customEntry) {
          customDoc = parser.parseFromString(await customEntry.async('string'), 'application/xml');
          const cnt = customDoc?.documentElement?.children?.length || 0;
          if (cnt) items.push({ label: say('خصائص مخصصة', 'Custom Properties'), value: String(cnt) });
        }

        const cleanAndSave = async () => {
          const wipe = (doc, tags) => {
            if (!doc?.documentElement) return;
            const s = new Set(tags);
            Array.from(doc.getElementsByTagName('*')).forEach((el) => { if (s.has(el.localName)) el.textContent = ''; });
          };
          if (coreDoc) {
            wipe(coreDoc, ['creator', 'lastModifiedBy', 'title', 'subject', 'description', 'keywords', 'category', 'revision', 'created', 'modified', 'lastPrinted']);
            zip.file('docProps/core.xml', ser.serializeToString(coreDoc));
          }
          if (appDoc) {
            wipe(appDoc, ['Company', 'Manager', 'Application', 'AppVersion', 'Template', 'TotalTime', 'HyperlinkBase']);
            zip.file('docProps/app.xml', ser.serializeToString(appDoc));
          }
          if (customDoc?.documentElement) {
            while (customDoc.documentElement.firstChild) customDoc.documentElement.removeChild(customDoc.documentElement.firstChild);
            zip.file('docProps/custom.xml', ser.serializeToString(customDoc));
          }
          const blob = await zip.generateAsync({ type: 'blob' });
          const cleanName = name.replace(/(\.(docx|docm|pptx|pptm))$/i, '-clean$1');
          downloadBlobFile(blob, cleanName);
          modal.close();
          toast(say('تم تنظيف الملف وتنزيله بنجاح ✓', 'Cleaned file downloaded ✓'));
        };

        renderHomeMetaModal(modal, items, cleanAndSave);
      }
    } catch (err) {
      console.error('Home clean metadata failed:', err);
      modal.render(
        h('div', { class: 'pt-err' }, say('تعذر قراءة الملف.', 'Could not read the file.')),
        h('div', { class: 'pt-actions' }, h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')))
      );
    }
  }

  function renderHomeMetaModal(modal, items, onClean) {
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

    modal.render(
      list,
      h('div', { class: 'pt-actions' },
        h('button', { type: 'button', class: 'btn', onclick: () => modal.close() }, say('إغلاق', 'Close')),
        h('button', { type: 'button', class: 'btn pri', disabled: !items.length, onclick: () => onClean() }, '🛡️ ' + say('تنظيف وتنزيل النسخة النظيفة', 'Clean & Download'))
      )
    );
  }

  /* ═══════════════════════════════════════════════════════════
     5. Wire Home Buttons & PDF Toolbar Button
     ═══════════════════════════════════════════════════════════ */
  $id('bCompareDocs')?.addEventListener('click', () => openCompareDocsModal(false));
  $id('bExtractImages')?.addEventListener('click', () => openExtractImagesModal(false));
  $id('bMergePdfsHome')?.addEventListener('click', () => startHomeMergePdfs());
  $id('bCleanMetaHome')?.addEventListener('click', () => startHomeCleanMetadata());

  if (typeof renderPdfEditor === 'function') {
    const origRenderPdf = renderPdfEditor;
    renderPdfEditor = async function (...args) {
      const res = await origRenderPdf.apply(this, args);
      try {
        if (kind !== 'pdf' || !curPdf?.bytes) return res;
        const toolbar = document.querySelector('.pdf-toolbar');
        if (toolbar && !toolbar._phase2Added) {
          toolbar._phase2Added = true;
          const bNormCrop = btn(say('📐 مقاس A4 / قص', '📐 A4 / Crop'), () => openPdfPageSizeCropModal('normalize'));
          const bExtract = btn(say('🖼️ استخراج الصور', '🖼️ Extract Images'), () => openExtractImagesModal(true));
          const btns = Array.from(toolbar.querySelectorAll('button'));
          const rotRightBtn = btns.find((b) => /↻/.test(b.textContent || ''));
          const exportPngBtn = btns.find((b) => /Export PNG|تصدير PNG/i.test(b.textContent || ''));
          const spacer = Array.from(toolbar.children).find((c) => c.tagName === 'SPAN' && c.style.flex === '1');

          if (rotRightBtn && rotRightBtn.parentNode === toolbar) {
            toolbar.insertBefore(bNormCrop, rotRightBtn.nextSibling);
          } else if (spacer) {
            toolbar.insertBefore(bNormCrop, spacer);
          } else {
            toolbar.append(bNormCrop);
          }

          if (exportPngBtn && exportPngBtn.parentNode === toolbar) {
            toolbar.insertBefore(bExtract, exportPngBtn.nextSibling);
          } else if (spacer) {
            toolbar.insertBefore(bExtract, spacer);
          } else {
            toolbar.append(bExtract);
          }
        }
      } catch (err) {
        console.warn('Phase 2 PDF toolbar integration failed:', err);
      }
      return res;
    };
  }

  window.Phase2Tools = {
    openCompareDocsModal,
    openExtractImagesModal,
    openPdfPageSizeCropModal
  };
})();
