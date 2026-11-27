'use strict';
/* ═════════════════════════════════════════════════════════════
   PDF Tools — ضغط، حماية/فك حماية بكلمة مرور، وتعبئة النماذج
   وحدة مستقلة تُحمَّل بعد js/pdf-editor.js:
   - Compress   : pdf-lib + canvas (إعادة ضغط الصور بمستويات) — بدون مكتبات إضافية
   - Protect    : qpdf.wasm (تشفير AES-256) يُحمَّل عند أول استخدام فقط
   - Unlock     : qpdf.wasm (فك التشفير) ويُستخدم أيضاً عند فتح ملف محمي في المحرر
   - Fill forms : pdf-lib AcroForm
   ═════════════════════════════════════════════════════════════ */
(() => {
  if (!window.PDFLib || !window.PDFLib.PDFDocument) return;

  const isAr = () => document.documentElement.lang === 'ar';
  const say = (ar, en) => (isAr() ? ar : en);
  const $id = (id) => document.getElementById(id);
  const setBusy = (on, msg) => { if (typeof busy === 'function') busy(on, msg); };
  const setBusyMsg = (msg) => { const el = $id('busyMsg'); if (el) el.textContent = msg; };
  const notify = (msg) => { if (typeof toast === 'function') toast(msg); else window.alert(msg); };
  const nextTick = () => new Promise((resolve) => setTimeout(resolve, 0));

  class WrongPasswordError extends Error {}
  class SilentCancel extends Error {}

  const baseName = (name) => String(name || 'document').replace(/\.pdf$/i, '');
  const fmtSize = (n) => {
    if (n >= 1048576) return (n / 1048576).toFixed(2) + ' MB';
    return Math.max(1, Math.round(n / 1024)) + ' KB';
  };

  function saveBytes(bytes, name) {
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function reportError(err) {
    if (err instanceof SilentCancel) return;
    console.error('PDF tool failed:', err);
    const msg = err instanceof UserError ? err.message : (say('تعذر تنفيذ العملية: ', 'Operation failed: ') + ((err && err.message) || err));
    notify(msg);
  }

  const loadDoc = (bytes) => window.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });

  function pickPdf() {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf,application/pdf';
      input.hidden = true;
      document.body.append(input);
      const done = (file) => { input.remove(); resolve(file); };
      input.onchange = () => done((input.files && input.files[0]) || null);
      input.oncancel = () => done(null);
      input.click();
    });
  }

  /* ───────────── نافذة منبثقة عامة ───────────── */
  function openModal(title, opts = {}) {
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
    const box = h('div', { class: 'md-modal pt-modal' },
      h('div', { class: 'row', style: { justifyContent: 'space-between' } },
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

  function passwordInput(placeholder) {
    const input = h('input', { type: 'password', autocomplete: 'off', spellcheck: 'false', placeholder: placeholder || '' });
    const toggle = h('button', {
      type: 'button',
      class: 'btn',
      title: say('إظهار / إخفاء', 'Show / Hide'),
      onclick: () => { input.type = input.type === 'password' ? 'text' : 'password'; }
    }, '👁');
    return { input, row: h('div', { class: 'pt-pass' }, input, toggle) };
  }

  function actionsRow(...buttons) {
    return h('div', { class: 'pt-actions' }, ...buttons);
  }

  function textButton(label, handler, cls) {
    return h('button', { type: 'button', class: 'btn ' + (cls || ''), onclick: handler }, label);
  }

  /* ───────────── محرك qpdf (WASM) — يُحمَّل عند الحاجة فقط ───────────── */
  let qpdfPromise = null;
  let qpdfCounter = 0;

  function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error(say('تعذر تحميل محرك التشفير. تأكد من اتصالك بالإنترنت لأول استخدام.', 'Could not load the encryption engine. Check your connection on first use.')));
      document.head.append(s);
    });
  }

  function getQpdf() {
    if (!qpdfPromise) {
      qpdfPromise = (async () => {
        if (typeof window.Module !== 'function') await loadScriptOnce('qpdf.js');
        if (typeof window.Module !== 'function') throw new Error('qpdf unavailable');
        const logs = [];
        const instance = await window.Module({
          locateFile: (p) => (/\.wasm$/.test(p) ? 'qpdf.wasm' : p),
          noInitialRun: true,
          print: (t) => logs.push(String(t)),
          printErr: (t) => logs.push(String(t))
        });
        return { q: instance, logs };
      })().catch((err) => { qpdfPromise = null; throw err; });
    }
    return qpdfPromise;
  }

  async function runQpdf(bytes, buildArgs, opts = {}) {
    const { q, logs } = await getQpdf();
    const id = '/job' + (++qpdfCounter);
    const inPath = id + '-in.pdf';
    const outPath = id + '-out.pdf';
    logs.length = 0;
    q.FS.writeFile(inPath, bytes);
    let out = null;
    try {
      try { q.callMain(buildArgs(inPath, outPath)); } catch (e) { /* emscripten exit status */ }
      try { out = q.FS.readFile(outPath).slice(); } catch (e) { out = null; }
    } finally {
      try { q.FS.unlink(inPath); } catch (e) { /* ignore */ }
      try { q.FS.unlink(outPath); } catch (e) { /* ignore */ }
    }
    const log = logs.join('\n');
    if (!out || !out.length) {
      if (/invalid password/i.test(log) || (opts.decrypt && !/(unable to find|damaged|not a PDF|can't find PDF header)/i.test(log))) throw new WrongPasswordError('invalid password');
      throw new Error(log.split('\n').filter(Boolean).slice(-2).join(' ') || 'qpdf failed');
    }
    return out;
  }

  const unlockBytes = (bytes, password) => runQpdf(bytes, (i, o) => ['--password=' + password, '--decrypt', i, o], { decrypt: true });
  const protectBytes = (bytes, password) => runQpdf(bytes, (i, o) => ['--encrypt', password, password, '256', '--', i, o]);

  /* ───────────── اكتشاف الملفات المحمية ───────────── */
  const latin1 = new TextDecoder('latin1');
  const TAIL_SCAN = 524288;
  const HEAD_SCAN = 131072;

  function bytesMayBeEncrypted(u8) {
    if (latin1.decode(u8.subarray(Math.max(0, u8.length - TAIL_SCAN))).includes('/Encrypt')) return true;
    return latin1.decode(u8.subarray(0, Math.min(u8.length, HEAD_SCAN))).includes('/Encrypt');
  }

  async function fileMayBeEncrypted(file) {
    const tail = latin1.decode(await file.slice(Math.max(0, file.size - TAIL_SCAN)).arrayBuffer());
    if (tail.includes('/Encrypt')) return true;
    return latin1.decode(await file.slice(0, HEAD_SCAN).arrayBuffer()).includes('/Encrypt');
  }

  async function pdfNeedsDecrypt(u8) {
    if (!bytesMayBeEncrypted(u8)) return false;
    try {
      const doc = await loadDoc(u8);
      return !!doc.isEncrypted;
    } catch (e) {
      return true;
    }
  }

  function askPassword({ title, message, okLabel, validate }) {
    return new Promise((resolve) => {
      const busyEl = $id('busy');
      const wasBusy = !!(busyEl && !busyEl.hidden);
      if (wasBusy) busyEl.hidden = true;
      let finished = false;
      const finish = (value) => {
        if (finished) return;
        finished = true;
        if (wasBusy && busyEl) busyEl.hidden = false;
        resolve(value);
      };
      const modal = openModal(title, { onClose: () => finish(null) });
      const pw = passwordInput(say('كلمة المرور', 'Password'));
      const err = h('div', { class: 'pt-err' });
      const okBtn = textButton(okLabel, () => submit(), 'pri');
      const submit = async () => {
        err.textContent = '';
        okBtn.disabled = true;
        const oldLabel = okBtn.textContent;
        okBtn.textContent = say('جارٍ التحقق…', 'Checking…');
        try {
          const result = await validate(pw.input.value);
          finish({ password: pw.input.value, result });
          modal.close();
        } catch (e) {
          okBtn.disabled = false;
          okBtn.textContent = oldLabel;
          if (e instanceof WrongPasswordError) {
            err.textContent = say('كلمة المرور غير صحيحة.', 'Incorrect password.');
            pw.input.select();
          } else {
            err.textContent = (e && e.message) || String(e);
          }
        }
      };
      pw.input.onkeydown = (e) => { if (e.key === 'Enter') submit(); };
      modal.render(
        h('div', { class: 'pt-note' }, message),
        h('div', { class: 'pt-field' }, pw.row),
        err,
        actionsRow(textButton(say('إلغاء', 'Cancel'), () => modal.close()), okBtn)
      );
      pw.input.focus();
    });
  }

  /** يرجّع بايتات قابلة للتعديل: إن كان الملف مشفراً يُفك بكلمة المرور (يُطلب من المستخدم عند الحاجة). */
  async function ensureReadable(bytes, name) {
    if (!(await pdfNeedsDecrypt(bytes))) return bytes;
    try {
      return await unlockBytes(bytes, '');
    } catch (e) {
      if (!(e instanceof WrongPasswordError)) throw e;
    }
    const res = await askPassword({
      title: say('🔒 ملف محمي بكلمة مرور', '🔒 Password-protected PDF'),
      message: say('هذا الملف محمي. أدخل كلمة المرور لفتحه: ', 'This file is protected. Enter its password to open it: ') + (name || ''),
      okLabel: say('فتح', 'Unlock'),
      validate: (pw) => unlockBytes(bytes, pw)
    });
    if (!res) throw new SilentCancel('cancelled');
    return res.result;
  }

  /* ═════════════ 1) ضغط PDF ═════════════ */
  const LEVELS = {
    low: { q: 0.85, maxSide: 3000 },
    medium: { q: 0.65, maxSide: 2000 },
    high: { q: 0.45, maxSide: 1300 },
    extreme: { raster: true, scale: 1.4, q: 0.5 }
  };

  const canvasToJpeg = (canvas, quality) => new Promise((resolve) => canvas.toBlob((b) => resolve(b), 'image/jpeg', quality));

  async function decodeJpegBitmap(bytes) {
    const blob = new Blob([bytes], { type: 'image/jpeg' });
    try {
      return await createImageBitmap(blob, { imageOrientation: 'none', colorSpaceConversion: 'none' });
    } catch (e) {
      return await createImageBitmap(blob);
    }
  }

  async function recompressImages(pdfDoc, level, onProgress) {
    const { PDFName, PDFRawStream, PDFNumber, PDFArray, PDFBool, decodePDFRawStream } = window.PDFLib;
    const ctx = pdfDoc.context;
    const nm = (s) => PDFName.of(s);
    const look = (dict, key) => ctx.lookup(dict.get(nm(key)));
    const isNamed = (obj, n) => !!obj && typeof obj.toString === 'function' && obj.toString() === '/' + n;

    const compsOf = (cs) => {
      if (!cs) return null;
      if (cs instanceof PDFName) {
        const n = cs.toString();
        if (n === '/DeviceRGB' || n === '/CalRGB') return 3;
        if (n === '/DeviceGray' || n === '/CalGray') return 1;
        return null;
      }
      if (cs instanceof PDFArray) {
        const arr = cs.asArray();
        const head = ctx.lookup(arr[0]);
        if (isNamed(head, 'CalRGB')) return 3;
        if (isNamed(head, 'CalGray')) return 1;
        if (isNamed(head, 'ICCBased')) {
          const st = ctx.lookup(arr[1]);
          const n = st && st.dict ? ctx.lookup(st.dict.get(nm('N'))) : null;
          const v = n instanceof PDFNumber ? n.asNumber() : null;
          return v === 3 || v === 1 ? v : null;
        }
      }
      return null;
    };

    const imgs = ctx.enumerateIndirectObjects().filter(([, obj]) =>
      obj instanceof PDFRawStream && isNamed(obj.dict.lookup(nm('Subtype')), 'Image') && obj.contents.length > 20000);

    let replaced = 0;
    for (let i = 0; i < imgs.length; i++) {
      const [ref, obj] = imgs[i];
      onProgress(i + 1, imgs.length);
      if (i % 3 === 2) await nextTick();
      try {
        const dict = obj.dict;
        const maskFlag = look(dict, 'ImageMask');
        if (maskFlag instanceof PDFBool && maskFlag.asBoolean()) continue;
        if (dict.has(nm('Decode')) || dict.has(nm('Mask'))) continue;
        const sm = look(dict, 'SMask');
        if (sm && sm.dict && sm.dict.has(nm('Matte'))) continue;

        const wObj = look(dict, 'Width');
        const hObj = look(dict, 'Height');
        if (!(wObj instanceof PDFNumber) || !(hObj instanceof PDFNumber)) continue;
        const w = wObj.asNumber();
        const hgt = hObj.asNumber();
        if (w * hgt > 60e6 || w > 16000 || hgt > 16000) continue;

        const filterObj = look(dict, 'Filter');
        let filters = [];
        if (filterObj instanceof PDFArray) filters = filterObj.asArray().map((f) => String(ctx.lookup(f)));
        else if (filterObj) filters = [String(filterObj)];
        const csRaw = dict.get(nm('ColorSpace'));
        const comps = compsOf(ctx.lookup(csRaw));
        if (!comps) continue;

        const scale = Math.min(1, level.maxSide / Math.max(w, hgt));
        const nw = Math.max(1, Math.round(w * scale));
        const nh = Math.max(1, Math.round(hgt * scale));
        const canvas = document.createElement('canvas');
        canvas.width = nw;
        canvas.height = nh;
        const c2d = canvas.getContext('2d');
        c2d.fillStyle = '#fff';
        c2d.fillRect(0, 0, nw, nh);

        if (filters.length === 1 && filters[0] === '/DCTDecode') {
          const bmp = await decodeJpegBitmap(obj.contents);
          c2d.imageSmoothingQuality = 'high';
          c2d.drawImage(bmp, 0, 0, nw, nh);
          if (bmp.close) bmp.close();
        } else if (filters.length === 1 && filters[0] === '/FlateDecode') {
          const bpc = look(dict, 'BitsPerComponent');
          if (!(bpc instanceof PDFNumber) || bpc.asNumber() !== 8) continue;
          const raw = decodePDFRawStream(obj).decode();
          if (raw.length < w * hgt * comps) continue;
          const rgba = new Uint8ClampedArray(w * hgt * 4);
          if (comps === 3) {
            for (let p = 0, s = 0, d = 0; p < w * hgt; p++, s += 3, d += 4) {
              rgba[d] = raw[s]; rgba[d + 1] = raw[s + 1]; rgba[d + 2] = raw[s + 2]; rgba[d + 3] = 255;
            }
          } else {
            for (let p = 0, d = 0; p < w * hgt; p++, d += 4) {
              const v = raw[p];
              rgba[d] = v; rgba[d + 1] = v; rgba[d + 2] = v; rgba[d + 3] = 255;
            }
          }
          if (scale === 1) {
            c2d.putImageData(new ImageData(rgba, w, hgt), 0, 0);
          } else {
            const full = document.createElement('canvas');
            full.width = w;
            full.height = hgt;
            full.getContext('2d').putImageData(new ImageData(rgba, w, hgt), 0, 0);
            c2d.imageSmoothingQuality = 'high';
            c2d.drawImage(full, 0, 0, nw, nh);
            full.width = 0;
          }
        } else {
          continue;
        }

        const blob = await canvasToJpeg(canvas, level.q);
        canvas.width = 0;
        if (!blob) continue;
        const jpg = new Uint8Array(await blob.arrayBuffer());
        if (jpg.length >= obj.contents.length * 0.95) continue;

        const newDict = ctx.obj({
          Type: 'XObject', Subtype: 'Image', Width: nw, Height: nh,
          BitsPerComponent: 8, Filter: 'DCTDecode', Length: jpg.length
        });
        newDict.set(nm('ColorSpace'), comps === 3 ? csRaw : nm('DeviceRGB'));
        const smRaw = dict.get(nm('SMask'));
        if (smRaw) newDict.set(nm('SMask'), smRaw);
        ctx.assign(ref, PDFRawStream.of(newDict, jpg));
        replaced++;
      } catch (e) {
        console.warn('Image skipped during compression:', e);
      }
    }
    return { total: imgs.length, replaced };
  }

  async function rasterizeCompress(bytes, level, onProgress) {
    const lib = typeof configurePdfJsWorker === 'function' ? configurePdfJsWorker() : window.pdfjsLib;
    if (!lib) throw new Error('pdf.js missing');
    const { PDFDocument } = window.PDFLib;
    const src = await loadDoc(bytes);
    const out = await PDFDocument.create();
    try {
      const t = src.getTitle(); if (t) out.setTitle(t);
      const a = src.getAuthor(); if (a) out.setAuthor(a);
      const s = src.getSubject(); if (s) out.setSubject(s);
      const k = src.getKeywords(); if (k) out.setKeywords(k.split(/\s+/).filter(Boolean));
      const cd = src.getCreationDate(); if (cd) out.setCreationDate(cd);
    } catch (e) { /* metadata is best-effort */ }
    const doc = await lib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
    try {
      for (let i = 1; i <= doc.numPages; i++) {
        onProgress(i, doc.numPages);
        const page = await doc.getPage(i);
        const base = page.getViewport({ scale: 1 });
        const sc = Math.min(level.scale, 4096 / Math.max(base.width, base.height));
        const vp = page.getViewport({ scale: sc });
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(vp.width));
        canvas.height = Math.max(1, Math.round(vp.height));
        const c2d = canvas.getContext('2d');
        c2d.fillStyle = '#fff';
        c2d.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: c2d, viewport: vp }).promise;
        const blob = await canvasToJpeg(canvas, level.q);
        canvas.width = 0;
        if (!blob) throw new Error('JPEG encode failed');
        const img = await out.embedJpg(new Uint8Array(await blob.arrayBuffer()));
        const p = out.addPage([base.width, base.height]);
        p.drawImage(img, { x: 0, y: 0, width: base.width, height: base.height });
        page.cleanup();
        await nextTick();
      }
    } finally {
      doc.destroy();
    }
    return out.save({ useObjectStreams: true });
  }

  /** ضغط بايتات PDF (جاهزة وغير مشفرة). يرجّع { bytes, before, after, images, replaced }. */
  async function compressBytes(bytes, levelKey, onProgress) {
    const level = LEVELS[levelKey] || LEVELS.medium;
    const progress = onProgress || (() => {});
    const before = bytes.length;
    if (level.raster) {
      const out = await rasterizeCompress(bytes, level, progress);
      return { bytes: out, before, after: out.length, images: 0, replaced: 0, raster: true };
    }
    const doc = await loadDoc(bytes);
    const stats = await recompressImages(doc, level, progress);
    const out = await doc.save({ useObjectStreams: true });
    return { bytes: out, before, after: out.length, images: stats.total, replaced: stats.replaced, raster: false };
  }

  async function startCompress(rawBytes, name) {
    const bytes = await ensureReadable(rawBytes, name);
    try { await loadDoc(bytes); } catch (e) { throw new UserError(say('ملف PDF غير صالح أو تالف.', 'The PDF file is invalid or damaged.')); }
    let chosen = 'medium';

    const showOptions = () => {
      const modal = openModal(say('🗜 ضغط PDF', '🗜 Compress PDF'));
      const opts = [
        ['low', say('ضغط خفيف', 'Light'), say('جودة عالية جداً، تصغير بسيط للحجم', 'Very high quality, small size reduction')],
        ['medium', say('ضغط متوسط (موصى به)', 'Medium (recommended)'), say('توازن جيد بين الحجم والوضوح', 'Good balance of size and clarity')],
        ['high', say('ضغط قوي', 'Strong'), say('أصغر حجم مع بقاء النص واضحاً وقابلاً للتحديد', 'Smallest size; text stays sharp and selectable')],
        ['extreme', say('أقصى ضغط (تحويل الصفحات إلى صور)', 'Maximum (pages become images)'), say('أصغر حجم ممكن، لكن النص لن يكون قابلاً للتحديد أو البحث', 'Smallest possible, but text is no longer selectable or searchable')]
      ];
      const list = h('div', { class: 'pt-scroll', style: { maxHeight: 'none' } });
      opts.forEach(([key, title, desc]) => {
        const radio = h('input', { type: 'radio', name: 'ptLevel', value: key });
        radio.checked = key === chosen;
        const row = h('label', { class: 'pt-opt' + (key === chosen ? ' sel' : '') }, radio, h('span', null, h('b', null, title), h('small', null, desc)));
        radio.onchange = () => {
          chosen = key;
          list.querySelectorAll('.pt-opt').forEach((el) => el.classList.remove('sel'));
          row.classList.add('sel');
        };
        list.append(row);
      });
      modal.render(
        h('div', { class: 'pt-note' }, `${name || ''} — ${fmtSize(bytes.length)}`),
        list,
        h('div', { class: 'pt-note' }, say('يتم الضغط على جهازك فقط. بيانات الملف (العنوان، المؤلف…) تبقى كما هي.', 'Compression runs on your device only. File metadata (title, author…) is preserved.')),
        actionsRow(
          textButton(say('إلغاء', 'Cancel'), () => modal.close()),
          textButton(say('ضغط', 'Compress'), () => { modal.close(); run(); }, 'pri')
        )
      );
    };

    const run = async () => {
      setBusy(true, say('جارٍ الضغط…', 'Compressing…'));
      let res;
      try {
        res = await compressBytes(bytes, chosen, (i, n) => setBusyMsg(
          chosen === 'extreme' ? say(`جارٍ معالجة الصفحة ${i}/${n}…`, `Processing page ${i}/${n}…`) : say(`جارٍ ضغط الصور ${i}/${n}…`, `Compressing images ${i}/${n}…`)
        ));
      } catch (err) {
        setBusy(false);
        reportError(err);
        showOptions();
        return;
      }
      setBusy(false);
      showResult(res);
    };

    const showResult = (res) => {
      const modal = openModal(say('🗜 نتيجة الضغط', '🗜 Compression result'));
      const gained = res.after < res.before;
      const pct = Math.round((1 - res.after / res.before) * 100);
      const nodes = [];
      if (gained) {
        nodes.push(h('div', { class: 'pt-result' },
          h('div', null, `${fmtSize(res.before)}  →  `, h('b', null, fmtSize(res.after))),
          h('div', { class: 'pt-note' }, say(`تم تصغير الحجم بنسبة ${pct}%`, `Size reduced by ${pct}%`))));
      } else {
        nodes.push(h('div', { class: 'pt-result' },
          h('div', null, h('b', null, say('الملف محسّن مسبقاً', 'The file is already optimized'))),
          h('div', { class: 'pt-note' }, say('لم يمكن تصغيره أكثر بهذا المستوى. جرّب مستوى أقوى.', 'It could not be reduced at this level. Try a stronger level.'))));
      }
      if (!res.raster && res.images === 0) {
        nodes.push(h('div', { class: 'pt-note' }, say('لا يحتوي الملف على صور كبيرة تستحق الضغط؛ الحجم الأساسي من النصوص والخطوط.', 'This file has no large images to compress; its size comes from text and fonts.')));
      }
      nodes.push(actionsRow(
        textButton(say('تجربة مستوى آخر', 'Try another level'), () => { modal.close(); showOptions(); }),
        gained ? textButton(say('⬇ تنزيل', '⬇ Download'), () => {
          saveBytes(res.bytes, baseName(name) + '-compressed.pdf');
          notify(say('تم تنزيل الملف المضغوط', 'Compressed file downloaded'));
        }, 'pri') : null,
        textButton(say('إغلاق', 'Close'), () => modal.close())
      ));
      modal.render(...nodes);
    };

    showOptions();
  }

  /* ═════════════ 2) حماية PDF بكلمة مرور ═════════════ */
  async function startProtect(rawBytes, name) {
    const bytes = await ensureReadable(rawBytes, name);
    const modal = openModal(say('🔒 حماية PDF بكلمة مرور', '🔒 Protect PDF with a password'));
    const p1 = passwordInput(say('كلمة المرور', 'Password'));
    const p2 = passwordInput(say('تأكيد كلمة المرور', 'Confirm password'));
    const err = h('div', { class: 'pt-err' });
    const okBtn = textButton(say('🔒 حماية وتنزيل', '🔒 Protect & download'), () => submit(), 'pri');
    const submit = async () => {
      err.textContent = '';
      if (p1.input.value.length < 4) { err.textContent = say('استخدم 4 أحرف على الأقل.', 'Use at least 4 characters.'); return; }
      if (p1.input.value !== p2.input.value) { err.textContent = say('كلمتا المرور غير متطابقتين.', 'Passwords do not match.'); return; }
      okBtn.disabled = true;
      const oldLabel = okBtn.textContent;
      okBtn.textContent = say('جارٍ التشفير…', 'Encrypting…');
      try {
        const out = await protectBytes(bytes, p1.input.value);
        saveBytes(out, baseName(name) + '-protected.pdf');
        modal.close();
        notify(say('تم تنزيل الملف المحمي', 'Protected file downloaded'));
      } catch (e) {
        okBtn.disabled = false;
        okBtn.textContent = oldLabel;
        err.textContent = (e && e.message) || String(e);
      }
    };
    p2.input.onkeydown = (e) => { if (e.key === 'Enter') submit(); };
    modal.render(
      h('div', { class: 'pt-note' }, `${name || ''} — ${fmtSize(bytes.length)}`),
      h('div', { class: 'pt-field' }, h('label', null, say('كلمة المرور', 'Password')), p1.row),
      h('div', { class: 'pt-field' }, h('label', null, say('تأكيد كلمة المرور', 'Confirm password')), p2.row),
      h('div', { class: 'pt-note' }, say('التشفير AES-256 يتم على جهازك. ⚠ إذا نسيت كلمة المرور فلا يمكن استرجاعها.', 'AES-256 encryption runs on your device. ⚠ If you forget the password it cannot be recovered.')),
      err,
      actionsRow(textButton(say('إلغاء', 'Cancel'), () => modal.close()), okBtn)
    );
    p1.input.focus();
  }

  /* ═════════════ 3) فك حماية PDF ═════════════ */
  async function startUnlock(bytes, name) {
    if (!(await pdfNeedsDecrypt(bytes))) {
      notify(say('هذا الملف غير محمي بكلمة مرور.', 'This file is not password-protected.'));
      return;
    }
    const ready = await ensureReadable(bytes, name);
    saveBytes(ready, baseName(name) + '-unlocked.pdf');
    notify(say('تم فك الحماية وتنزيل الملف', 'Unlocked file downloaded'));
  }

  /* ═════════════ 4) تعبئة نماذج PDF ═════════════ */
  const RTL_RE = /[\u0590-\u08FF]/;

  function canEncodeLatin(text) {
    if (typeof canEncodeWinAnsi === 'function') return canEncodeWinAnsi(text);
    return !/[^\u0009\u000A\u000D\u0020-\u007E\u00A0-\u00FF]/.test(text || '');
  }

  function fieldLabel(field) {
    try {
      const tu = field.acroField.dict.lookup(window.PDFLib.PDFName.of('TU'));
      if (tu && typeof tu.decodeText === 'function') {
        const t = tu.decodeText().trim();
        if (t) return t;
      }
    } catch (e) { /* fall back to the field name */ }
    return field.getName().split('.').pop().replace(/\[\d+\]$/, '') || field.getName();
  }

  function buildPageMap(pdfDoc) {
    const map = new Map();
    pdfDoc.getPages().forEach((page, idx) => {
      try {
        const annots = page.node.Annots();
        if (!annots) return;
        for (let i = 0; i < annots.size(); i++) {
          const ref = annots.get(i);
          if (ref && ref.toString) map.set(ref.toString(), idx);
        }
      } catch (e) { /* ignore */ }
    });
    return map;
  }

  function fieldPageIndex(pdfDoc, field, pageMap) {
    try {
      const widgets = field.acroField.getWidgets();
      if (!widgets.length) return 0;
      const ref = pdfDoc.context.getObjectRef(widgets[0].dict);
      if (ref && pageMap.has(ref.toString())) return pageMap.get(ref.toString());
      const p = widgets[0].P();
      if (p) {
        const idx = pdfDoc.getPages().findIndex((pg) => pg.ref === p);
        if (idx >= 0) return idx;
      }
    } catch (e) { /* ignore */ }
    return 0;
  }

  function renderTextToPng(text, rect, multiline) {
    const k = 3;
    const w = Math.max(8, Math.round(rect.width * k));
    const hgt = Math.max(8, Math.round(rect.height * k));
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = hgt;
    const c2d = canvas.getContext('2d');
    const rtl = RTL_RE.test(text);
    const fs = multiline ? 12 : Math.max(8, Math.min(14, rect.height * 0.65));
    c2d.font = `${fs * k}px "Segoe UI", Tahoma, Arial, sans-serif`;
    c2d.fillStyle = '#000';
    c2d.direction = rtl ? 'rtl' : 'ltr';
    c2d.textAlign = rtl ? 'right' : 'left';
    const pad = 2 * k;
    const x = rtl ? w - pad : pad;
    if (!multiline) {
      c2d.textBaseline = 'middle';
      c2d.fillText(text, x, hgt / 2, w - pad * 2);
    } else {
      c2d.textBaseline = 'top';
      const lines = [];
      text.split(/\r?\n/).forEach((para) => {
        let line = '';
        para.split(/\s+/).forEach((word) => {
          const test = line ? line + ' ' + word : word;
          if (line && c2d.measureText(test).width > w - pad * 2) { lines.push(line); line = word; } else line = test;
        });
        lines.push(line);
      });
      const lh = fs * k * 1.3;
      lines.forEach((ln, i) => { if (pad + i * lh < hgt) c2d.fillText(ln, x, pad + i * lh); });
    }
    return new Promise((resolve) => canvas.toBlob(async (b) => resolve(b ? new Uint8Array(await b.arrayBuffer()) : null), 'image/png'));
  }

  async function startFill(rawBytes, name, opts = {}) {
    const bytes = await ensureReadable(rawBytes, name);
    const lib = window.PDFLib;
    let doc;
    let fields;
    let hasXfa = false;
    try {
      doc = await loadDoc(bytes);
      const form = doc.getForm();
      fields = form.getFields().filter((f) => !(f instanceof lib.PDFButton) && !(f instanceof lib.PDFSignature));
      try { hasXfa = form.acroForm.dict.has(lib.PDFName.of('XFA')); } catch (e) { hasXfa = false; }
    } catch (e) {
      throw new UserError(say('تعذر قراءة النموذج: ', 'Could not read the form: ') + (e.message || e));
    }

    const modal = openModal(say('📝 تعبئة نموذج PDF', '📝 Fill PDF form'));
    if (!fields.length) {
      modal.render(
        h('div', { class: 'pt-note' }, say('لا يحتوي هذا الملف على حقول نموذج قابلة للتعبئة. إذا كان نموذجاً ممسوحاً ضوئياً فافتحه في المحرر واستخدم «إضافة نص».', 'This PDF has no fillable form fields. If it is a scanned form, open it in the editor and use “Add Text”.')),
        actionsRow(
          opts.fromHome && typeof handleFile === 'function' ? textButton(say('فتح في المحرر', 'Open in editor'), () => { modal.close(); handleFile(new File([bytes], name || 'document.pdf', { type: 'application/pdf' })); }, 'pri') : null,
          textButton(say('إغلاق', 'Close'), () => modal.close())
        )
      );
      return;
    }

    const pageMap = buildPageMap(doc);
    const entries = fields.map((field) => {
      const entry = { name: field.getName(), label: fieldLabel(field), page: fieldPageIndex(doc, field, pageMap), readOnly: false };
      try { entry.readOnly = field.isReadOnly(); } catch (e) { /* ignore */ }
      if (field instanceof lib.PDFTextField) {
        entry.kind = 'text';
        entry.multiline = field.isMultiline();
        entry.orig = field.getText() || '';
        entry.maxLen = field.getMaxLength();
        entry.control = entry.multiline ? h('textarea', { rows: 3 }) : h('input', { type: 'text' });
        entry.control.value = entry.orig;
        if (typeof entry.maxLen === 'number' && entry.maxLen > 0) entry.control.maxLength = entry.maxLen;
      } else if (field instanceof lib.PDFCheckBox) {
        entry.kind = 'check';
        entry.orig = field.isChecked();
        entry.control = h('input', { type: 'checkbox' });
        entry.control.checked = entry.orig;
      } else if (field instanceof lib.PDFRadioGroup || field instanceof lib.PDFDropdown) {
        entry.kind = field instanceof lib.PDFRadioGroup ? 'radio' : 'dropdown';
        const selected = field.getSelected();
        entry.orig = (Array.isArray(selected) ? selected[0] : selected) || '';
        entry.control = h('select', null, h('option', { value: '' }, '—'), ...field.getOptions().map((o) => h('option', { value: o }, o)));
        entry.control.value = entry.orig;
      } else if (field instanceof lib.PDFOptionList) {
        entry.kind = 'list';
        entry.orig = field.getSelected();
        entry.control = h('select', { multiple: true, size: Math.min(6, Math.max(2, field.getOptions().length)) }, ...field.getOptions().map((o) => h('option', { value: o }, o)));
        Array.from(entry.control.options).forEach((opt) => { opt.selected = entry.orig.includes(opt.value); });
      } else {
        return null;
      }
      if (entry.readOnly) entry.control.disabled = true;
      return entry;
    }).filter(Boolean);

    const flattenChk = h('input', { type: 'checkbox' });
    const err = h('div', { class: 'pt-err' });
    const okBtn = textButton(say('⬇ حفظ وتنزيل', '⬇ Save & download'), () => save(), 'pri');

    const listNodes = [];
    let lastPage = -1;
    entries.sort((a, b) => a.page - b.page).forEach((entry) => {
      if (entry.page !== lastPage) {
        lastPage = entry.page;
        listNodes.push(h('div', { class: 'pt-page-h' }, say(`الصفحة ${entry.page + 1}`, `Page ${entry.page + 1}`)));
      }
      if (entry.kind === 'check') {
        listNodes.push(h('label', { class: 'pt-check', title: entry.name }, entry.control, h('span', null, entry.label)));
      } else {
        listNodes.push(h('div', { class: 'pt-field' }, h('label', { title: entry.name }, entry.label), entry.control));
      }
    });

    const save = async () => {
      err.textContent = '';
      okBtn.disabled = true;
      const oldLabel = okBtn.textContent;
      okBtn.textContent = say('جارٍ الحفظ…', 'Saving…');
      try {
        const out = await applyFormValues();
        saveBytes(out, baseName(name) + '-filled.pdf');
        modal.close();
        notify(say('تم تنزيل النموذج المعبأ', 'Filled form downloaded'));
      } catch (e) {
        console.error('Form fill failed:', e);
        err.textContent = say('تعذر حفظ النموذج: ', 'Could not save the form: ') + ((e && e.message) || e);
      } finally {
        okBtn.disabled = false;
        okBtn.textContent = oldLabel;
      }
    };

    async function applyFormValues() {
      const pdf = await loadDoc(bytes);
      const form = pdf.getForm();
      const pages = pdf.getPages();
      const overlays = [];
      for (const entry of entries) {
        if (entry.readOnly) continue;
        const field = form.getField(entry.name);
        const ctrl = entry.control;
        if (entry.kind === 'text') {
          const value = ctrl.value;
          if (value === entry.orig) continue;
          if (value && !canEncodeLatin(value)) {
            const widget = field.acroField.getWidgets()[0];
            if (widget) {
              overlays.push({ page: entry.page, rect: widget.getRectangle(), text: value, multiline: entry.multiline });
              form.removeField(field);
            } else {
              field.setText(undefined);
            }
          } else {
            field.setText(value || undefined);
            if (entry.multiline && value) { try { field.setFontSize(12); } catch (e) { /* keep default size */ } }
          }
        } else if (entry.kind === 'check') {
          if (ctrl.checked === entry.orig) continue;
          if (ctrl.checked) field.check(); else field.uncheck();
        } else if (entry.kind === 'radio' || entry.kind === 'dropdown') {
          if (ctrl.value === entry.orig) continue;
          if (ctrl.value) field.select(ctrl.value); else field.clear();
        } else if (entry.kind === 'list') {
          const sel = Array.from(ctrl.selectedOptions).map((o) => o.value);
          if (sel.join('|') === entry.orig.join('|')) continue;
          if (sel.length) field.select(sel); else field.clear();
        }
      }
      const font = await pdf.embedFont(lib.StandardFonts.Helvetica);
      form.updateFieldAppearances(font);
      for (const ov of overlays) {
        const png = await renderTextToPng(ov.text, ov.rect, ov.multiline);
        const page = pages[ov.page];
        if (!png || !page) continue;
        const img = await pdf.embedPng(png);
        page.drawImage(img, { x: ov.rect.x, y: ov.rect.y, width: ov.rect.width, height: ov.rect.height });
      }
      if (flattenChk.checked) form.flatten({ updateFieldAppearances: false });
      return pdf.save({ updateFieldAppearances: false, useObjectStreams: false });
    }

    modal.render(
      h('div', { class: 'pt-note' }, `${name || ''} — ${say(`${entries.length} حقل`, `${entries.length} field(s)`)}`),
      hasXfa ? h('div', { class: 'pt-note' }, say('⚠ هذا نموذج ديناميكي (XFA)؛ قد تظهر بعض القيم بشكل مختلف في Adobe Acrobat.', '⚠ This is a dynamic (XFA) form; some values may display differently in Adobe Acrobat.')) : null,
      h('div', { class: 'pt-scroll' }, ...listNodes),
      h('label', { class: 'pt-check' }, flattenChk, h('span', null, say('تثبيت القيم (لا يمكن تعديل الحقول بعد التنزيل)', 'Flatten (fields cannot be edited after download)'))),
      err,
      actionsRow(textButton(say('إلغاء', 'Cancel'), () => modal.close()), okBtn)
    );
  }

  /* ═════════════ ربط الواجهة ═════════════ */
  function bindHomeButton(id, runner) {
    const button = $id(id);
    if (!button) return;
    button.addEventListener('click', async () => {
      const file = await pickPdf();
      if (!file) return;
      try {
        assertFileSize(file);
        setBusy(true, say('جارٍ قراءة الملف…', 'Reading file…'));
        const bytes = new Uint8Array(await file.arrayBuffer());
        setBusy(false);
        await runner(bytes, file.name || 'document.pdf', { fromHome: true });
      } catch (err) {
        setBusy(false);
        reportError(err);
      }
    });
  }

  bindHomeButton('bCompressPdf', startCompress);
  bindHomeButton('bProtectPdf', startProtect);
  bindHomeButton('bUnlockPdf', startUnlock);
  bindHomeButton('bFillPdf', startFill);

  // أزرار داخل شريط أدوات محرر PDF (تعمل على النسخة الحالية بكل تعديلاتك دون تغيير المحرر نفسه)
  if (typeof renderPdfEditor === 'function' && typeof buildMergedPdfBytesWithAnnotations === 'function') {
    const origRender = renderPdfEditor;
    renderPdfEditor = async function (...args) {
      const res = await origRender.apply(this, args);
      try {
        if (typeof kind === 'undefined' || kind !== 'pdf' || !curPdf.bytes) return res;
        const toolbar = document.querySelector('.pdf-toolbar');
        if (!toolbar || toolbar._pdfToolsAdded) return res;
        toolbar._pdfToolsAdded = true;
        const withCurrent = (runner) => async () => {
          setBusy(true, say('جارٍ تجهيز الملف…', 'Preparing file…'));
          let bytes;
          try {
            bytes = await buildMergedPdfBytesWithAnnotations();
          } catch (err) {
            setBusy(false);
            reportError(err);
            return;
          } finally {
            setBusy(false);
          }
          try { await runner(bytes, fileName || 'document.pdf', {}); } catch (err) { reportError(err); }
        };
        const group = [
          btn(say('🗜 ضغط', '🗜 Compress'), withCurrent(startCompress)),
          btn(say('🔒 حماية', '🔒 Protect'), withCurrent(startProtect)),
          btn(say('📝 تعبئة نموذج', '📝 Fill form'), withCurrent(startFill))
        ];
        const zoomNode = toolbar.querySelector('.zoom-select');
        const spacer = Array.from(toolbar.children).find((c) => c.tagName === 'SPAN' && c.style.flex === '1');
        const anchor = zoomNode || spacer;
        group.forEach((node) => (anchor ? toolbar.insertBefore(node, anchor) : toolbar.append(node)));
      } catch (err) {
        console.warn('PDF tools toolbar integration failed:', err);
      }
      return res;
    };
  }

  // فتح ملف PDF محمي بكلمة مرور في المحرر: يُطلب كلمة المرور ثم يُفك التشفير محلياً
  if (typeof initPdfFromFile === 'function') {
    const origInit = initPdfFromFile;
    initPdfFromFile = async function (file) {
      let target = file;
      let unlocked = false;
      try {
        if (file && (await fileMayBeEncrypted(file))) {
          assertFileSize(file);
          const bytes = new Uint8Array(await file.arrayBuffer());
          if (await pdfNeedsDecrypt(bytes)) {
            const ready = await ensureReadable(bytes, file.name);
            target = new File([ready], file.name || 'document.pdf', { type: 'application/pdf' });
            unlocked = true;
          }
        }
      } catch (err) {
        if (err instanceof SilentCancel) throw new UserError(say('تم إلغاء فتح الملف المحمي.', 'Opening the protected file was cancelled.'));
        throw err;
      }
      const result = await origInit.call(this, target);
      if (unlocked) notify(say('تم فك الحماية للتعديل — الملف الذي ستنزّله لن يكون محمياً (استخدم 🔒 حماية)', 'Unlocked for editing — the file you download will not be protected (use 🔒 Protect)'));
      return result;
    };
  }

  // دمج ملفات محمية: نفك التشفير أولاً حتى لا يُدمج محتوى مشفّر تالف
  if (typeof mergePdfFiles === 'function') {
    const origMerge = mergePdfFiles;
    mergePdfFiles = async function (filesList) {
      const list = Array.from(filesList || []);
      const prepared = [];
      try {
        for (const f of list) {
          if (await fileMayBeEncrypted(f)) {
            const bytes = new Uint8Array(await f.arrayBuffer());
            if (await pdfNeedsDecrypt(bytes)) {
              const ready = await ensureReadable(bytes, f.name);
              prepared.push(new File([ready], f.name || 'document.pdf', { type: 'application/pdf' }));
              continue;
            }
          }
          prepared.push(f);
        }
      } catch (err) {
        if (err instanceof SilentCancel) return;
        reportError(err);
        return;
      }
      return origMerge.call(this, prepared);
    };
  }

  window.PdfTools = {
    ensureReadable,
    compress: startCompress,
    protect: startProtect,
    unlock: startUnlock,
    fill: startFill,
    _compressBytes: compressBytes
  };
})();
