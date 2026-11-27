'use strict';
/* PDF Split — isolated utility module */
(() => {
  const input = document.querySelector('#splitPdfIn');
  const button = document.querySelector('#bSplitPdf');
  if (!input || !button || !window.PDFLib || !window.JSZip) return;

  const isArabic = () => document.documentElement.lang === 'ar';
  const say = (ar, en) => isArabic() ? ar : en;
  const setBusy = (on, message) => {
    if (typeof busy === 'function') busy(on, message);
  };
  const notify = (message) => {
    if (typeof toast === 'function') toast(message);
    else window.alert(message);
  };

  function parseRanges(raw, pageCount) {
    const chunks = raw.split(',').map((part) => part.trim()).filter(Boolean);
    if (!chunks.length) throw new Error(say('أدخل نطاق صفحة واحد على الأقل.', 'Enter at least one page range.'));
    return chunks.map((chunk) => {
      const match = chunk.match(/^(\d+)\s*(?:-\s*(\d+))?$/);
      if (!match) throw new Error(say(`صيغة النطاق غير صحيحة: ${chunk}`, `Invalid range: ${chunk}`));
      const start = Number(match[1]);
      const end = Number(match[2] || match[1]);
      if (start < 1 || end < start || end > pageCount) {
        throw new Error(say(`النطاق ${chunk} خارج حدود الملف (1-${pageCount}).`, `Range ${chunk} is outside the document (1-${pageCount}).`));
      }
      return { start: start - 1, end: end - 1 };
    });
  }

  async function split(file) {
    if (!file) return;
    setBusy(true, say('جارٍ قراءة ملف PDF…', 'Reading PDF…'));
    try {
      const { PDFDocument } = window.PDFLib;
      let fileBytes = new Uint8Array(await file.arrayBuffer());
      if (window.PdfTools) fileBytes = await window.PdfTools.ensureReadable(fileBytes, file.name);
      const source = await PDFDocument.load(fileBytes, { ignoreEncryption: true });
      const count = source.getPageCount();
      if (!count) throw new Error(say('ملف PDF لا يحتوي على صفحات.', 'The PDF contains no pages.'));
      const raw = window.prompt(
        say('أدخل نطاقات الصفحات المطلوب فصلها، مفصولة بفواصل. مثال: 1-3, 4-6, 8', 'Enter page ranges to split, separated by commas. Example: 1-3, 4-6, 8'),
        count > 1 ? `1-${count}` : '1'
      );
      if (raw === null) return;
      const ranges = parseRanges(raw, count);
      const zip = new window.JSZip();
      for (let i = 0; i < ranges.length; i++) {
        const { start, end } = ranges[i];
        const output = await PDFDocument.create();
        const indices = Array.from({ length: end - start + 1 }, (_, offset) => start + offset);
        const pages = await output.copyPages(source, indices);
        pages.forEach((page) => output.addPage(page));
        const bytes = await output.save();
        zip.file(`part-${String(i + 1).padStart(2, '0')}-pages-${start + 1}-${end + 1}.pdf`, bytes);
      }
      const archive = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
      const url = URL.createObjectURL(archive);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${(file.name || 'document').replace(/\.pdf$/i, '')}-split.zip`;
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      notify(say(`تم تقسيم الملف إلى ${ranges.length} ملف/ملفات`, `PDF split into ${ranges.length} file(s)`));
    } catch (error) {
      console.error('PDF split failed:', error);
      notify(say('تعذر تقسيم الملف: ', 'Could not split PDF: ') + (error.message || 'Unknown error'));
    } finally {
      setBusy(false);
      input.value = '';
    }
  }

  button.addEventListener('click', () => { input.value = ''; input.click(); });
  input.addEventListener('change', () => split(input.files && input.files[0]));
})();
