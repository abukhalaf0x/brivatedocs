'use strict';
/* Images to PDF — isolated utility module */
(() => {
  const $ = (s) => document.querySelector(s);
  const input = $('#imagesToPdfIn');
  const button = $('#bImagesToPdf');
  if (!input || !button || !window.PDFLib) return;

  const isArabic = () => document.documentElement.lang === 'ar';
  const say = (ar, en) => isArabic() ? ar : en;
  const setBusy = (on, message) => {
    if (typeof busy === 'function') busy(on, message);
  };
  const notify = (message) => {
    if (typeof toast === 'function') toast(message);
    else window.alert(message);
  };
  const safeBaseName = (name) => (name || 'images').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80) || 'images';

  async function imageFileToPngBytes(file) {
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.decoding = 'async';
      img.src = url;
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = () => reject(new Error('Unsupported or unreadable image'));
      });
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('Canvas is unavailable');
      ctx.drawImage(img, 0, 0);
      const blob = await new Promise((resolve, reject) =>
        canvas.toBlob((value) => value ? resolve(value) : reject(new Error('Image conversion failed')), 'image/png')
      );
      return new Uint8Array(await blob.arrayBuffer());
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  async function convert(files) {
    const selected = Array.from(files || []).filter((file) => file && file.type.startsWith('image/'));
    if (!selected.length) return;
    setBusy(true, say('جارٍ تحويل الصور إلى PDF…', 'Converting images to PDF…'));
    try {
      const { PDFDocument } = window.PDFLib;
      const pdf = await PDFDocument.create();
      const pageW = 595.28;
      const pageH = 841.89;
      const margin = 28;
      for (const file of selected) {
        if (file.size > 40 * 1024 * 1024) throw new Error(`${file.name}: ${say('حجم الصورة يتجاوز 40MB', 'image exceeds 40MB')}`);
        const pngBytes = await imageFileToPngBytes(file);
        const image = await pdf.embedPng(pngBytes);
        const page = pdf.addPage([pageW, pageH]);
        const scale = Math.min((pageW - margin * 2) / image.width, (pageH - margin * 2) / image.height);
        const width = image.width * scale;
        const height = image.height * scale;
        page.drawImage(image, {
          x: (pageW - width) / 2,
          y: (pageH - height) / 2,
          width,
          height
        });
      }
      const bytes = await pdf.save();
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${safeBaseName(selected[0].name)}${selected.length > 1 ? '-images' : ''}.pdf`;
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      notify(say('تم تحويل الصور إلى PDF بنجاح', 'Images converted to PDF successfully'));
    } catch (error) {
      console.error('Images to PDF failed:', error);
      notify(say('تعذر تحويل الصور: ', 'Could not convert images: ') + (error.message || 'Unknown error'));
    } finally {
      setBusy(false);
      input.value = '';
    }
  }

  button.addEventListener('click', () => { input.value = ''; input.click(); });
  input.addEventListener('change', () => convert(input.files));
})();
