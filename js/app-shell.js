'use strict';
/* ══════════════════════════════════════════════════════════════
   brivatedocs — Shared Media Triggers, Find/Replace, Export & Shell Init
   ══════════════════════════════════════════════════════════════ */
/* ── أدوات رفع الصور والوسائط والبحث والتحميل ── */
function triggerImageInsert(onReady) {
  const inp = $('#imgIn');
  inp.value = '';
  inp.onchange = async () => {
    const f = inp.files?.[0];
    if (!f) return;
    busy(true, 'جارٍ تجهيز الصورة…');
    try { onReady(await prepImage(f)); }
    catch (err) { toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة'); }
    finally { busy(false); }
  };
  inp.click();
}

function triggerReplaceImage(onReady) {
  const inp = $('#repImgIn');
  inp.value = '';
  inp.onchange = async () => {
    const f = inp.files?.[0];
    if (!f) return;
    busy(true, 'جارٍ تبديل الصورة…');
    try { onReady(await prepImage(f)); }
    catch (err) { toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة'); }
    finally { busy(false); }
  };
  inp.click();
}

function triggerBgImageInsert(onReady) {
  const inp = $('#bgImgIn');
  inp.value = '';
  inp.onchange = async () => {
    const f = inp.files?.[0];
    if (!f) return;
    busy(true, 'جارٍ تجهيز صورة الخلفية…');
    try { onReady(await prepImage(f)); }
    catch (err) { toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة'); }
    finally { busy(false); }
  };
  inp.click();
}

function triggerMediaInsert(onReady) {
  const inp = $('#mediaIn');
  inp.value = '';
  inp.onchange = async () => {
    const f = inp.files?.[0];
    if (!f) return;
    busy(true, 'جارٍ تجهيز ملف الوسائط…');
    try {
      const bytes = new Uint8Array(await f.arrayBuffer());
      const ext = (f.name.split('.').pop() || 'mp4').toLowerCase();
      const isAudio = f.type.startsWith('audio/') || ['mp3', 'wav'].includes(ext);
      onReady({ bytes, ext, mime: f.type || (isAudio ? 'audio/mpeg' : 'video/mp4'), isAudio });
    } catch (_) { toast('تعذر قراءة ملف الوسائط'); }
    finally { busy(false); }
  };
  inp.click();
}

let lastTopBarHeight = 0;
function syncEditorTopBarForKind() {
  const isPdf = kind === 'pdf';
  const findButton = $('#bFind');
  if (findButton) { findButton.disabled = false; findButton.hidden = false; }
  const snapButton = $('#bSnapshots');
  if (snapButton) { snapButton.disabled = isPdf; snapButton.hidden = isPdf; }
  const topEl = $('#editor .top');
  if (topEl) {
    const hPx = topEl.offsetHeight;
    if (hPx > 0 && hPx !== lastTopBarHeight) {
      lastTopBarHeight = hPx;
      document.documentElement.style.setProperty('--top-h', hPx + 'px');
    }
  }
}

function renderEditor() {
  syncEditorTopBarForKind();
  if (kind === 'pptx') renderPptxEditor();
  else if (kind === 'docx') renderDocxEditor();
}

function replaceAcrossTextNodes(container, ns, localName, search, replacement) {
  const nodes = desc(container, ns, localName).filter(n => n.textContent);
  if (!nodes.length) return 0;
  const values = nodes.map(n => n.textContent);
  const full = values.join('');
  const matches = [];
  let at = 0;
  while ((at = full.indexOf(search, at)) !== -1) { matches.push(at); at += search.length; }
  const locate = (pos) => {
    let offset = pos;
    for (let i = 0; i < nodes.length; i++) {
      if (offset <= values[i].length) return { i, offset };
      offset -= values[i].length;
    }
    return { i: nodes.length - 1, offset: values.at(-1).length };
  };
  for (const pos of matches.reverse()) {
    const a = locate(pos), b = locate(pos + search.length);
    if (a.i === b.i) nodes[a.i].textContent = values[a.i].slice(0, a.offset) + replacement + values[a.i].slice(b.offset);
    else {
      nodes[a.i].textContent = values[a.i].slice(0, a.offset) + replacement;
      for (let i = a.i + 1; i < b.i; i++) nodes[i].textContent = '';
      nodes[b.i].textContent = values[b.i].slice(b.offset);
    }
    for (let i = 0; i < nodes.length; i++) values[i] = nodes[i].textContent;
  }
  return matches.length;
}

function executeFindAndReplace(findStr, repStr) {
  if (!findStr || repStr == null || !pk || kind === 'pdf') return 0;
  if (kind === 'docx') {
    const paper = $('#app .paper');
    if (paper) syncPaperToDocxXml(paper);
  }
  pk.commit();
  let count = 0;
  for (const [name, content] of pk.files.entries()) {
    if (typeof content !== 'string' || !name.endsWith('.xml')) continue;
    const d = pk.doc(name);
    const ns = name.startsWith('ppt/') ? A : W;
    const paragraphs = desc(d.documentElement, ns, 'p');
    for (const para of paragraphs) count += replaceAcrossTextNodes(para, ns, 't', findStr, repStr);
  }
  if (count > 0) {
    pk.commit();
    pushHistory('replace');
    renderEditor();
    toast(currentLanguage === 'en' ? `Replaced ${count} occurrence(s)` : `تم استبدال ${count} موضع`);
  } else {
    toast('لم يتم العثور على النص المطلوب');
  }
  return count;
}

function findAndReplace() {
  if (kind === 'pdf') {
    if (typeof openPdfSearchModal === 'function') openPdfSearchModal();
    return;
  }
  if (!pk) {
    toast(currentLanguage === 'en' ? 'Open a file first' : 'افتح ملفاً أولاً');
    return;
  }
  if (kind === 'docx') {
    const paper = $('#app .paper');
    if (paper) syncPaperToDocxXml(paper);
  }
  $('#findReplaceBackdrop')?.remove();
  const findIn = h('input', {
    id: 'frFindInput',
    type: 'text',
    placeholder: tr('النص المطلوب البحث عنه'),
    style: { width: '100%', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px 10px', background: 'var(--soft)', color: 'var(--ink)' }
  });
  const repIn = h('input', {
    id: 'frReplaceInput',
    type: 'text',
    placeholder: tr('النص البديل'),
    style: { width: '100%', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px 10px', background: 'var(--soft)', color: 'var(--ink)' }
  });
  const runReplace = () => {
    const fVal = findIn.value;
    if (!fVal) { findIn.focus(); return; }
    const rVal = repIn.value ?? '';
    backdrop.remove();
    executeFindAndReplace(fVal, rVal);
  };
  findIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); repIn.focus(); } });
  repIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); runReplace(); } });

  const backdrop = h('div', { id: 'findReplaceBackdrop', class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) backdrop.remove(); } },
    h('div', { class: 'md-modal', style: { width: 'min(440px, 95vw)' } },
      h('strong', null, '🔎 ' + tr('البحث والاستبدال')),
      h('div', { class: 'fld' }, h('label', null, tr('النص المطلوب البحث عنه')), findIn),
      h('div', { class: 'fld' }, h('label', null, tr('النص البديل')), repIn),
      h('div', { class: 'row', style: { justifyContent: 'flex-end', marginTop: '4px' } },
        btn('إلغاء', () => backdrop.remove()),
        btn('استبدال الكل', runReplace, 'pri')
      )
    )
  );
  document.body.append(backdrop);
  findIn.focus();
}

function sortWordXmlChildren(parent, order) {
  const rank = new Map(order.map((name, index) => [name, index]));
  const children = kids(parent, W).map((node, index) => ({ node, index, rank: rank.has(node.localName) ? rank.get(node.localName) : 10000 + index }));
  children.sort((a, b) => a.rank - b.rank || a.index - b.index);
  for (const item of children) parent.appendChild(item.node);
}
function normalizeWordXmlPackage(wordPkg) {
  const pPrOrder = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','jc','textDirection','textAlignment','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
  const rPrOrder = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange'];
  const tcPrOrder = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark','headers','cellIns','cellDel','cellMerge','tcPrChange'];
  const sectPrOrder = ['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];
  for (const [name, content] of wordPkg.files.entries()) {
    if (typeof content !== 'string' || !name.startsWith('word/') || !name.endsWith('.xml')) continue;
    const doc = wordPkg.doc(name);
    for (const el of desc(doc.documentElement, W, 'pPr')) sortWordXmlChildren(el, pPrOrder);
    for (const el of desc(doc.documentElement, W, 'rPr')) sortWordXmlChildren(el, rPrOrder);
    for (const el of desc(doc.documentElement, W, 'tcPr')) sortWordXmlChildren(el, tcPrOrder);
    for (const el of desc(doc.documentElement, W, 'sectPr')) sortWordXmlChildren(el, sectPrOrder);
    for (const tbl of desc(doc.documentElement, W, 'tbl')) {
      sortWordXmlChildren(tbl, ['tblPr', 'tblGrid', 'tr', 'customXml', 'sdt', 'tblPrEx', 'tblGridChange', 'trPrChange']);
      if (kid(tbl, W, 'tblGrid')) continue;
      const firstRow = kid(tbl, W, 'tr');
      if (!firstRow) continue;
      const count = Math.max(1, kids(firstRow, W, 'tc').length);
      const grid = mk(doc, W, 'w:tblGrid');
      for (let i = 0; i < count; i++) grid.appendChild(mk(doc, W, 'w:gridCol', { 'w:w': String(Math.floor(9360 / count)) }));
      const tblPr = kid(tbl, W, 'tblPr');
      if (tblPr) insertAfter(tblPr, grid); else tbl.insertBefore(grid, firstRow);
    }
  }
  wordPkg.commit();
}

async function downloadEditedFile() {
  if (!pk) return;
  busy(true, 'جارٍ تجهيز الملف للتحميل…');
  try {
    const mime = kind === 'pptx'
      ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    if (kind === 'docx') normalizeWordXmlPackage(pk);
    const blob = await pk.toBlob(mime);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace(/(\.(docx|pptx|docm|pptm))?$/i, (m) => m || (kind === 'pptx' ? '.pptx' : '.docx'));
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    if (activeDraftId) await DraftDB.removeById(activeDraftId);
    toast('تم تحميل الملف');
  } catch (_) { toast('حدث خطأ أثناء تصدير الملف'); }
  finally { busy(false); }
}


(function initAppShell() {
/* ── ربط الأحداث الرئيسية، تحجيم صور DOCX، سجل الإصدارات، واختصارات الكيبورد ── */
const drop = $('#drop'), fileIn = $('#fileIn');
if (drop && fileIn) {
  drop.addEventListener('click', () => fileIn.click());
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault(); drop.classList.remove('over');
    handleFile(e.dataTransfer?.files?.[0]);
  });
  fileIn.addEventListener('change', () => { handleFile(fileIn.files?.[0]); fileIn.value = ''; });
}

$('#bNewPptx')?.addEventListener('click', () => startNewDeck($('#newThemeSel')?.value || 'light'));
$('#bNewDocx')?.addEventListener('click', () => startNewDocx($('#newDocxTplSel')?.value || 'blank'));

$('#bDiscardDraft')?.addEventListener('click', async () => {
  await DraftDB.clear();
  const banner = $('#draftBanner');
  if (banner) banner.hidden = true;
  toast('تم تجاهل المسودة السابقة');
});

$('#bHome')?.addEventListener('click', () => showView('home'));
$('#bFind')?.addEventListener('click', findAndReplace);
$('#bUndo')?.addEventListener('click', () => applySnapshot(histIdx - 1));
$('#bRedo')?.addEventListener('click', () => applySnapshot(histIdx + 1));
$('#bDown')?.addEventListener('click', () => downloadEditedFile());

const fnEl = $('#fname');
if (fnEl) {
  fnEl.addEventListener('click', () => {
    if (!fileName) return;
    const extMatch = fileName.match(/\.(docx|pptx|docm|pptm|pdf)$/i);
    const ext = extMatch ? extMatch[0] : (kind === 'pdf' ? '.pdf' : (kind === 'pptx' ? '.pptx' : '.docx'));
    const base = fileName.replace(/\.(docx|pptx|docm|pptm|pdf)$/i, '');
    const nextName = prompt(currentLanguage === 'en' ? 'File name:' : 'اسم الملف الجديد:', base);
    if (nextName && nextName.trim()) {
      fileName = nextName.trim().replace(/\.(docx|pptx|docm|pptm|pdf)$/i, '') + ext;
      fnEl.textContent = fileName;
      if (pk) pushHistory('renameFile');
      toast(currentLanguage === 'en' ? 'File name updated' : 'تم تحديث اسم الملف');
    }
  });
}

const topBar = $('#editor .top');
if (topBar && typeof ResizeObserver !== 'undefined') {
  let topBarRaf = 0;
  const ro = new ResizeObserver(() => {
    if (topBarRaf) return;
    topBarRaf = requestAnimationFrame(() => {
      topBarRaf = 0;
      const hPx = topBar.offsetHeight;
      if (hPx > 0 && hPx !== lastTopBarHeight) {
        lastTopBarHeight = hPx;
        document.documentElement.style.setProperty('--top-h', hPx + 'px');
      }
    });
  });
  ro.observe(topBar);
}
if (topBar) {
  const downBtn = $('#bDown');
  if (!$('#bFocusMode')) {
    const focusBtn = btn('🎯 وضع التركيز', () => {
      const on = document.body.classList.toggle('focus-mode');
      focusBtn.classList.toggle('on', on);
    });
    focusBtn.id = 'bFocusMode';
    if (downBtn) topBar.insertBefore(focusBtn, downBtn);
    else topBar.append(focusBtn);
  }
  if (!$('#bSnapshots')) {
    const snapshots = [];
    const snapBtn = btn('🕒 النسخ', () => {
      if (!pk) return;
      const promptText = snapshots.length > 0
        ? (currentLanguage === 'en'
            ? `Saved Versions (${snapshots.length}):\nType a name to save current version, OR enter version number (1-${snapshots.length}) to restore:`
            : `النسخ المحفوظة (${snapshots.length}):\nاكتب اسماً لحفظ نسخة الآن، أو اكتب رقم النسخة (1-${snapshots.length}) لاستعادتها:`)
        : (currentLanguage === 'en'
            ? 'No saved versions yet.\nType a name to save the current version:'
            : 'لا توجد نسخ محفوظة بعد.\nاكتب اسماً لحفظ نسخة من المستند الآن:');
      const choice = prompt(promptText);
      if (!choice || !choice.trim()) return;
      const num = parseInt(choice.trim(), 10);
      if (!isNaN(num) && snapshots[num - 1]) {
        pk.restore(snapshots[num - 1].snap);
        pushHistory('restoreVersion');
        renderEditor();
        toast('تمت استعادة النسخة بنجاح');
      } else {
        snapshots.push({ name: choice.trim(), snap: pk.snapshot(), time: new Date().toLocaleTimeString() });
        toast(`تم حفظ النسخة (${snapshots.length}): ${choice.trim()}`);
      }
    });
    snapBtn.id = 'bSnapshots';
    if (downBtn) topBar.insertBefore(snapBtn, downBtn);
    else topBar.append(snapBtn);
  }
}

let activeResizer = null;
const removeResizer = () => { if (activeResizer) { activeResizer.remove(); activeResizer = null; } };

document.addEventListener('pointerdown', (e) => {
  if (e.target.closest?.('.docx-img-resizer')) return;
  removeResizer();
  if (kind !== 'docx' || $('#editor').hidden) return;
  const img = e.target.closest?.('.paper img');
  const paper = $('.paper');
  if (!img || !paper) return;

  const positionResizer = () => {
    if (!activeResizer || !paper.contains(img)) return removeResizer();
    const pRect = paper.getBoundingClientRect();
    const iRect = img.getBoundingClientRect();
    const scale = (+curDocx.zoom || 100) / 100;
    activeResizer.style.left = ((iRect.left - pRect.left) / scale - 7) + 'px';
    activeResizer.style.top = ((iRect.bottom - pRect.top) / scale - 7) + 'px';
  };

  activeResizer = h('div', { class: 'docx-img-resizer', contenteditable: 'false' });
  paper.append(activeResizer);
  positionResizer();

  activeResizer.addEventListener('pointerdown', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const startX = ev.clientX;
    const startW = img.getBoundingClientRect().width / ((+curDocx.zoom || 100) / 100);
    const onMove = (mv) => {
      const dx = startX - mv.clientX;
      const nextW = Math.max(40, Math.min(paper.clientWidth - 40, Math.round(startW + dx)));
      img.style.width = nextW + 'px';
      positionResizer();
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      syncPaperToDocxXml(paper);
      pushHistory('docImgResize');
      updateDocxPagination(paper);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  });
});

document.addEventListener('paste', async (e) => {
  if ($('#editor')?.hidden) return;
  const clipItems = Array.from(e.clipboardData?.items || []);
  const imgItem = clipItems.find(it => it.kind === 'file' && /^image\//i.test(it.type || ''));
  if (imgItem) {
    const imgFile = imgItem.getAsFile();
    if (imgFile) {
      if (kind === 'docx') {
        const paper = $('#app .paper');
        if (paper) {
          e.preventDefault();
          busy(true, 'جارٍ تجهيز الصورة…');
          try {
            saveDocxSelection(paper);
            const info = await prepImage(imgFile);
            insertDocxImage(paper, info);
            updateDocxPagination(paper);
            toast('تم لصق الصورة من الحافظة');
          } catch (err) {
            toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة');
          } finally {
            busy(false);
          }
          return;
        }
      } else if (kind === 'pptx' && curPptx.deck && !document.activeElement?.isContentEditable) {
        e.preventDefault();
        busy(true, 'جارٍ تجهيز الصورة…');
        try {
          const info = await prepImage(imgFile);
          addImageToSlide(curPptx.deck, curPptx.slideIdx, info);
          toast('تم لصق الصورة من الحافظة');
        } catch (err) {
          toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة');
        } finally {
          busy(false);
        }
        return;
      } else if (kind === 'pdf' && curPdf?.bytes && typeof insertPdfImageFileOnCurrentPage === 'function' && !document.activeElement?.isContentEditable) {
        e.preventDefault();
        await insertPdfImageFileOnCurrentPage(imgFile, true);
        return;
      }
    }
  }

  if (kind !== 'docx' || !e.target.closest?.('.paper')) return;
  const html = e.clipboardData?.getData('text/html');
  const plain = e.clipboardData?.getData('text/plain');
  if (html && plain && /<(meta|style|script|font|o:p)/i.test(html)) {
    e.preventDefault();
    document.execCommand('insertText', false, plain);
  }
});

document.addEventListener('keydown', (e) => {
  if (!$('#editor').hidden && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault();
    downloadEditedFile();
    return;
  }
  if (!$('#editor').hidden && $('#slideshowOverlay').hidden && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
    e.preventDefault();
    findAndReplace();
    return;
  }
  if (!$('#editor').hidden && kind === 'docx' && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
    e.preventDefault();
    window.print();
    return;
  }

  if ($('#editor').hidden || !$('#slideshowOverlay').hidden) return;
  if (['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName)) return;

  if (kind === 'pdf' && !document.activeElement?.isContentEditable) {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      (e.shiftKey ? $('#bRedo') : $('#bUndo'))?.click();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      $('#bRedo')?.click();
      return;
    }
    return;
  }

  if (kind === 'docx' && !document.activeElement?.isContentEditable) {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      applySnapshot(e.shiftKey ? histIdx + 1 : histIdx - 1);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      applySnapshot(histIdx + 1);
      return;
    }
    return;
  }

  if (kind !== 'pptx' || document.activeElement?.isContentEditable) return;

  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
    e.preventDefault();
    if (e.shiftKey) applySnapshot(histIdx + 1);
    else applySnapshot(histIdx - 1);
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
    e.preventDefault();
    applySnapshot(histIdx + 1);
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
    if (curPptx.deck) copySelectedShape(curPptx.deck, curPptx.slideIdx);
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
    if (curPptx.deck) pasteShapeToSlide(curPptx.deck, curPptx.slideIdx);
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
    if (curPptx.deck && curPptx.selShapeIdx != null) {
      e.preventDefault();
      duplicateSelectedShape(curPptx.deck, curPptx.slideIdx);
    }
  } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
    if (curPptx.deck) {
      e.preventDefault();
      const list = shapesOf(pk.doc(curPptx.deck.slides[curPptx.slideIdx].path));
      curPptx.multiSel = new Set(list.map((_, i) => i));
      curPptx.selShapeIdx = list.length ? list.length - 1 : null;
      renderPptxEditor();
    }
  } else if (e.key === 'Delete' || e.key === 'Backspace') {
    if (curPptx.deck && (curPptx.selShapeIdx != null || curPptx.multiSel.size)) {
      e.preventDefault();
      deleteSelectedShape(curPptx.deck, curPptx.slideIdx);
    }
  } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
    if (curPptx.deck && curPptx.selShapeIdx != null) {
      e.preventDefault();
      const doc = pk.doc(curPptx.deck.slides[curPptx.slideIdx].path);
      const sh = shapesOf(doc)[curPptx.selShapeIdx];
      if (!sh || isShapeLocked(sh)) return;
      const g = readXfrm(xfrmOf(sh));
      if (!g) return;
      const step = (e.shiftKey ? 10 : 1) * 9525;
      if (e.key === 'ArrowLeft') g.x -= step;
      if (e.key === 'ArrowRight') g.x += step;
      if (e.key === 'ArrowUp') g.y -= step;
      if (e.key === 'ArrowDown') g.y += step;
      writeGeomToShape(sh, g);
      const selEl = $(`.stagewrap .shape[data-i="${curPptx.selShapeIdx}"]`);
      if (selEl) {
        selEl._g = g;
        placeEl(selEl, g, { ox: 0, oy: 0, w: curPptx.deck.W, h: curPptx.deck.H });
      }
      scheduleHistory('nudge', 350);
    }
  }
});

setupTheme();
setupLanguage();
showView('home');
})();

/* ── تنظيف نصوص undefined و null من لوحة خصائص PPTX ── */
(function () {
  const cleanPropsPanel = (panel) => {
    if (!panel) return;
    const walker = document.createTreeWalker(panel, NodeFilter.SHOW_TEXT);
    const toRemove = [];
    while (walker.nextNode()) {
      const n = walker.currentNode;
      const t = (n.nodeValue || '').trim();
      if (t === 'undefined' || t === 'null') {
        toRemove.push(n);
      } else if (n.nodeValue && (n.nodeValue.includes('undefined') || n.nodeValue.includes('null'))) {
        n.nodeValue = n.nodeValue
          .replace(/\bundefined\b/g, '')
          .replace(/\bnull\b/g, '')
          .replace(/\s{2,}/g, ' ');
      }
    }
    toRemove.forEach((n) => n.remove());
  };

  const attachObserver = () => {
    const appEl = document.getElementById('app');
    if (!appEl) return;
    const obs = new MutationObserver(() => {
      appEl.querySelectorAll('.props').forEach(cleanPropsPanel);
    });
    obs.observe(appEl, { childList: true, subtree: true, characterData: true });
    appEl.querySelectorAll('.props').forEach(cleanPropsPanel);
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', attachObserver);
  } else {
    attachObserver();
  }
})();
