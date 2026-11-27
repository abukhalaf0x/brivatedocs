'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 6 (أ): وضع مقدم العرض، العرض التقديمي، وتصدير PNG / ZIP / PDF
   ═════════════════════════════════════════════════════════════ */

/* ── وضع مقدم العرض (Presenter View) ── */
let presenterWin = null;
let presenterState = { deck: null, idx: 0 };

function openPresenterView(deck, startIdx) {
  presenterState = { deck, idx: startIdx };
  curPptx.slideIdx = startIdx;
  openInlinePresenterModal(deck);
}

function openPopupPresenterWindow(deck, startIdx) {
  presenterState = { deck, idx: startIdx };
  if (presenterWin && !presenterWin.closed) {
    presenterWin.focus();
    syncPresenterWindow();
    return;
  }

  presenterWin = window.open('', 'DocStudioPresenterView', 'width=1180,height=740');
  if (!presenterWin) {
    toast('اسمح بالنوافذ المنبثقة لفتح وضع مقدم العرض');
    return;
  }

  const isEn = currentLanguage === 'en';
  const dir = isEn ? 'ltr' : 'rtl';
  const styles = Array.from(document.querySelectorAll('link[rel="stylesheet"], style')).map(s => {
    if (s.tagName === 'LINK') {
      const c = s.cloneNode(true);
      c.href = s.href;
      return c.outerHTML;
    }
    return s.outerHTML;
  }).join('');

  presenterWin.document.write(`<!doctype html>
<html lang="${currentLanguage}" dir="${dir}" data-theme="${currentTheme}">
<head>
<meta charset="utf-8">
<title>${isEn ? 'Presenter View — brivatedocs' : 'وضع مقدم العرض — brivatedocs'}</title>
${styles}
<style>
  body{margin:0;background:var(--bg);color:var(--ink);display:flex;flex-direction:column;height:100vh;overflow:hidden}
  .pv-top{background:var(--card);border-bottom:1px solid var(--line);padding:10px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px}
  .pv-timer{font-family:monospace;font-size:1.35rem;font-weight:800;color:var(--acc);background:var(--soft);border:1px solid var(--line);padding:4px 12px;border-radius:8px}
  .pv-controls{display:flex;align-items:center;gap:8px}
  .pv-main{flex:1;display:grid;grid-template-columns:1.55fr 1fr;gap:14px;padding:14px;overflow:hidden}
  .pv-col{display:flex;flex-direction:column;gap:12px;min-height:0}
  .pv-card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px;display:flex;flex-direction:column;gap:8px;min-height:0}
  .pv-card h4{margin:0;font-size:.88rem;color:var(--mut)}
  .pv-stage{flex:1;display:flex;align-items:center;justify-content:center;background:var(--stage-bg);border-radius:8px;padding:10px;overflow:hidden}
  .pv-stage .slide{max-width:100%;max-height:100%;width:auto!important;height:100%!important;pointer-events:none}
  .pv-next-wrap{height:210px;display:flex;align-items:center;justify-content:center;background:var(--stage-bg);border-radius:8px;padding:8px;overflow:hidden}
  .pv-next-wrap .slide{max-width:100%;max-height:100%;width:auto!important;height:100%!important;pointer-events:none}
  .pv-notes{flex:1;overflow:auto;background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:12px;font-size:1.05rem;line-height:1.7;white-space:pre-wrap}
</style>
</head>
<body>
  <div class="pv-top">
    <div style="display:flex;align-items:center;gap:12px">
      <strong>🖥 ${isEn ? 'Presenter View' : 'وضع مقدم العرض'}</strong>
      <span class="pv-timer" id="pvTimer">00:00:00</span>
      <button class="btn" id="pvResetTimer">${isEn ? 'Reset' : 'تصفير'}</button>
    </div>
    <div class="pv-controls">
      <button class="btn" id="pvPrev">◀ ${isEn ? 'Previous' : 'السابق'}</button>
      <span id="pvCounter" style="font-weight:700">1 / 1</span>
      <button class="btn pri" id="pvNext">${isEn ? 'Next' : 'التالي'} ▶</button>
    </div>
  </div>
  <div class="pv-main">
    <div class="pv-col">
      <div class="pv-card" style="flex:1">
        <h4>${isEn ? 'Current Slide' : 'الشريحة الحالية'}</h4>
        <div class="pv-stage" id="pvCurrentStage"></div>
      </div>
    </div>
    <div class="pv-col">
      <div class="pv-card">
        <h4>${isEn ? 'Next Slide' : 'الشريحة القادمة'}</h4>
        <div class="pv-next-wrap" id="pvNextStage"></div>
      </div>
      <div class="pv-card" style="flex:1">
        <h4>📝 ${isEn ? 'Speaker Notes' : 'ملاحظات المحاضر'}</h4>
        <div class="pv-notes" id="pvNotesBox"></div>
      </div>
    </div>
  </div>
  <script>
    let sec = 0;
    const tEl = document.getElementById('pvTimer');
    const fmt = (s) => [Math.floor(s/3600), Math.floor((s%3600)/60), s%60].map(n => String(n).padStart(2,'0')).join(':');
    setInterval(() => { sec++; tEl.textContent = fmt(sec); }, 1000);
    document.getElementById('pvResetTimer').onclick = () => { sec = 0; tEl.textContent = fmt(0); };
    document.getElementById('pvPrev').onclick = () => window.opener && window.opener.__pvStep && window.opener.__pvStep(-1);
    document.getElementById('pvNext').onclick = () => window.opener && window.opener.__pvStep && window.opener.__pvStep(1);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') window.opener && window.opener.__pvStep && window.opener.__pvStep(1);
      else if (e.key === 'ArrowLeft' || e.key === 'PageUp') window.opener && window.opener.__pvStep && window.opener.__pvStep(-1);
    });
  <\/script>
</body>
</html>`);
  presenterWin.document.close();
  setTimeout(syncPresenterWindow, 80);
}

window.__pvStep = function (delta) {
  if (!curPptx.deck) return;
  const next = curPptx.slideIdx + delta;
  if (next < 0 || next >= curPptx.deck.slides.length) return;
  curPptx.slideIdx = next;
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  renderPptxEditor();
  if (!$('#slideshowOverlay').hidden && ssState.visibleIndices) {
    const pos = ssState.visibleIndices.indexOf(next);
    if (pos >= 0) { ssState.pos = pos; renderSlideshowSlide(); }
  }
};

function syncPresenterWindow() {
  if (!presenterWin || presenterWin.closed || !curPptx.deck) return;
  const doc = presenterWin.document;
  const curStage = doc.getElementById('pvCurrentStage');
  const nextStage = doc.getElementById('pvNextStage');
  const notesBox = doc.getElementById('pvNotesBox');
  const counter = doc.getElementById('pvCounter');
  if (!curStage || !nextStage) return;

  presenterWin.document.documentElement.setAttribute('data-theme', currentTheme);

  const deck = curPptx.deck;
  const idx = curPptx.slideIdx;
  const total = deck.slides.length;
  if (counter) counter.textContent = `${idx + 1} / ${total}`;

  curStage.innerHTML = '';
  const curEl = renderSlideEl(deck, idx, false);
  curEl.classList.remove('show-grid');
  curStage.appendChild(doc.importNode(curEl, true));

  nextStage.innerHTML = '';
  if (idx + 1 < total) {
    const nxtEl = renderSlideEl(deck, idx + 1, false);
    nxtEl.classList.remove('show-grid');
    nextStage.appendChild(doc.importNode(nxtEl, true));
  } else {
    nextStage.textContent = currentLanguage === 'en' ? 'End of presentation' : 'نهاية العرض التقديمي';
  }

  if (notesBox) {
    const notes = getSlideNotes(deck.slides[idx].path).trim();
    notesBox.textContent = notes || (currentLanguage === 'en' ? 'No speaker notes for this slide.' : 'لا توجد ملاحظات لهذه الشريحة.');
  }
}

/* ── وضع العرض التقديمي (Fullscreen Slideshow + حركات العناصر + الليزر + القلم) ── */
let ssState = { visibleIndices: [], pos: 0, deck: null, laser: false, pen: false, drawing: false, autoTimer: null };

function stopSlideshowAutoplay() {
  if (ssState.autoTimer) {
    clearInterval(ssState.autoTimer);
    ssState.autoTimer = null;
  }
  $('#ssAutoBtn')?.classList.remove('on');
}

function setupSlideshowPenCanvas() {
  const cvs = $('#ssPenCanvas');
  if (!cvs) return;
  cvs.width = window.innerWidth;
  cvs.height = window.innerHeight;
  const ctx = cvs.getContext('2d');
  ctx.clearRect(0, 0, cvs.width, cvs.height);
  ctx.strokeStyle = '#ef4444';
  ctx.lineWidth = 3.5;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
}

function startSlideshow(deck, startIdx) {
  const freshDeck = readDeck();
  const visible = freshDeck.slides.map((s, i) => (!s.hidden ? i : -1)).filter(i => i >= 0);
  if (!visible.length) { toast('جميع الشرائح مخفية!'); return; }
  const startPos = Math.max(0, visible.indexOf(startIdx));
  stopSlideshowAutoplay();
  ssState = { visibleIndices: visible, pos: startPos, deck: freshDeck, laser: false, pen: false, drawing: false, autoTimer: null };
  const ov = $('#slideshowOverlay');
  ov.hidden = false;
  ov.style.display = 'flex';
  if ($('#ssPenCanvas')) $('#ssPenCanvas').hidden = true;
  if ($('#ssClearPenBtn')) $('#ssClearPenBtn').hidden = true;
  $('#ssLaserBtn')?.classList.remove('on');
  $('#ssPenBtn')?.classList.remove('on');
  renderSlideshowSlide();
  if (ov.requestFullscreen) {
    ov.requestFullscreen().then(() => {
      requestAnimationFrame(() => renderSlideshowSlide());
    }).catch(() => {});
  }
}

function renderSlideshowSlide() {
  const { deck, visibleIndices, pos } = ssState;
  if (!deck || !visibleIndices.length) return;
  const realIdx = visibleIndices[pos];
  const slPath = deck.slides[realIdx].path;
  const st = $('#ssStage');
  st.innerHTML = '';

  const slEl = renderSlideEl(deck, realIdx, false);
  slEl.classList.remove('show-grid');
  const tr = getSlideTransition(slPath);
  if (tr && tr !== 'none') slEl.classList.add('tr-' + tr);

  const vw = window.innerWidth * 0.94;
  const vh = window.innerHeight * 0.88;
  const ratio = (deck.W || 12192000) / (deck.H || 6858000);
  let targetW = vw;
  let targetH = targetW / ratio;
  if (targetH > vh) {
    targetH = vh;
    targetW = targetH * ratio;
  }
  slEl.style.width = Math.round(targetW) + 'px';
  slEl.style.height = Math.round(targetH) + 'px';

  const xmlShapes = shapesOf(pk.doc(slPath));
  let delay = 0;
  $$('.shape:not(.static)', slEl).forEach((el) => {
    const idx = +(el.dataset.i || -1);
    const anim = xmlShapes[idx] && kid(nvOf(xmlShapes[idx]), P, 'cNvPr')?.getAttribute('data-anim');
    if (anim && anim !== 'none') {
      el.style.opacity = '0';
      el.style.animationDelay = (delay * 0.35) + 's';
      el.classList.add('anim-preview-' + anim);
      delay++;
    }
  });

  st.append(slEl);
  $('#ssCounter').textContent = `${pos + 1} / ${visibleIndices.length}`;
  if (ssState.pen) setupSlideshowPenCanvas();
  curPptx.slideIdx = realIdx;
  syncPresenterWindow();
}

function exitSlideshow() {
  stopSlideshowAutoplay();
  const ov = $('#slideshowOverlay');
  if (ov) {
    ov.hidden = true;
    ov.style.display = '';
  }
  $('#laserDot').hidden = true;
  if ($('#ssPenCanvas')) $('#ssPenCanvas').hidden = true;
  if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
}

window.addEventListener('resize', () => {
  if (!$('#slideshowOverlay')?.hidden) renderSlideshowSlide();
});

$('#ssPrev')?.addEventListener('click', () => {
  if (ssState.pos > 0) { ssState.pos--; renderSlideshowSlide(); }
});
$('#ssNext')?.addEventListener('click', () => {
  if (ssState.pos < ssState.visibleIndices.length - 1) { ssState.pos++; renderSlideshowSlide(); }
});
$('#ssAutoBtn')?.addEventListener('click', () => {
  if (ssState.autoTimer) {
    stopSlideshowAutoplay();
  } else {
    $('#ssAutoBtn')?.classList.add('on');
    ssState.autoTimer = setInterval(() => {
      if ($('#slideshowOverlay')?.hidden || !ssState.visibleIndices.length) {
        stopSlideshowAutoplay();
        return;
      }
      ssState.pos = (ssState.pos + 1) % ssState.visibleIndices.length;
      renderSlideshowSlide();
    }, 4500);
  }
});
$('#ssExit')?.addEventListener('click', exitSlideshow);
$('#ssLaserBtn')?.addEventListener('click', () => {
  ssState.laser = !ssState.laser;
  if (ssState.laser) { ssState.pen = false; $('#ssPenCanvas').hidden = true; $('#ssClearPenBtn').hidden = true; $('#ssPenBtn')?.classList.remove('on'); }
  $('#laserDot').hidden = !ssState.laser;
  $('#ssLaserBtn').classList.toggle('on', ssState.laser);
});
$('#ssPenBtn')?.addEventListener('click', () => {
  ssState.pen = !ssState.pen;
  const cvs = $('#ssPenCanvas');
  if (ssState.pen) {
    ssState.laser = false;
    $('#laserDot').hidden = true;
    $('#ssLaserBtn')?.classList.remove('on');
    if (cvs) { cvs.hidden = false; setupSlideshowPenCanvas(); }
    if ($('#ssClearPenBtn')) $('#ssClearPenBtn').hidden = false;
  } else {
    if (cvs) cvs.hidden = true;
    if ($('#ssClearPenBtn')) $('#ssClearPenBtn').hidden = true;
  }
  $('#ssPenBtn')?.classList.toggle('on', ssState.pen);
});
$('#ssClearPenBtn')?.addEventListener('click', setupSlideshowPenCanvas);

const penCvs = $('#ssPenCanvas');
if (penCvs) {
  penCvs.addEventListener('pointerdown', (e) => {
    if (!ssState.pen) return;
    ssState.drawing = true;
    const ctx = penCvs.getContext('2d');
    ctx.beginPath();
    ctx.moveTo(e.clientX, e.clientY);
  });
  penCvs.addEventListener('pointermove', (e) => {
    if (!ssState.pen || !ssState.drawing) return;
    const ctx = penCvs.getContext('2d');
    ctx.lineTo(e.clientX, e.clientY);
    ctx.stroke();
  });
  const stopDraw = () => { ssState.drawing = false; };
  penCvs.addEventListener('pointerup', stopDraw);
  penCvs.addEventListener('pointercancel', stopDraw);
}

$('#slideshowOverlay')?.addEventListener('pointermove', (e) => {
  if (!ssState.laser) return;
  const dot = $('#laserDot');
  dot.style.left = e.clientX + 'px';
  dot.style.top = e.clientY + 'px';
});

(() => {
  const st = $('#ssStage');
  if (!st) return;
  let sx = 0, sy = 0, tracking = false;
  st.addEventListener('pointerdown', (e) => {
    if (ssState.pen || $('#slideshowOverlay')?.hidden) return;
    sx = e.clientX; sy = e.clientY; tracking = true;
  });
  st.addEventListener('pointerup', (e) => {
    if (!tracking || ssState.pen) return;
    tracking = false;
    const dx = e.clientX - sx;
    const dy = e.clientY - sy;
    if (Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(dy) * 1.25) {
      if (dx < 0 && ssState.pos < ssState.visibleIndices.length - 1) {
        ssState.pos++; renderSlideshowSlide();
      } else if (dx > 0 && ssState.pos > 0) {
        ssState.pos--; renderSlideshowSlide();
      }
    }
  });
  st.addEventListener('pointercancel', () => { tracking = false; });
})();
document.addEventListener('keydown', (e) => {
  if ($('#slideshowOverlay').hidden) return;
  if (e.key === 'Escape') exitSlideshow();
  else if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
    if (ssState.pos < ssState.visibleIndices.length - 1) { ssState.pos++; renderSlideshowSlide(); }
  } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
    if (ssState.pos > 0) { ssState.pos--; renderSlideshowSlide(); }
  }
});

/* ── تحويل روابط blob: إلى Base64 وتصدير PNG / ZIP / PDF ── */
async function blobUrlToDataUrl(blobUrl) {
  const res = await fetch(blobUrl);
  const blob = await res.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function inlineBlobAssets(clone) {
  const cache = new Map();
  const convert = async (u) => {
    if (!u || !u.startsWith('blob:')) return u;
    if (cache.has(u)) return cache.get(u);
    const dataUrl = await blobUrlToDataUrl(u);
    cache.set(u, dataUrl);
    return dataUrl;
  };

  for (const img of $$('img', clone)) {
    const src = img.getAttribute('src') || img.src || '';
    if (src.startsWith('blob:')) {
      try { img.setAttribute('src', await convert(src)); } catch (_) {}
    }
  }

  for (const el of [clone, ...$$('*', clone)]) {
    const bgImg = el.style && el.style.backgroundImage;
    if (bgImg && bgImg.includes('blob:')) {
      const m = /url\(["']?(blob:[^"')]+)["']?\)/.exec(bgImg);
      if (m) {
        try {
          const dataUrl = await convert(m[1]);
          el.style.backgroundImage = bgImg.replace(m[1], dataUrl);
        } catch (_) {}
      }
    }
  }
}

async function slideToCanvas(deck, idx) {
  const slEl = renderSlideEl(deck, idx, false);
  slEl.classList.remove('show-grid');
  const w = 1280, hPx = Math.round(1280 * (deck.H / deck.W));
  slEl.style.width = w + 'px';
  slEl.style.height = hPx + 'px';
  const holder = h('div', { style: { position: 'fixed', left: '-9999px', top: '0', width: w + 'px', height: hPx + 'px' } }, slEl);
  document.body.appendChild(holder);

  const clone = slEl.cloneNode(true);
  clone.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
  await inlineBlobAssets(clone);

  const xml = new XMLSerializer().serializeToString(clone);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hPx}"><foreignObject width="100%" height="100%">${xml}</foreignObject></svg>`;
  const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);

  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = hPx;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, hPx);
    ctx.drawImage(img, 0, 0);
    return canvas;
  } finally {
    holder.remove();
  }
}

async function exportCurrentSlidePng(deck, idx) {
  busy(true, 'جارٍ تصدير الشريحة كصورة PNG…');
  try {
    const canvas = await slideToCanvas(deck, idx);
    const a = document.createElement('a');
    a.href = canvas.toDataURL('image/png');
    a.download = `slide-${idx + 1}.png`;
    a.click();
    toast('تم تصدير الشريحة كـ PNG');
  } catch (_) { toast('تعذر تصدير الصورة'); }
  finally { busy(false); }
}

async function exportAllSlidesZip(deck) {
  if (!deck || !deck.slides.length) return;
  busy(true, 'جارٍ تصدير جميع الشرائح في ملف ZIP…');
  try {
    const zip = new JSZip();
    const total = deck.slides.length;
    for (let i = 0; i < total; i++) {
      busy(true, `جارٍ تصدير الشريحة ${i + 1} من ${total}…`);
      const canvas = await slideToCanvas(deck, i);
      const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
      if (blob) zip.file(`slide-${i + 1}.png`, blob);
    }
    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    const base = fileName.replace(/\.(pptx|pptm)$/i, '') || 'slides';
    a.href = url;
    a.download = `${base}-png-slides.zip`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast('تم تصدير جميع الشرائح كملف ZIP');
  } catch (_) {
    toast('تعذر تصدير ملف ZIP');
  } finally {
    busy(false);
  }
}

function openInlinePresenterModal(deck) {
  $('#inlinePresenterBackdrop')?.remove();
  let sec = 0;
  const fmtTime = (s) => [Math.floor(s / 3600), Math.floor((s % 3600) / 60), s % 60].map(n => String(n).padStart(2, '0')).join(':');
  const timerEl = h('span', { style: { fontFamily: 'monospace', fontSize: '1.15rem', fontWeight: '800', color: 'var(--acc)', background: 'var(--soft)', border: '1px solid var(--line)', padding: '3px 10px', borderRadius: '8px' } }, '00:00:00');
  const counterEl = h('strong', { style: { minWidth: '52px', textAlign: 'center' } }, '');
  const curStage = h('div', { style: { background: 'var(--stage-bg)', borderRadius: '10px', padding: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '260px' } });
  const nextStage = h('div', { style: { background: 'var(--stage-bg)', borderRadius: '10px', padding: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '140px' } });
  const notesBox = h('div', { style: { flex: '1', background: 'var(--soft)', border: '1px solid var(--line)', borderRadius: '8px', padding: '10px', minHeight: '90px', whiteSpace: 'pre-wrap', overflow: 'auto', fontSize: '.92rem' } });

  let prevBtn, nextBtn;
  const syncModal = () => {
    const total = deck.slides.length;
    const idx = Math.max(0, Math.min(total - 1, curPptx.slideIdx || 0));
    counterEl.textContent = `${idx + 1} / ${total}`;
    if (prevBtn) prevBtn.disabled = idx <= 0;
    if (nextBtn) nextBtn.disabled = idx >= total - 1;
    curStage.innerHTML = '';
    const curEl = renderSlideEl(deck, idx, false);
    curEl.style.maxWidth = '100%';
    curStage.append(curEl);
    nextStage.innerHTML = '';
    if (idx + 1 < total) {
      const nxtEl = renderSlideEl(deck, idx + 1, false);
      nxtEl.style.maxWidth = '100%';
      nextStage.append(nxtEl);
    } else {
      nextStage.append(h('span', { class: 'muted' }, currentLanguage === 'en' ? 'End of presentation' : 'نهاية العرض'));
    }
    const notes = getSlideNotes(deck.slides[idx].path);
    notesBox.textContent = notes || (currentLanguage === 'en' ? 'No speaker notes for this slide.' : 'لا توجد ملاحظات لهذه الشريحة.');
  };

  let timerId = null;
  const close = () => { if (timerId) clearInterval(timerId); backdrop.remove(); renderPptxEditor(); };
  prevBtn = btn('◀ السابق', () => { if (curPptx.slideIdx > 0) { curPptx.slideIdx--; syncModal(); } });
  nextBtn = btn('التالي ▶', () => { if (curPptx.slideIdx < deck.slides.length - 1) { curPptx.slideIdx++; syncModal(); } }, 'pri');

  const backdrop = h('div', { id: 'inlinePresenterBackdrop', class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) close(); } },
    h('div', { class: 'md-modal', style: { width: 'min(980px, 96vw)', maxHeight: '92vh', overflow: 'auto' } },
      h('div', { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap', borderBottom: '1px solid var(--line)', paddingBottom: '8px' } },
        h('div', { style: { display: 'flex', alignItems: 'center', gap: '10px' } },
          h('strong', null, tr('🖥 وضع مقدم العرض')),
          timerEl,
          btn(currentLanguage === 'en' ? 'Reset' : 'تصفير', () => { sec = 0; timerEl.textContent = '00:00:00'; })
        ),
        h('div', { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
          prevBtn, counterEl, nextBtn,
          btn(currentLanguage === 'en' ? '↗ Separate Window' : '↗ نافذة مستقلة', () => { close(); openPopupPresenterWindow(deck, curPptx.slideIdx); }),
          btn('إغلاق', close, 'danger')
        )
      ),
      h('div', { class: 'pv-inline-grid' },
        h('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } },
          h('small', { class: 'muted' }, currentLanguage === 'en' ? 'Current Slide' : 'الشريحة الحالية'),
          curStage
        ),
        h('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
          h('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } },
            h('small', { class: 'muted' }, currentLanguage === 'en' ? 'Next Slide' : 'الشريحة القادمة'),
            nextStage
          ),
          h('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } },
            h('small', { class: 'muted' }, tr('📝 ملاحظات المحاضر (Speaker Notes)')),
            notesBox
          )
        )
      )
    )
  );
  document.body.append(backdrop);
  syncModal();
  timerId = setInterval(() => { sec++; timerEl.textContent = fmtTime(sec); }, 1000);
}

async function exportDeckToPdf(deck) {
  const visibleSlides = deck.slides.filter(s => !s.hidden);
  if (!visibleSlides.length) { toast('جميع الشرائح مخفية!'); return; }

  busy(true, 'جارٍ تجهيز ملف PDF…');
  try {
    const { PDFDocument } = getPdfLib();
    const pdfDoc = await PDFDocument.create();
    const pageWPt = Math.max(72, Math.round((deck.W / 914400) * 72));
    const pageHPt = Math.max(72, Math.round((deck.H / 914400) * 72));

    for (let i = 0; i < visibleSlides.length; i++) {
      busy(true, `جارٍ تصدير الشريحة ${i + 1} من ${visibleSlides.length}…`);
      const realIdx = deck.slides.indexOf(visibleSlides[i]);
      const canvas = await slideToCanvas(deck, realIdx);
      const pngBytes = await new Promise((resolve, reject) => {
        canvas.toBlob(async (blob) => {
          if (!blob) return reject(new Error('PNG blob failed'));
          resolve(new Uint8Array(await blob.arrayBuffer()));
        }, 'image/png');
      });
      const pngImg = await pdfDoc.embedPng(pngBytes);
      const page = pdfDoc.addPage([pageWPt, pageHPt]);
      page.drawImage(pngImg, { x: 0, y: 0, width: pageWPt, height: pageHPt });
    }

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const base = (fileName || 'presentation').replace(/\.(pptx|pptm)$/i, '') || 'presentation';
    a.href = url;
    a.download = `${base}.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast('تم تصدير ملف PDF بنجاح');
  } catch (err) {
    console.error('Deck PDF export failed:', err);
    toast('حدث خطأ أثناء تصدير الملف');
  } finally {
    busy(false);
  }
}
