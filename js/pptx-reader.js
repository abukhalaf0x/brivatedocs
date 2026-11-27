'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 3 من 6: محرك قراءة ورسم الشرائح، الجداول، والرسوم البيانية
   ═════════════════════════════════════════════════════════════ */

let pk = null, kind = null, fileName = '';

const SHAPE_TAGS = new Set(['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp']);
function spTreeOf(doc) { return kid(kid(doc.documentElement, P, 'cSld'), P, 'spTree'); }
function shapesOf(doc) { return kids(spTreeOf(doc), P).filter(e => SHAPE_TAGS.has(e.localName)); }
function nvOf(sh) {
  for (const n of ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr']) { const e = kid(sh, P, n); if (e) return e; }
  return null;
}
function phOf(sh) {
  const ph = kid(kid(nvOf(sh), P, 'nvPr'), P, 'ph');
  return ph ? { type: ph.getAttribute('type') || 'obj', idx: ph.getAttribute('idx') } : null;
}
const txBodyOf = sh => kid(sh, P, 'txBody') || kid(sh, A, 'txBody');
function normType(t) { return t === 'ctrTitle' ? 'title' : (t === 'subTitle' || t === 'obj') ? 'body' : t; }
function findPlaceholder(doc, ph, isMaster) {
  const list = shapesOf(doc).map(s => [s, phOf(s)]).filter(x => x[1]);
  if (isMaster) { const f = list.find(x => normType(x[1].type) === normType(ph.type)); return f ? f[0] : null; }
  let f = ph.idx != null ? list.find(x => x[1].idx === ph.idx) : null;
  if (!f) f = list.find(x => x[1].type === ph.type && ph.type !== 'obj');
  return f ? f[0] : null;
}

function readDeck() {
  const pres = pk.mainPart, root = pk.doc(pres).documentElement;
  const sz = kid(root, P, 'sldSz');
  const rels = readRels(pk, pres);
  const slides = [];
  for (const e of kids(kid(root, P, 'sldIdLst'), P, 'sldId')) {
    const rid = e.getAttributeNS(R, 'id'), rel = rels.get(rid);
    if (rel && rel.path && pk.has(rel.path)) {
      slides.push({ el: e, id: e.getAttribute('id'), rid, path: rel.path, hidden: e.getAttribute('show') === '0' });
    }
  }
  return { pres, root, W: (sz && +sz.getAttribute('cx')) || 9144000, H: (sz && +sz.getAttribute('cy')) || 6858000, slides, rels };
}

function buildCtx(deck, slidePath) {
  const layer = path => (path && pk.has(path)) ? { path, doc: pk.doc(path), rels: readRels(pk, path) } : null;
  const doc = pk.doc(slidePath), rels = readRels(pk, slidePath);
  const lrel = relByType(rels, 'slideLayout'), layout = layer(lrel && lrel.path);
  const mrel = layout && relByType(layout.rels, 'slideMaster'), master = layer(mrel && mrel.path);
  const trel = master && relByType(master.rels, 'theme');
  const themeDoc = trel && pk.has(trel.path) ? pk.doc(trel.path) : null;
  const theme = {};
  if (themeDoc) {
    const cs = desc(themeDoc.documentElement, A, 'clrScheme')[0];
    for (const c of kids(cs)) {
      const ce = kids(c)[0]; if (!ce) continue;
      theme[c.localName] = ce.localName === 'srgbClr' ? ce.getAttribute('val') : (ce.getAttribute('lastClr') || '000000');
    }
  }
  const clrMap = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  const cm = master && kid(master.doc.documentElement, P, 'clrMap');
  if (cm) for (const at of Array.from(cm.attributes)) clrMap[at.name] = at.value;
  for (const d of [layout && layout.doc, doc]) {
    const ov = d && kid(kid(d.documentElement, P, 'clrMapOvr'), A, 'overrideClrMapping');
    if (ov) for (const at of Array.from(ov.attributes)) clrMap[at.name] = at.value;
  }
  return {
    deck, slide: { path: slidePath, doc, rels }, layout, master, theme, clrMap, wpt: deck.W / 12700,
    defTextStyle: kid(pk.doc(deck.pres).documentElement, P, 'defaultTextStyle')
  };
}

function xfrmOf(sh) {
  const ln = sh.localName;
  if (ln === 'graphicFrame') return kid(sh, P, 'xfrm');
  if (ln === 'grpSp') return kid(kid(sh, P, 'grpSpPr'), A, 'xfrm');
  return kid(kid(sh, P, 'spPr'), A, 'xfrm');
}
function readXfrm(x) {
  if (!x) return null;
  const off = kid(x, A, 'off'), ext = kid(x, A, 'ext');
  if (!off || !ext) return null;
  const g = {
    x: +off.getAttribute('x') || 0, y: +off.getAttribute('y') || 0,
    w: +ext.getAttribute('cx') || 0, h: +ext.getAttribute('cy') || 0,
    rot: (+x.getAttribute('rot') || 0) / 60000,
    flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1'
  };
  const co = kid(x, A, 'chOff'), ce = kid(x, A, 'chExt');
  if (co && ce) { g.chx = +co.getAttribute('x') || 0; g.chy = +co.getAttribute('y') || 0; g.chw = +ce.getAttribute('cx') || g.w || 1; g.chh = +ce.getAttribute('cy') || g.h || 1; }
  return g;
}
function geomOf(sh, ctx) {
  let g = readXfrm(xfrmOf(sh));
  if (g) return g;
  const ph = phOf(sh);
  if (ph) for (const [L, m] of [[ctx.layout, false], [ctx.master, true]]) {
    if (!L) continue;
    const s = findPlaceholder(L.doc, ph, m), gg = s && readXfrm(xfrmOf(s));
    if (gg) return gg;
  }
  return null;
}

function chainsFor(ctx, sh, lvl) {
  const key = 'lvl' + (lvl + 1) + 'pPr', ph = phOf(sh), hi = [], lo = [];
  const lv = lst => kid(lst, A, key);
  hi.push(lv(kid(txBodyOf(sh), A, 'lstStyle')));
  if (ph) for (const [L, m] of [[ctx.layout, false], [ctx.master, true]]) {
    if (!L) continue;
    const s = findPlaceholder(L.doc, ph, m);
    if (s) hi.push(lv(kid(txBodyOf(s), A, 'lstStyle')));
  }
  const tx = ctx.master && kid(ctx.master.doc.documentElement, P, 'txStyles');
  const nm = ph ? (/^(title|ctrTitle)$/.test(ph.type) ? 'titleStyle' : /^(dt|ftr|sldNum)$/.test(ph.type) ? 'otherStyle' : 'bodyStyle') : 'otherStyle';
  lo.push(lv(kid(tx, P, nm)), lv(ctx.defTextStyle));
  return { hi: hi.filter(Boolean), lo: lo.filter(Boolean) };
}

function runProps(ctx, sh, rPr, lvl) {
  const ch = chainsFor(ctx, sh, lvl);
  const hiR = [rPr, ...ch.hi.map(c => kid(c, A, 'defRPr'))].filter(Boolean);
  const loR = ch.lo.map(c => kid(c, A, 'defRPr')).filter(Boolean);
  const all = [...hiR, ...loR];
  const get = n => { for (const r of all) if (r.hasAttribute(n)) return r.getAttribute(n); return null; };
  let color = null;
  for (const r of hiR) { const f = fillFrom(r, ctx); if (f && f.type === 'solid') { color = f; break; } }
  if (!color) {
    const fr = kid(kid(sh, P, 'style'), A, 'fontRef'), ce = fr && colorChild(fr);
    const col = ce && parseColorEl(ce, ctx); if (col) color = { hex: col.hex, alpha: col.alpha };
  }
  if (!color) for (const r of loR) { const f = fillFrom(r, ctx); if (f && f.type === 'solid') { color = f; break; } }
  if (!color) { const hx = schemeHex('tx1', ctx); color = { hex: hx || '000000', alpha: 1 }; }
  let font = null;
  for (const r of all) {
    const lat = kid(r, A, 'latin'), ea = kid(r, A, 'ea');
    if (lat && lat.getAttribute('typeface')) { font = lat.getAttribute('typeface'); break; }
    if (ea && ea.getAttribute('typeface')) { font = ea.getAttribute('typeface'); break; }
  }
  return {
    sz: +get('sz') || 1800, b: get('b') === '1', i: get('i') === '1',
    u: (get('u') || 'none') !== 'none', strike: (get('strike') || 'noStrike') !== 'noStrike',
    baseline: +get('baseline') || 0, cap: get('cap') === 'all', color, font
  };
}

let FS = 1, CUR_WPT = 720;
const fsCss = sz => `${(sz / 100) * FS / CUR_WPT * 100}cqw`;

function renderParagraph(ctx, sh, p, state) {
  const pPr = kid(p, A, 'pPr'), lvl = Math.min(8, +(pPr && pPr.getAttribute('lvl')) || 0);
  const ch = chainsFor(ctx, sh, lvl), pl = [pPr, ...ch.hi, ...ch.lo].filter(Boolean);
  const pget = n => { for (const e of pl) if (e.hasAttribute(n)) return e.getAttribute(n); return null; };
  const pchild = n => { for (const e of pl) { const c = kid(e, A, n); if (c) return c; } return null; };
  const div = h('div', { class: 'tp' });
  const rtl = pget('rtl');
  if (rtl !== null) { div.style.direction = rtl === '1' ? 'rtl' : 'ltr'; div.style.unicodeBidi = 'isolate'; }
  const algn = pget('algn');
  div.style.textAlign = algn ? ({ l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'start') : 'start';

  let first = null, any = false;
  for (const n of kids(p, A)) {
    if (n.localName === 'r' || n.localName === 'fld') {
      const rp = runProps(ctx, sh, kid(n, A, 'rPr'), lvl), t = kid(n, A, 't');
      const sp = h('span');
      sp.textContent = t ? t.textContent : '';
      sp.style.fontSize = fsCss(rp.sz * (rp.baseline ? 0.65 : 1));
      sp.style.color = cssColor(rp.color);
      if (rp.font && !rp.font.startsWith('+')) sp.style.fontFamily = `"${rp.font}", sans-serif`;
      if (rp.b) sp.style.fontWeight = '700';
      if (rp.i) sp.style.fontStyle = 'italic';
      const td = []; if (rp.u) td.push('underline'); if (rp.strike) td.push('line-through');
      if (td.length) sp.style.textDecoration = td.join(' ');
      if (rp.baseline) sp.style.verticalAlign = rp.baseline > 0 ? 'super' : 'sub';
      if (rp.cap) sp.style.textTransform = 'uppercase';
      div.append(sp); first = first || rp; any = true;
    } else if (n.localName === 'br') { div.append(h('br')); }
  }
  const end = kid(p, A, 'endParaRPr');
  if (!first) first = runProps(ctx, sh, end, lvl);
  div.style.fontSize = fsCss(first.sz);
  if (!any) { const sp = h('span'); sp.textContent = '\u200b'; sp.style.fontSize = fsCss(first.sz); div.append(sp); }

  const lnSpc = pchild('lnSpc'), sp1 = lnSpc && kid(lnSpc, A, 'spcPct');
  div.style.lineHeight = sp1 ? String(Math.max(.6, (+sp1.getAttribute('val') / 100000) * 1.2)) : '1.2';
  const spB = pchild('spcBef'), spA = pchild('spcAft');
  const spc = e => {
    if (!e) return null;
    const pts = kid(e, A, 'spcPts'), pct = kid(e, A, 'spcPct');
    if (pts) return `${(+pts.getAttribute('val') / 100) / CUR_WPT * 100}cqw`;
    if (pct) return `${(+pct.getAttribute('val') / 100000) * 1.2}em`;
    return null;
  };
  if (spc(spB)) div.style.marginTop = spc(spB);
  if (spc(spA)) div.style.marginBottom = spc(spA);

  const marL = +pget('marL') || 0, indent = +pget('indent') || 0;
  let bu = null;
  for (const e of pl) {
    if (kid(e, A, 'buNone')) { bu = null; break; }
    const bc = kid(e, A, 'buChar'), ba = kid(e, A, 'buAutoNum');
    if (bc) { const f = kid(e, A, 'buFont'), tf = f ? (f.getAttribute('typeface') || '') : ''; bu = { char: /Wingdings|Symbol/i.test(tf) ? '•' : (bc.getAttribute('char') || '•') }; break; }
    if (ba) { bu = { auto: true, start: +ba.getAttribute('startAt') || 1, type: ba.getAttribute('type') }; break; }
  }
  if (bu && any) {
    let txt = bu.char;
    if (bu.auto) {
      state.counters[lvl] = (state.counters[lvl] || 0) + 1;
      const n = state.counters[lvl] + bu.start - 1;
      txt = (/alpha/i.test(bu.type || '') ? String.fromCharCode(96 + ((n - 1) % 26) + 1) : String(n)) + '.';
    }
    const sp = h('span', { class: 'bu' }, txt + ' ');
    sp.style.fontSize = fsCss(first.sz); sp.style.color = cssColor(first.color);
    div.prepend(sp);
    if (!bu.auto) state.counters[lvl] = 0;
  } else if (!bu) state.counters[lvl] = 0;
  if (marL || indent) {
    const cq = v => `${v / (CUR_WPT * 12700) * 100}cqw`;
    div.style.paddingInlineStart = cq(marL); div.style.textIndent = cq(indent);
  }
  return div;
}

function renderTextBody(ctx, sh, txb) {
  const bodyPr = kid(txb, A, 'bodyPr');
  const anchor = (bodyPr && bodyPr.getAttribute('anchor')) || 't';
  const ins = (n, d) => { const v = bodyPr && bodyPr.getAttribute(n); return (v != null && v !== '' ? +v : d) / ctx.deck.W * 100; };
  const box = h('div', { class: 'tb' });
  box.style.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
  box.style.padding = `${ins('tIns', 45720)}cqw ${ins('rIns', 91440)}cqw ${ins('bIns', 45720)}cqw ${ins('lIns', 91440)}cqw`;
  const na = kid(bodyPr, A, 'normAutofit');
  const prev = FS; FS = na && na.getAttribute('fontScale') ? +na.getAttribute('fontScale') / 100000 : 1;
  const state = { counters: [] };
  for (const p of kids(txb, A, 'p')) box.append(renderParagraph(ctx, sh, p, state));
  FS = prev;
  return box;
}

function lineSpec(sh, ctx) {
  const spPr = kid(sh, P, 'spPr'), ln = kid(spPr, A, 'ln');
  if (ln) {
    if (kid(ln, A, 'noFill')) return null;
    const f = fillFrom(ln, ctx);
    if (f && f.type === 'solid') return { w: +ln.getAttribute('w') || 12700, color: f };
  }
  const lr = kid(kid(sh, P, 'style'), A, 'lnRef');
  if (lr && lr.getAttribute('idx') !== '0' && !(ln && kid(ln, A, 'noFill'))) {
    const ce = colorChild(lr), col = ce && parseColorEl(ce, ctx);
    if (col) return { w: (ln && +ln.getAttribute('w')) || 12700, color: { hex: col.hex, alpha: col.alpha } };
  }
  return null;
}

function shapeFill(sh, ctx) {
  const f = fillFrom(kid(sh, P, 'spPr'), ctx);
  if (f) return f;
  const fr = kid(kid(sh, P, 'style'), A, 'fillRef');
  if (fr && fr.getAttribute('idx') !== '0') {
    const ce = colorChild(fr), col = ce && parseColorEl(ce, ctx);
    if (col) return { type: 'solid', hex: col.hex, alpha: col.alpha };
  }
  return null;
}

function readShapeAppearance(sh, ctx) {
  const fill = shapeFill(sh, ctx);
  const ln = lineSpec(sh, ctx);
  const spPr = kid(sh, P, 'spPr');
  const eff = spPr && kid(spPr, A, 'effectLst');
  const hasShadow = !!(eff && kid(eff, A, 'outerShdw'));
  let fillType = 'solid', fillHex = '#F5F2EC', gradHex2 = '#B4531A', gradAngle = 90, noFill = false, opacity = 100;

  if (!fill || fill.type === 'none') {
    noFill = true;
  } else if (fill.type === 'solid') {
    fillType = 'solid';
    fillHex = '#' + fill.hex;
    if (fill.alpha != null) opacity = Math.round(fill.alpha * 100);
  } else if (fill.type === 'grad' && fill.rawStops?.length) {
    fillType = 'grad';
    fillHex = '#' + fill.rawStops[0].hex;
    gradHex2 = '#' + (fill.rawStops[fill.rawStops.length - 1]?.hex || fill.rawStops[0].hex);
    gradAngle = fill.angle || 0;
    if (fill.rawStops[0].alpha != null) opacity = Math.round(fill.rawStops[0].alpha * 100);
  }

  return {
    fillType,
    fillHex,
    gradHex2,
    gradAngle,
    noFill,
    opacity: Math.max(10, Math.min(100, opacity)),
    borderHex: ln ? '#' + ln.color.hex : '#1F1C17',
    borderW: ln ? Math.max(1, Math.round(ln.w / 12700)) : 0,
    hasShadow
  };
}

function getShapePlainText(sh) {
  const txb = txBodyOf(sh);
  if (!txb) return '';
  return kids(txb, A, 'p').map(p => desc(p, A, 't').map(t => t.textContent).join('')).join('\n');
}

function setShapePlainText(sh, text) {
  const txb = txBodyOf(sh);
  if (!txb) return;
  const doc = sh.ownerDocument;
  const existingPs = kids(txb, A, 'p');
  const basePPr = existingPs[0] && kid(existingPs[0], A, 'pPr')?.cloneNode(true);
  const baseRPr = desc(txb, A, 'rPr')[0]?.cloneNode(true);
  for (const p of existingPs) removeNode(p);
  const lines = String(text ?? '').split(/\r?\n/);
  for (const line of (lines.length ? lines : [''])) {
    const p = mk(doc, A, 'a:p');
    if (basePPr) p.appendChild(basePPr.cloneNode(true));
    const r = mk(doc, A, 'a:r');
    if (baseRPr) r.appendChild(baseRPr.cloneNode(true));
    const t = mk(doc, A, 'a:t');
    if (/^\s|\s$/.test(line)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = line;
    r.appendChild(t);
    p.appendChild(r);
    txb.appendChild(p);
  }
}

function setCellTextXml(tc, text) {
  if (!tc) return;
  const doc = tc.ownerDocument;
  let txb = txBodyOf(tc);
  if (!txb) {
    txb = mk(doc, A, 'a:txBody');
    txb.append(mk(doc, A, 'a:bodyPr'), mk(doc, A, 'a:lstStyle'), mk(doc, A, 'a:p'));
    tc.appendChild(txb);
  }
  const existingPs = kids(txb, A, 'p');
  const basePPr = existingPs[0] && kid(existingPs[0], A, 'pPr')?.cloneNode(true);
  const baseRPr = desc(txb, A, 'rPr')[0]?.cloneNode(true);
  for (const p of existingPs) removeNode(p);
  const lines = String(text ?? '').replace(/\u200b/g, '').split(/\r?\n/);
  for (const line of (lines.length ? lines : [''])) {
    const p = mk(doc, A, 'a:p');
    if (basePPr) p.appendChild(basePPr.cloneNode(true));
    const r = mk(doc, A, 'a:r');
    if (baseRPr) r.appendChild(baseRPr.cloneNode(true));
    const t = mk(doc, A, 'a:t');
    if (/^\s|\s$/.test(line)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = line;
    r.appendChild(t);
    p.appendChild(r);
    txb.appendChild(p);
  }
}

const lenCss = (emu, W0) => `${emu / W0 * 100}cqw`;

function placeEl(el, g, sp) {
  el.style.left = ((g.x - sp.ox) / sp.w * 100) + '%';
  el.style.top = ((g.y - sp.oy) / sp.h * 100) + '%';
  el.style.width = (g.w / sp.w * 100) + '%';
  el.style.height = (g.h / sp.h * 100) + '%';
}

/* رسم الصور والفيديو والصوت والقص الدائري */
function renderPic(el, sh, L, ctx, g) {
  el.dataset.kind = 'pic';
  const spPr = kid(sh, P, 'spPr');
  const prst = kid(spPr, A, 'prstGeom')?.getAttribute('prst');
  if (prst === 'ellipse') el.style.borderRadius = '50%';

  const blip = desc(kid(sh, P, 'blipFill'), A, 'blip')[0];
  const rid = blip && blip.getAttributeNS(R, 'embed'), rel = rid && L.rels.get(rid);
  const u8 = rel && rel.path && pk.files.get(rel.path), url = u8 && mediaURL(u8, rel.path);
  const nvPr = kid(nvOf(sh), P, 'nvPr');
  const vidFile = nvPr && (kid(nvPr, A, 'videoFile') || kid(nvPr, A, 'audioFile'));
  const mRid = vidFile && vidFile.getAttributeNS(R, 'link');
  const mRel = mRid && L.rels.get(mRid);
  const mU8 = mRel && mRel.path && pk.files.get(mRel.path);
  const mUrl = mU8 && mediaURL(mU8, mRel.path);

  const crop = h('div', { class: 'crop' });
  if (prst === 'ellipse') crop.style.borderRadius = '50%';

  if (mUrl) {
    el.classList.add('media-shape');
    const isAudio = vidFile.localName === 'audioFile';
    crop.append(isAudio ? h('audio', { src: mUrl, controls: true, style: { width: '100%' } }) : h('video', { src: mUrl, controls: true, poster: url || '' }));
  } else if (url) {
    const img = h('img', { src: url, draggable: 'false', alt: '' });
    const sr = kid(kid(sh, P, 'blipFill'), A, 'srcRect');
    const l = (sr && +sr.getAttribute('l') || 0) / 100000, t = (sr && +sr.getAttribute('t') || 0) / 100000, r = (sr && +sr.getAttribute('r') || 0) / 100000, b = (sr && +sr.getAttribute('b') || 0) / 100000;
    const kw = Math.max(.01, 1 - l - r), kh = Math.max(.01, 1 - t - b);
    img.style.width = (100 / kw) + '%'; img.style.height = (100 / kh) + '%';
    img.style.left = (-l / kw * 100) + '%'; img.style.top = (-t / kh * 100) + '%';
    if (g.flipH || g.flipV) img.style.transform = `scale(${g.flipH ? -1 : 1},${g.flipV ? -1 : 1})`;
    crop.append(img);
  } else {
    el.classList.add('ph-box');
    crop.textContent = '🖼';
    crop.style.display = 'flex'; crop.style.alignItems = 'center'; crop.style.justifyContent = 'center';
  }
  el.append(crop);
  const ln = lineSpec(sh, ctx);
  if (ln) el.style.border = `max(1px,${lenCss(ln.w, ctx.deck.W)}) solid ${cssColor(ln.color)}`;
}

/* قراءة ورسم الرسوم البيانية متعددة السلاسل (Bar / Line / Pie) بـ SVG */
function parseChartData(chartPath) {
  if (!chartPath || !pk.has(chartPath)) return null;
  const doc = pk.doc(chartPath);
  const plot = desc(doc.documentElement, C, 'plotArea')[0];
  if (!plot) return null;
  let type = 'bar';
  let chartNode = kid(plot, C, 'barChart');
  if (!chartNode) { chartNode = kid(plot, C, 'lineChart'); if (chartNode) type = 'line'; }
  if (!chartNode) { chartNode = kid(plot, C, 'pieChart'); if (chartNode) type = 'pie'; }
  if (!chartNode) return null;

  const sers = kids(chartNode, C, 'ser');
  if (!sers.length) return null;
  const cats = desc(kid(sers[0], C, 'cat'), C, 'v').map(v => v.textContent);
  const series = sers.map((s, idx) => {
    const nameEl = desc(kid(s, C, 'tx'), C, 'v')[0];
    const name = nameEl ? nameEl.textContent : (currentLanguage === 'en' ? `Series ${idx + 1}` : `السلسلة ${idx + 1}`);
    const vals = desc(kid(s, C, 'val'), C, 'v').map(v => +v.textContent || 0);
    return { name, vals };
  });
  return { type, cats, vals: series[0].vals, series, chartPath };
}

function renderChart(el, gf, L, ctx) {
  el.dataset.kind = 'chart';
  const cEl = desc(gf, C, 'chart')[0];
  const rid = cEl && cEl.getAttributeNS(R, 'id');
  const rel = rid && L.rels.get(rid);
  const data = rel && parseChartData(rel.path);

  if (!data || !data.series.length || !data.cats.length) {
    el.classList.add('ph-box');
    el.textContent = tr('📊 رسم بياني');
    return;
  }

  el.style.background = '#ffffffeb';
  el.style.border = '1px solid #d5cfc2';
  el.style.borderRadius = '8px';
  el.style.padding = '8px';

  const colors = ['#B4531A', '#2563EB', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899'];
  const svg = svgEl('svg', { viewBox: '0 0 320 205', preserveAspectRatio: 'none' });
  const allVals = data.series.flatMap(s => s.vals);
  const maxVal = Math.max(1, ...allVals);

  if (data.type === 'bar') {
    const n = data.cats.length;
    const sCount = data.series.length;
    const slot = 260 / n;
    const groupW = slot * 0.78;
    const barW = groupW / sCount;

    data.cats.forEach((cat, i) => {
      const baseX = 35 + i * slot + slot * 0.11;
      data.series.forEach((ser, si) => {
        const v = ser.vals[i] ?? 0;
        const barH = (v / maxVal) * 135;
        const x = baseX + si * barW;
        const y = 165 - barH;
        svg.append(svgEl('rect', { x, y, width: Math.max(3, barW - 1.5), height: barH, fill: colors[si % colors.length], rx: 2 }));
        if (sCount <= 2) {
          const valTxt = svgEl('text', { x: x + barW / 2, y: y - 3, 'text-anchor': 'middle', 'font-size': '9', fill: '#444' });
          valTxt.textContent = v;
          svg.append(valTxt);
        }
      });
      const lbl = svgEl('text', { x: 35 + i * slot + slot / 2, y: 184, 'text-anchor': 'middle', 'font-size': '10.5', fill: '#333' });
      lbl.textContent = cat || (i + 1);
      svg.append(lbl);
    });
  } else if (data.type === 'line') {
    const n = data.cats.length;
    const step = n > 1 ? 250 / (n - 1) : 125;
    data.series.forEach((ser, si) => {
      const col = colors[si % colors.length];
      const pts = data.cats.map((_, i) => `${35 + i * step},${165 - ((ser.vals[i] ?? 0) / maxVal) * 130}`).join(' ');
      svg.append(svgEl('polyline', { fill: 'none', stroke: col, 'stroke-width': '2.8', points: pts }));
      data.cats.forEach((_, i) => {
        const cx = 35 + i * step, cy = 165 - ((ser.vals[i] ?? 0) / maxVal) * 130;
        svg.append(svgEl('circle', { cx, cy, r: 4, fill: col }));
      });
    });
    data.cats.forEach((cat, i) => {
      const lbl = svgEl('text', { x: 35 + i * step, y: 184, 'text-anchor': 'middle', 'font-size': '10.5', fill: '#333' });
      lbl.textContent = cat || (i + 1);
      svg.append(lbl);
    });
  } else if (data.type === 'pie') {
    const vals = data.series[0].vals;
    const total = vals.reduce((a, b) => a + Math.max(0, b), 0) || 1;
    let angle = -Math.PI / 2;
    vals.forEach((v, i) => {
      const slice = (Math.max(0, v) / total) * Math.PI * 2;
      const x1 = 110 + 70 * Math.cos(angle), y1 = 100 + 70 * Math.sin(angle);
      angle += slice;
      const x2 = 110 + 70 * Math.cos(angle), y2 = 100 + 70 * Math.sin(angle);
      const large = slice > Math.PI ? 1 : 0;
      const d = `M 110 100 L ${x1} ${y1} A 70 70 0 ${large} 1 ${x2} ${y2} Z`;
      svg.append(svgEl('path', { d, fill: colors[i % colors.length] }));
      svg.append(svgEl('rect', { x: 205, y: 25 + i * 22, width: 10, height: 10, fill: colors[i % colors.length] }));
      const lbl = svgEl('text', { x: 220, y: 34 + i * 22, 'font-size': '11', fill: '#333' });
      lbl.textContent = `${data.cats[i] || ''}: ${v}`;
      svg.append(lbl);
    });
  }

  if (data.type !== 'pie' && data.series.length > 1) {
    data.series.forEach((ser, si) => {
      const lx = 35 + si * 95;
      svg.append(svgEl('rect', { x: lx, y: 6, width: 9, height: 9, fill: colors[si % colors.length], rx: 2 }));
      const lt = svgEl('text', { x: lx + 13, y: 14, 'font-size': '9.5', fill: '#333' });
      lt.textContent = ser.name;
      svg.append(lt);
    });
  }
  el.append(svg);
}

/* رسم الجداول التفاعلية */
function renderTable(el, gf, L, ctx, shapeIdx, isInteractive) {
  el.dataset.kind = 'table';
  el.classList.add('table-shape');
  const tbl = desc(gf, A, 'tbl')[0];
  if (!tbl) return;
  const grid = kids(kid(tbl, A, 'tblGrid'), A, 'gridCol').map(c => +c.getAttribute('w') || 1);
  const tot = grid.reduce((a, b) => a + b, 0) || 1;
  const tp = kid(tbl, A, 'tblPr');
  const hasStyle = !!(tp && kid(tp, A, 'tableStyleId')), firstRow = tp && tp.getAttribute('firstRow') === '1', band = tp && tp.getAttribute('bandRow') === '1';
  const accent = schemeHex('accent1', ctx) || '4472C4';
  const t = h('table', { class: 'pt' });
  const cg = h('colgroup');
  grid.forEach(w => cg.append(h('col', { style: { width: (w / tot * 100) + '%' } })));
  t.append(cg);

  const rows = kids(tbl, A, 'tr'), totH = rows.reduce((a, r) => a + (+r.getAttribute('h') || 0), 0) || 1;
  rows.forEach((tr, ri) => {
    const row = h('tr', { style: { height: ((+tr.getAttribute('h') || 0) / totH * 100) + '%' } });
    let ci = 0;
    for (const tc of kids(tr, A, 'tc')) {
      const colIndex = ci++;
      if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
      const td = h('td');
      const cs = +tc.getAttribute('gridSpan') || 1, rs = +tc.getAttribute('rowSpan') || 1;
      if (cs > 1) td.colSpan = cs;
      if (rs > 1) td.rowSpan = rs;
      const tcPr = kid(tc, A, 'tcPr');
      let f = fillFrom(tcPr, ctx);
      const isHead = hasStyle && firstRow && ri === 0;
      if (!f && hasStyle) {
        if (isHead) f = { type: 'solid', hex: accent };
        else if (band && ri % 2 === 1) { const [r, g, b] = hex2rgb(accent); f = { type: 'solid', hex: rgb2hex(r * .2 + 204, g * .2 + 204, b * .2 + 204) }; }
      }
      if (f) applyFill(td, f, L, pk);

      const txb = txBodyOf(tc), st = { counters: [] };
      for (const p of kids(txb, A, 'p')) {
        const d = renderParagraph(ctx, gf, p, st);
        if (isHead && !f) { d.style.fontWeight = '700'; d.style.color = '#fff'; $$('span', d).forEach(s => { s.style.color = '#fff'; s.style.fontWeight = '700'; }); }
        td.append(d);
      }

      if (isInteractive) {
        td.dataset.r = ri;
        td.dataset.c = colIndex;
        if (curPptx.selShapeIdx === shapeIdx && curPptx.selCell && curPptx.selCell.r === ri && curPptx.selCell.c === colIndex) {
          td.classList.add('cell-sel');
        }
        const enterCellEdit = (ev) => {
          if (ev) ev.stopPropagation();
          td.contentEditable = 'true';
          td.focus();
          td.onblur = () => {
            td.contentEditable = 'false';
            setCellTextXml(tc, td.innerText);
            pushHistory('editCell');
            renderPptxEditor();
          };
        };
        td.addEventListener('pointerdown', (ev) => {
          ev.stopPropagation();
          if (ev.pointerType === 'touch' && td.contentEditable !== 'true') {
            const now = Date.now();
            const prev = td._lastTap;
            if (prev && (now - prev.t) < 340 && Math.hypot(ev.clientX - prev.x, ev.clientY - prev.y) < 24) {
              td._lastTap = null;
              ev.preventDefault();
              enterCellEdit(ev);
              return;
            }
            td._lastTap = { t: now, x: ev.clientX, y: ev.clientY };
          }
          curPptx.selShapeIdx = shapeIdx;
          curPptx.selCell = { r: ri, c: colIndex };
          $$('.cell-sel').forEach(x => x.classList.remove('cell-sel'));
          td.classList.add('cell-sel');
          $$('.shape').forEach(s => s.classList.toggle('sel', s === el));
          const propsEl = $('#app .props');
          if (propsEl && curPptx.deck) renderShapeProps(propsEl, el.closest('.slide'), curPptx.deck, curPptx.slideIdx);
        });
        td.addEventListener('dblclick', enterCellEdit);
      }
      row.append(td);
    }
    t.append(row);
  });
  el.append(t);
}

function renderLine(el, sh, g, ctx) {
  el.dataset.kind = 'line';
  const ln = lineSpec(sh, ctx);
  const s = svgEl('svg', { viewBox: '0 0 100 100', preserveAspectRatio: 'none' });
  const l = svgEl('line', { x1: g.flipH ? 100 : 0, x2: g.flipH ? 0 : 100, y1: g.flipV ? 100 : 0, y2: g.flipV ? 0 : 100, 'vector-effect': 'non-scaling-stroke' });
  if (ln) {
    l.style.stroke = cssColor(ln.color);
    l.style.strokeWidth = `max(2px,${lenCss(ln.w, ctx.deck.W)})`;
    const d = kid(kid(kid(sh, P, 'spPr'), A, 'ln'), A, 'prstDash');
    if (d && d.getAttribute('val') !== 'solid') l.style.strokeDasharray = '6 4';
  } else l.style.stroke = '#1f1c17';
  s.append(l); el.append(s);
}

const CLIPS = {
  triangle: 'polygon(50% 0,100% 100%,0 100%)',
  diamond: 'polygon(50% 0,100% 50%,50% 100%,0 50%)',
  rtTriangle: 'polygon(0 0,100% 100%,0 100%)',
  hexagon: 'polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%)',
  parallelogram: 'polygon(20% 0,100% 0,80% 100%,0 100%)',
  star5: 'polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)',
  wedgeRoundRectCallout: 'polygon(0% 0%,100% 0%,100% 78%,58% 78%,42% 100%,42% 78%,0% 78%)',
  rightArrow: 'polygon(0 25%,65% 25%,65% 0,100% 50%,65% 100%,65% 75%,0 75%)',
  leftArrow: 'polygon(35% 0,35% 25%,100% 25%,100% 75%,35% 75%,35% 100%,0 50%)',
  upArrow: 'polygon(50% 0,100% 35%,75% 35%,75% 100%,25% 100%,25% 35%,0 35%)',
  downArrow: 'polygon(25% 0,75% 0,75% 65%,100% 65%,50% 100%,0 65%,25% 65%)'
};

function isShapeLocked(sh) {
  const cNvPr = kid(nvOf(sh), P, 'cNvPr');
  return cNvPr?.getAttribute('descr') === 'docstudio:locked';
}

function toggleLockShape(deck, slideIdx) {
  if (curPptx.selShapeIdx == null) return;
  const doc = pk.doc(deck.slides[slideIdx].path);
  const sh = shapesOf(doc)[curPptx.selShapeIdx];
  const cNvPr = sh && kid(nvOf(sh), P, 'cNvPr');
  if (!cNvPr) return;
  const locked = isShapeLocked(sh);
  if (locked) cNvPr.removeAttribute('descr');
  else cNvPr.setAttribute('descr', 'docstudio:locked');
  pushHistory('lockShape');
  renderPptxEditor();
  toast(locked ? 'تم فك قفل العنصر' : 'تم قفل العنصر');
}

function renderSp(el, sh, L, ctx, g) {
  const spPr = kid(sh, P, 'spPr'), geom = kid(spPr, A, 'prstGeom');
  const prst = geom ? geom.getAttribute('prst') : (kid(spPr, A, 'custGeom') ? 'cust' : 'rect');
  if (/^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst)) return renderLine(el, sh, g, ctx);
  const fill = shapeFill(sh, ctx), ln = lineSpec(sh, ctx);
  applyFill(el, fill, L, pk);
  if (ln && prst !== 'cust') el.style.border = `max(1px,${lenCss(ln.w, ctx.deck.W)}) solid ${cssColor(ln.color)}`;
  if (prst === 'ellipse') {
    el.style.borderRadius = '50%';
  } else if (prst === 'roundRect') {
    const gd = kid(kid(geom, A, 'avLst'), A, 'gd');
    const fmla = gd?.getAttribute('fmla') || '';
    const m = /val\s+(\d+)/.exec(fmla);
    const pct = m ? Math.min(50, Math.max(0, Math.round((+m[1] / 50000) * 50))) : 12;
    el.style.borderRadius = pct + '%';
  } else if (CLIPS[prst]) {
    el.style.clipPath = CLIPS[prst];
  }

  const eff = kid(spPr, A, 'effectLst');
  if (eff && kid(eff, A, 'outerShdw')) el.style.boxShadow = '0 6px 18px rgba(0,0,0,0.22)';

  const txb = txBodyOf(sh);
  el.dataset.kind = txb ? 'text' : 'other';
  if (txb) el.append(renderTextBody(ctx, sh, txb));
}

function renderShape(sh, L, ctx, space, isStatic, shapeIdx, isInteractive) {
  const g = geomOf(sh, ctx);
  const el = h('div', { class: 'shape' });
  if (!g) { el.style.display = 'none'; el.dataset.kind = 'other'; return el; }
  el._g = g;
  if (!isStatic && isShapeLocked(sh)) el.classList.add('locked');
  placeEl(el, g, space);
  if (g.rot) el.style.transform = `rotate(${g.rot}deg)`;
  switch (sh.localName) {
    case 'pic': renderPic(el, sh, L, ctx, g); break;
    case 'cxnSp': renderLine(el, sh, g, ctx); break;
    case 'sp': renderSp(el, sh, L, ctx, g); break;
    case 'grpSp': {
      el.dataset.kind = 'group';
      const sp2 = { ox: g.chx != null ? g.chx : g.x, oy: g.chy != null ? g.chy : g.y, w: g.chw || g.w || 1, h: g.chh || g.h || 1 };
      for (const c of kids(sh, P).filter(e => SHAPE_TAGS.has(e.localName))) el.append(renderShape(c, L, ctx, sp2, isStatic, -1, false));
      break;
    }
    case 'graphicFrame': {
      if (desc(sh, A, 'tbl').length) renderTable(el, sh, L, ctx, shapeIdx, isInteractive);
      else renderChart(el, sh, L, ctx);
      break;
    }
    default: break;
  }
  return el;
}

function backgroundOf(ctx) {
  for (const L of [ctx.slide, ctx.layout, ctx.master]) {
    if (!L) continue;
    const bg = kid(kid(L.doc.documentElement, P, 'cSld'), P, 'bg'); if (!bg) continue;
    const bgPr = kid(bg, P, 'bgPr');
    if (bgPr) { const f = fillFrom(bgPr, ctx); if (f) return { fill: f, L }; }
    const ref = kid(bg, P, 'bgRef');
    if (ref) { const ce = colorChild(ref), col = ce && parseColorEl(ce, ctx); if (col) return { fill: { type: 'solid', hex: col.hex }, L }; }
  }
  return null;
}

function renderSlideEl(deck, idx, interactive) {
  const sl = deck.slides[idx], ctx = buildCtx(deck, sl.path);
  CUR_WPT = ctx.wpt;
  const root = h('div', { class: 'slide' + (interactive ? ' inter' : '') });
  if (interactive && curPptx.showGrid) root.classList.add('show-grid');
  root.style.aspectRatio = `${deck.W}/${deck.H}`;
  const bg = backgroundOf(ctx);
  root.style.background = '#fff';
  if (bg) applyFill(root, bg.fill, bg.L, pk);
  const space = { ox: 0, oy: 0, w: deck.W, h: deck.H };
  const showLayout = ctx.slide.doc.documentElement.getAttribute('showMasterSp') !== '0';
  const showMaster = showLayout && !(ctx.layout && ctx.layout.doc.documentElement.getAttribute('showMasterSp') === '0');
  for (const [L, ok] of [[ctx.master, showMaster], [ctx.layout, showLayout]]) {
    if (!L || !ok) continue;
    for (const sh of shapesOf(L.doc)) {
      if (phOf(sh)) continue;
      try { const e = renderShape(sh, L, ctx, space, true, -1, false); e.classList.add('static'); root.append(e); } catch (_) {}
    }
  }
  const all = shapesOf(ctx.slide.doc);
  all.forEach((sh, i) => {
    let e;
    try { e = renderShape(sh, ctx.slide, ctx, space, false, i, interactive); }
    catch (err) { console.warn('shape render failed', err); e = h('div', { class: 'shape', style: { display: 'none' } }); }
    e.dataset.i = i; e._sh = sh; root.append(e);
  });
  root._ctx = ctx;
  return root;
}
