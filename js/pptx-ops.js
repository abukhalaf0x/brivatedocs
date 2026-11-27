'use strict';
/* ═════════════════════════════════════════════════════════════
   القسم 4 من 6: عمليات PPTX المتكاملة، الثيمات، والمخططات الذكية
   ═════════════════════════════════════════════════════════════ */

function nextShapeId(spTree) {
  let maxId = 1;
  for (const el of desc(spTree, P, 'cNvPr')) {
    const id = parseInt(el.getAttribute('id') || '0', 10);
    if (!isNaN(id)) maxId = Math.max(maxId, id);
  }
  return maxId + 1;
}

function registerSlidePart(deck, slidePath, insertAfterIdx) {
  setOverride(pk, '/' + slidePath, CT_SLIDE);
  const rid = addRel(pk, deck.pres, REL + 'slide', slidePath);
  const presDoc = pk.doc(deck.pres);
  const sldIdLst = kid(presDoc.documentElement, P, 'sldIdLst') || presDoc.documentElement.appendChild(mk(presDoc, P, 'p:sldIdLst'));
  let maxSlideId = 255;
  for (const s of kids(sldIdLst, P, 'sldId')) {
    const id = parseInt(s.getAttribute('id') || '0', 10);
    if (!isNaN(id)) maxSlideId = Math.max(maxSlideId, id);
  }
  const sldIdEl = mk(presDoc, P, 'p:sldId', { id: String(maxSlideId + 1) });
  sldIdEl.setAttributeNS(R, 'r:id', rid);
  const existing = kids(sldIdLst, P, 'sldId');
  if (insertAfterIdx != null && existing[insertAfterIdx]) {
    insertAfter(existing[insertAfterIdx], sldIdEl);
  } else {
    sldIdLst.appendChild(sldIdEl);
  }
  return rid;
}

function slideLayoutPath(slidePath) {
  const rels = readRels(pk, slidePath);
  const lRel = relByType(rels, 'slideLayout');
  return lRel && lRel.path ? lRel.path : null;
}

function changeSlideLayout(slidePath, newLayoutPath) {
  if (!slidePath || !newLayoutPath || !pk.has(newLayoutPath)) {
    toast('تعذر تحديد تخطيط الشريحة');
    return;
  }
  const rels = readRels(pk, slidePath);
  const oldRel = relByType(rels, 'slideLayout');
  if (oldRel) removeRel(pk, slidePath, oldRel.id);
  addRel(pk, slidePath, REL + 'slideLayout', newLayoutPath);
}

/* ── مبدّل الثيم الكامل للعرض (Global Deck Theme Switcher) ── */
const PPTX_PRO_THEMES = {
  classic: { name: 'كلاسيكي أبيض', bg: 'FFFFFF', ink: '1F1C17', acc: 'B4531A' },
  midnight: { name: 'ليلي فاخر', bg: '0F172A', ink: 'F8FAFC', acc: '38BDF8' },
  royal: { name: 'أزرق ملكي', bg: '0F2547', ink: 'FFFFFF', acc: 'F59E0B' },
  emerald: { name: 'أخضر زمردي', bg: '064E3B', ink: 'ECFDF5', acc: '34D399' },
  sand: { name: 'رملي دافئ', bg: 'FAF6EE', ink: '2C241B', acc: '0D9488' }
};

function applyGlobalDeckTheme(themeId) {
  const th = PPTX_PRO_THEMES[themeId];
  if (!th || !curPptx.deck) return;
  const deck = curPptx.deck;

  for (const sl of deck.slides) {
    const doc = pk.doc(sl.path);
    const cSld = kid(doc.documentElement, P, 'cSld');
    if (cSld) {
      const oldBg = kid(cSld, P, 'bg');
      if (oldBg) removeNode(oldBg);
      cSld.insertBefore(frag(doc, `<p:bg xmlns:p="${P}" xmlns:a="${A}"><p:bgPr><a:solidFill><a:srgbClr val="${th.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`), cSld.firstChild);
    }

    for (const sh of shapesOf(doc)) {
      const txb = txBodyOf(sh);
      if (!txb) continue;
      for (const r of desc(txb, A, 'r')) {
        let rp = kid(r, A, 'rPr');
        if (!rp) { rp = mk(doc, A, 'a:rPr'); r.insertBefore(rp, r.firstChild); }
        const isBold = rp.getAttribute('b') === '1';
        const sz = +(rp.getAttribute('sz') || 1800);
        const colHex = (isBold && sz >= 2800) ? th.acc : th.ink;
        for (const f of kids(rp, A)) if (f.localName === 'solidFill') removeNode(f);
        const sf = mk(doc, A, 'a:solidFill');
        sf.appendChild(mk(doc, A, 'a:srgbClr', { val: colHex }));
        rp.appendChild(sf);
      }
    }
  }
  pushHistory('globalTheme:' + themeId);
  renderPptxEditor();
  toast('تم تطبيق القالب على العرض بالكامل');
}

/* ── خلفية الشريحة (لون ثابت / تدرج / صورة / تطبيق على الكل) ── */
function setSlideBackgroundColor(slidePath, hexColor) {
  const doc = pk.doc(slidePath);
  const cSld = kid(doc.documentElement, P, 'cSld');
  if (!cSld) return;
  const oldBg = kid(cSld, P, 'bg');
  if (oldBg) removeNode(oldBg);

  const hex = hexColor.replace('#', '').toUpperCase();
  const bg = frag(doc, `<p:bg xmlns:p="${P}" xmlns:a="${A}">
    <p:bgPr>
      <a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>
      <a:effectLst/>
    </p:bgPr>
  </p:bg>`);
  cSld.insertBefore(bg, cSld.firstChild);

  pushHistory('bgColor');
  renderPptxEditor();
  toast('تم تغيير لون الخلفية');
}

function setSlideBackgroundGradient(slidePath, hex1, hex2, angleDeg) {
  const doc = pk.doc(slidePath);
  const cSld = kid(doc.documentElement, P, 'cSld');
  if (!cSld) return;
  const oldBg = kid(cSld, P, 'bg');
  if (oldBg) removeNode(oldBg);

  const bg = mk(doc, P, 'p:bg');
  const bgPr = mk(doc, P, 'p:bgPr');
  bgPr.appendChild(buildGradFillXml(doc, hex1, hex2, angleDeg, 100));
  bgPr.appendChild(mk(doc, A, 'a:effectLst'));
  bg.appendChild(bgPr);
  cSld.insertBefore(bg, cSld.firstChild);

  pushHistory('bgGradient');
  renderPptxEditor();
  toast('تم تطبيق التدرج اللوني للخلفية');
}

function setSlideBackgroundImage(slidePath, imgInfo) {
  const doc = pk.doc(slidePath);
  const cSld = kid(doc.documentElement, P, 'cSld');
  if (!cSld) return;

  const mediaPath = uniquePath(pk, 'ppt/media', 'bgImage', imgInfo.ext);
  pk.setBin(mediaPath, imgInfo.bytes);
  ensureDefault(pk, imgInfo.ext, CT_IMG[imgInfo.ext] || 'image/png');
  const rid = addRel(pk, slidePath, REL + 'image', mediaPath);

  const oldBg = kid(cSld, P, 'bg');
  if (oldBg) removeNode(oldBg);

  const bg = frag(doc, `<p:bg xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">
    <p:bgPr>
      <a:blipFill rotWithShape="1">
        <a:blip r:embed="${rid}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </a:blipFill>
      <a:effectLst/>
    </p:bgPr>
  </p:bg>`);
  cSld.insertBefore(bg, cSld.firstChild);

  pushHistory('bgImage');
  renderPptxEditor();
  toast('تم تعيين صورة الخلفية');
}

function applyBgToAllSlides(deck, srcSlidePath) {
  const srcDoc = pk.doc(srcSlidePath);
  const srcCSld = kid(srcDoc.documentElement, P, 'cSld');
  const srcBg = srcCSld && kid(srcCSld, P, 'bg');
  if (!srcBg) {
    toast('الشريحة الحالية ليس لها خلفية مخصصة');
    return;
  }

  const srcRels = readRels(pk, srcSlidePath);
  const blip = desc(srcBg, A, 'blip')[0];
  const srcRid = blip && blip.getAttributeNS(R, 'embed');
  const srcMediaPath = srcRid && srcRels.get(srcRid)?.path;

  for (const sl of deck.slides) {
    if (sl.path === srcSlidePath) continue;
    const dstDoc = pk.doc(sl.path);
    const dstCSld = kid(dstDoc.documentElement, P, 'cSld');
    if (!dstCSld) continue;
    const oldBg = kid(dstCSld, P, 'bg');
    if (oldBg) removeNode(oldBg);

    const clonedBg = dstDoc.importNode(srcBg, true);
    if (srcMediaPath && pk.has(srcMediaPath)) {
      const dstBlip = desc(clonedBg, A, 'blip')[0];
      if (dstBlip) {
        const newRid = addRel(pk, sl.path, REL + 'image', srcMediaPath);
        dstBlip.setAttributeNS(R, 'r:embed', newRid);
      }
    }
    dstCSld.insertBefore(clonedBg, dstCSld.firstChild);
  }

  pushHistory('bgApplyAll');
  renderPptxEditor();
  toast('تم تطبيق الخلفية على جميع الشرائح');
}

/* ── إدارة الشرائح (إضافة / نسخ / حذف / ترتيب) ── */
function addSlide(deck) {
  const layouts = availableLayouts();
  const blankLayout = layouts.find(l => /blank|فارغة/i.test(l.name)) || layouts[layouts.length - 1] || layouts[0];
  const newPath = uniquePath(pk, 'ppt/slides', 'slide', 'xml');

  pk.setText(newPath, XMLDECL + `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree></p:cSld>
  </p:sld>`);

  if (blankLayout) addRel(pk, newPath, REL + 'slideLayout', blankLayout.path);
  registerSlidePart(deck, newPath, curPptx.slideIdx);

  curPptx.slideIdx = Math.min(deck.slides.length, curPptx.slideIdx + 1);
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  curPptx.selCell = null;
  pushHistory('addSlide');
  renderPptxEditor();
  toast('تمت إضافة شريحة جديدة');
}

function duplicateSlide(deck, idx) {
  const src = deck.slides[idx];
  if (!src) return;
  const newPath = uniquePath(pk, 'ppt/slides', 'slide', 'xml');
  const srcDoc = pk.doc(src.path);
  pk.setText(newPath, serialize(srcDoc.cloneNode(true)));

  const srcRels = readRels(pk, src.path);
  const rp = relsPathOf(newPath);
  pk.setText(rp, XMLDECL + `<Relationships xmlns="${NS.rel}"/>`);
  const relDoc = pk.doc(rp);
  for (const r of srcRels.values()) {
    if (r.type.endsWith('/notesSlide')) continue;
    let target = r.target;
    if (r.type.endsWith('/chart') && r.path && pk.has(r.path)) {
      const chartPath = uniquePath(pk, 'ppt/charts', 'chart', 'xml');
      const chartData = pk.files.get(r.path);
      if (typeof chartData === 'string') pk.setText(chartPath, chartData); else pk.setBin(chartPath, chartData);
      const sourceChartRels = relsPathOf(r.path);
      if (pk.has(sourceChartRels)) pk.setText(relsPathOf(chartPath), pk.files.get(sourceChartRels));
      const override = kids(ctRoot(pk), null, 'Override').find(x => x.getAttribute('PartName') === '/' + r.path);
      if (override) setOverride(pk, '/' + chartPath, override.getAttribute('ContentType'));
      target = relPath(newPath, chartPath);
    } else if (r.mode !== 'External') target = relPath(newPath, r.path);
    const e = relDoc.createElementNS(NS.rel, 'Relationship');
    e.setAttribute('Id', r.id); e.setAttribute('Type', r.type); e.setAttribute('Target', target);
    if (r.mode === 'External') e.setAttribute('TargetMode', 'External');
    relDoc.documentElement.appendChild(e);
  }

  registerSlidePart(deck, newPath, idx);
  curPptx.slideIdx = idx + 1;
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  pushHistory('duplicateSlide');
  renderPptxEditor();
  toast('تم نسخ الشريحة');
}

function deleteSlide(deck, idx) {
  if (deck.slides.length <= 1) {
    toast('لا يمكن حذف آخر شريحة');
    return;
  }
  const sl = deck.slides[idx];
  if (!sl) return;
  removeRel(pk, deck.pres, sl.rid);
  removeNode(sl.el);
  removeOverride(pk, '/' + sl.path);
  pk.remove(sl.path);
  pk.remove(relsPathOf(sl.path));

  const oldCurrent = curPptx.slideIdx;
  if (idx < oldCurrent) curPptx.slideIdx = oldCurrent - 1;
  else if (idx === oldCurrent) curPptx.slideIdx = Math.min(idx, deck.slides.length - 2);
  curPptx.slideIdx = Math.max(0, Math.min(curPptx.slideIdx, deck.slides.length - 1));
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  curPptx.selCell = null;
  pushHistory('deleteSlide');
  renderPptxEditor();
  toast('تم حذف الشريحة');
}

function moveSlide(deck, fromIdx, toIdx) {
  if (fromIdx === toIdx || toIdx < 0 || toIdx >= deck.slides.length) return;
  const presDoc = pk.doc(deck.pres);
  const sldIdLst = kid(presDoc.documentElement, P, 'sldIdLst');
  const items = kids(sldIdLst, P, 'sldId');
  const movedEl = items[fromIdx];
  if (!movedEl) return;

  if (fromIdx < toIdx) insertAfter(items[toIdx], movedEl);
  else sldIdLst.insertBefore(movedEl, items[toIdx]);

  curPptx.slideIdx = toIdx;
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  pushHistory('moveSlide');
  renderPptxEditor();
  toast('تم تغيير ترتيب الشرائح');
}

/* ── إضافة العناصر والمخططات الذكية SmartArt إلى الشريحة ── */
function addTextBoxToSlide(deck, slideIdx) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);
  const w = Math.round(deck.W * 0.42), h = Math.round(deck.H * 0.14);
  const x = Math.round((deck.W - w) / 2), y = Math.round((deck.H - h) / 2);

  const xml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr anchor="ctr"/>
      <a:lstStyle/>
      <a:p><a:pPr algn="ctr"${currentLanguage === 'en' ? '' : ' rtl="1"'}/><a:r><a:rPr lang="${currentLanguage === 'en' ? 'en-US' : 'ar-SA'}" sz="2400"><a:solidFill><a:srgbClr val="1F1C17"/></a:solidFill></a:rPr><a:t>${currentLanguage === 'en' ? 'Type text here' : 'اكتب النص هنا'}</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addTextBox');
  renderPptxEditor();
  toast('تمت إضافة مربع نص');
}

function addPresetShapeToSlide(deck, slideIdx, prst) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);
  const w = Math.round(deck.W * 0.22), h = Math.round(deck.W * 0.16);
  const x = Math.round((deck.W - w) / 2), y = Math.round((deck.H - h) / 2);

  const xml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
      <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="B4531A"/></a:solidFill>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="8F3F12"/></a:solidFill></a:ln>
    </p:spPr>
  </p:sp>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addShape:' + prst);
  renderPptxEditor();
  toast('تمت إضافة الشكل');
}

function insertSmartArtDiagram(type) {
  if (!curPptx.deck) return;
  const deck = curPptx.deck;
  const slPath = deck.slides[curPptx.slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  let id = nextShapeId(spTree);
  const W0 = deck.W, H0 = deck.H;

  const addBox = (x, y, w, h, bg, txt, prst = 'roundRect') => {
    const xml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
      <p:nvSpPr><p:cNvPr id="${id++}" name="SmartArt ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="${Math.round(W0 * x)}" y="${Math.round(H0 * y)}"/><a:ext cx="${Math.round(W0 * w)}" cy="${Math.round(H0 * h)}"/></a:xfrm>
        <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${bg}"/></a:solidFill>
      </p:spPr>
      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>${txt}</a:t></a:r></a:p></p:txBody>
    </p:sp>`;
    spTree.appendChild(frag(doc, xml));
  };

  if (type === 'pyramid') {
    addBox(0.38, 0.20, 0.24, 0.16, 'B4531A', 'القمة / الرؤية');
    addBox(0.28, 0.40, 0.44, 0.16, '2563EB', 'الإدارة والتخطيط');
    addBox(0.18, 0.60, 0.64, 0.16, '0F2547', 'التنفيذ والعمليات');
  } else if (type === 'cycle') {
    addBox(0.40, 0.16, 0.20, 0.16, 'B4531A', '1. تحليل', 'ellipse');
    addBox(0.62, 0.42, 0.20, 0.16, '2563EB', '2. تصميم', 'ellipse');
    addBox(0.40, 0.68, 0.20, 0.16, '10B981', '3. تطوير', 'ellipse');
    addBox(0.18, 0.42, 0.20, 0.16, '8B5CF6', '4. قياس', 'ellipse');
  } else {
    ['المرحلة الأولى', 'المرحلة الثانية', 'المرحلة الثالثة'].forEach((label, i) => {
      addBox(0.10 + i * 0.28, 0.38, 0.24, 0.22, i === 0 ? 'B4531A' : i === 1 ? '2563EB' : '10B981', label, 'rightArrow');
    });
  }
  pushHistory('smartArt:' + type);
  renderPptxEditor();
  toast('تم إدراج المخطط الذكي');
}

function addIconToSlide(deck, slideIdx, iconObj) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);
  const size = Math.round(deck.W * 0.1);
  const x = Math.round((deck.W - size) / 2), y = Math.round((deck.H - size) / 2);

  const xml = `<p:sp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvSpPr><p:cNvPr id="${id}" name="Icon ${iconObj.id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${size}" cy="${size}"/></a:xfrm>
      <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="FFF3E8"/></a:solidFill>
      <a:ln w="19050"><a:solidFill><a:srgbClr val="B4531A"/></a:solidFill></a:ln>
    </p:spPr>
    <p:txBody>
      <a:bodyPr anchor="ctr"/>
      <a:lstStyle/>
      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="3600" b="1"><a:solidFill><a:srgbClr val="B4531A"/></a:solidFill></a:rPr><a:t>${iconObj.glyph}</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addIcon');
  renderPptxEditor();
  toast('تمت إضافة الأيقونة');
}

function addLineToSlide(deck, slideIdx) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);
  const w = Math.round(deck.W * 0.35), h = 38100;
  const x = Math.round((deck.W - w) / 2), y = Math.round(deck.H / 2);

  const xml = `<p:cxnSp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvCxnSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
      <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
      <a:ln w="25400"><a:solidFill><a:srgbClr val="1F1C17"/></a:solidFill></a:ln>
    </p:spPr>
  </p:cxnSp>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addLine');
  renderPptxEditor();
  toast('تمت إضافة خط');
}

function addTableToSlide(deck, slideIdx) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);
  const rows = 3, cols = 3;
  const w = Math.round(deck.W * 0.6), h = Math.round(deck.H * 0.3);
  const x = Math.round((deck.W - w) / 2), y = Math.round((deck.H - h) / 2);
  const colW = Math.round(w / cols), rowH = Math.round(h / rows);

  const gridCols = Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('');
  const makeCell = (txt, isHdr) => `<a:tc>
    <a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr sz="1600"${isHdr ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${isHdr ? 'FFFFFF' : '1F1C17'}"/></a:solidFill></a:rPr><a:t>${txt}</a:t></a:r></a:p></a:txBody>
    <a:tcPr>${isHdr ? '<a:solidFill><a:srgbClr val="B4531A"/></a:solidFill>' : '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'}</a:tcPr>
  </a:tc>`;

  let trsXml = '';
  for (let r = 0; r < rows; r++) {
    let cells = '';
    for (let c = 0; c < cols; c++) {
      cells += makeCell(r === 0 ? `عنوان ${c + 1}` : `خلية ${r}-${c + 1}`, r === 0);
    }
    trsXml += `<a:tr h="${rowH}">${cells}</a:tr>`;
  }

  const xml = `<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}">
    <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${gridCols}</a:tblGrid>${trsXml}</a:tbl>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  curPptx.selCell = { r: 0, c: 0 };
  pushHistory('addTable');
  renderPptxEditor();
  toast('تمت إضافة جدول (انقر مرتين على أي خلية لتعديلها)');
}

function modifyTableStructure(tbl, action) {
  const doc = tbl.ownerDocument;
  const grid = kid(tbl, A, 'tblGrid');
  const rows = kids(tbl, A, 'tr');
  if (!grid || !rows.length) return;

  if (action === 'addRow') {
    const templateRow = rows[rows.length - 1];
    const cloned = templateRow.cloneNode(true);
    for (const tc of kids(cloned, A, 'tc')) setCellTextXml(tc, '');
    tbl.appendChild(cloned);
  } else if (action === 'delRow') {
    if (rows.length <= 1) return;
    const idx = curPptx.selCell ? Math.min(curPptx.selCell.r, rows.length - 1) : rows.length - 1;
    removeNode(rows[idx]);
    curPptx.selCell = null;
  } else if (action === 'addCol') {
    const cols = kids(grid, A, 'gridCol');
    const colW = cols[cols.length - 1]?.getAttribute('w') || '1500000';
    grid.appendChild(mk(doc, A, 'a:gridCol', { w: colW }));
    for (const r of rows) {
      const cells = kids(r, A, 'tc');
      const clonedTc = cells[cells.length - 1].cloneNode(true);
      setCellTextXml(clonedTc, '');
      r.appendChild(clonedTc);
    }
  } else if (action === 'delCol') {
    const cols = kids(grid, A, 'gridCol');
    if (cols.length <= 1) return;
    const cIdx = curPptx.selCell ? Math.min(curPptx.selCell.c, cols.length - 1) : cols.length - 1;
    removeNode(cols[cIdx]);
    for (const r of rows) {
      const cells = kids(r, A, 'tc');
      if (cells[cIdx]) removeNode(cells[cIdx]);
    }
    curPptx.selCell = null;
  }

  pushHistory('tableStruct:' + action);
  renderPptxEditor();
}

function buildChartXmlDoc(type, cats, seriesList) {
  const tagMap = { bar: 'c:barChart', line: 'c:lineChart', pie: 'c:pieChart' };
  const chartTag = tagMap[type] || 'c:barChart';
  const extraOpts = type === 'bar'
    ? '<c:barDir val="col"/><c:grouping val="clustered"/>'
    : type === 'line' ? '<c:grouping val="standard"/>' : '';

  const catPts = cats.map((c, i) => `<c:pt idx="${i}"><c:v>${String(c).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</c:v></c:pt>`).join('');
  const sersXml = seriesList.map((ser, si) => {
    const valPts = ser.vals.map((v, i) => `<c:pt idx="${i}"><c:v>${+v || 0}</c:v></c:pt>`).join('');
    return `<c:ser>
      <c:idx val="${si}"/><c:order val="${si}"/>
      <c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${String(ser.name).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:ptCount val="${cats.length}"/>${catPts}</c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ser.vals.length}"/>${valPts}</c:numCache></c:numRef></c:val>
    </c:ser>`;
  }).join('');

  const axesXml = type === 'pie' ? '' : `<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const axisIds = type === 'pie' ? '' : '<c:axId val="48650112"/><c:axId val="48672768"/>';
  return XMLDECL + `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}">
    <c:chart><c:plotArea><c:layout/>
      <${chartTag}>${extraOpts}${sersXml}${axisIds}</${chartTag}>${axesXml}
    </c:plotArea></c:chart>
  </c:chartSpace>`;
}

function addChartToSlide(deck, slideIdx) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);

  const chartPath = uniquePath(pk, 'ppt/charts', 'chart', 'xml');
  const initCats = currentLanguage === 'en' ? ['Q1', 'Q2', 'Q3', 'Q4'] : ['الربع 1', 'الربع 2', 'الربع 3', 'الربع 4'];
  const initSeries = [{ name: currentLanguage === 'en' ? 'Sales' : 'المبيعات', vals: [45, 72, 60, 95] }];
  pk.setText(chartPath, buildChartXmlDoc('bar', initCats, initSeries));
  setOverride(pk, '/' + chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
  const rid = addRel(pk, slPath, REL + 'chart', chartPath);

  const w = Math.round(deck.W * 0.52), h = Math.round(deck.H * 0.48);
  const x = Math.round((deck.W - w) / 2), y = Math.round((deck.H - h) / 2);

  const xml = `<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}">
    <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart r:id="${rid}"/>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addChart');
  renderPptxEditor();
  toast('تمت إضافة الرسم البياني (عدّل أرقامه من القائمة الجانبية)');
}

function updateChartXml(chartPath, type, cats, seriesList) {
  if (!chartPath) return;
  pk.setText(chartPath, buildChartXmlDoc(type, cats, seriesList));
}

function addImageToSlide(deck, slideIdx, imgInfo, customGeom) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);

  const mediaPath = uniquePath(pk, 'ppt/media', 'image', imgInfo.ext);
  pk.setBin(mediaPath, imgInfo.bytes);
  ensureDefault(pk, imgInfo.ext, CT_IMG[imgInfo.ext] || 'image/png');
  const rid = addRel(pk, slPath, REL + 'image', mediaPath);

  let x, y, w, h;
  if (customGeom) {
    ({ x, y, w, h } = customGeom);
  } else {
    const maxW = deck.W * 0.45, maxH = deck.H * 0.45;
    const ratio = (imgInfo.w || 300) / (imgInfo.h || 200);
    w = maxW;
    h = w / ratio;
    if (h > maxH) { h = maxH; w = h * ratio; }
    w = Math.round(w); h = Math.round(h);
    x = Math.round((deck.W - w) / 2); y = Math.round((deck.H - h) / 2);
  }

  const xml = `<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">
    <p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    </p:spPr>
  </p:pic>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addImage');
  renderPptxEditor();
  toast('تمت إضافة الصورة');
}

function replaceSelectedImage(slPath, sh, imgInfo) {
  const mediaPath = uniquePath(pk, 'ppt/media', 'image', imgInfo.ext);
  pk.setBin(mediaPath, imgInfo.bytes);
  ensureDefault(pk, imgInfo.ext, CT_IMG[imgInfo.ext] || 'image/png');
  const rid = addRel(pk, slPath, REL + 'image', mediaPath);
  const blip = desc(kid(sh, P, 'blipFill'), A, 'blip')[0];
  if (blip) blip.setAttributeNS(R, 'r:embed', rid);
  pushHistory('replaceImage');
  renderPptxEditor();
  toast('تم تبديل الصورة بنجاح');
}

function setPicCropPreset(sh, prst) {
  const spPr = kid(sh, P, 'spPr');
  if (!spPr) return;
  let geom = kid(spPr, A, 'prstGeom');
  if (!geom) {
    geom = mk(sh.ownerDocument, A, 'a:prstGeom', { prst });
    geom.appendChild(mk(sh.ownerDocument, A, 'a:avLst'));
    spPr.appendChild(geom);
  } else {
    geom.setAttribute('prst', prst);
  }
}

function triggerBatchImagesToSlides(deck) {
  const inp = $('#batchImgIn');
  if (!inp) return;
  inp.value = '';
  inp.onchange = async () => {
    const files = Array.from(inp.files || []);
    if (!files.length) return;
    busy(true, `جارٍ تحويل ${files.length} صور إلى شرائح…`);
    try {
      const layouts = availableLayouts();
      const blankLayout = layouts.find(l => /blank|فارغة/i.test(l.name)) || layouts[layouts.length - 1] || layouts[0];
      for (const f of files) {
        const info = await prepImage(f);
        const newPath = uniquePath(pk, 'ppt/slides', 'slide', 'xml');
        pk.setText(newPath, XMLDECL + `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
          </p:spTree></p:cSld>
        </p:sld>`);
        if (blankLayout) addRel(pk, newPath, REL + 'slideLayout', blankLayout.path);
        registerSlidePart(deck, newPath);
        const freshDeck = readDeck();
        addImageToSlide(freshDeck, freshDeck.slides.length - 1, info, { x: 0, y: 0, w: freshDeck.W, h: freshDeck.H });
      }
      pushHistory('batchImages');
      renderPptxEditor();
      toast(`تم تحويل ${files.length} صور إلى شرائح بنجاح`);
    } catch (err) {
      toast(err instanceof UserError ? err.message : 'تعذر تحميل الصورة');
    } finally {
      busy(false);
    }
  };
  inp.click();
}

function addMediaToSlide(deck, slideIdx, mediaInfo) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const id = nextShapeId(spTree);

  const mediaPath = uniquePath(pk, 'ppt/media', 'media', mediaInfo.ext);
  pk.setBin(mediaPath, mediaInfo.bytes);
  ensureDefault(pk, mediaInfo.ext, mediaInfo.mime);
  const mRid = addRel(pk, slPath, REL + (mediaInfo.isAudio ? 'audio' : 'video'), mediaPath);
  const mediaRid = addRel(pk, slPath, REL + 'media', mediaPath);

  const posterCanvas = document.createElement('canvas');
  posterCanvas.width = 320; posterCanvas.height = 180;
  const pCtx = posterCanvas.getContext('2d');
  pCtx.fillStyle = '#1E293B'; pCtx.fillRect(0, 0, 320, 180);
  pCtx.fillStyle = '#FFFFFF'; pCtx.font = 'bold 28px sans-serif'; pCtx.textAlign = 'center';
  pCtx.fillText(mediaInfo.isAudio ? '🔊 Audio' : '▶ Video', 160, 98);
  const dataUrl = posterCanvas.toDataURL('image/png');
  const binStr = atob(dataUrl.split(',')[1]);
  const posterBytes = new Uint8Array(binStr.length);
  for (let i = 0; i < binStr.length; i++) posterBytes[i] = binStr.charCodeAt(i);
  const posterPath = uniquePath(pk, 'ppt/media', 'poster', 'png');
  pk.setBin(posterPath, posterBytes);
  ensureDefault(pk, 'png', 'image/png');
  const pRid = addRel(pk, slPath, REL + 'image', posterPath);

  const w = Math.round(deck.W * 0.42), h = mediaInfo.isAudio ? Math.round(deck.H * 0.12) : Math.round(deck.H * 0.35);
  const x = Math.round((deck.W - w) / 2), y = Math.round((deck.H - h) / 2);
  const tag = mediaInfo.isAudio ? 'a:audioFile' : 'a:videoFile';

  const xml = `<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}" xmlns:p14="${NS.p14}">
    <p:nvPicPr><p:cNvPr id="${id}" name="Media ${id}"/><p:cNvPicPr/><p:nvPr><${tag} r:link="${mRid}"/><p:extLst><p:ext uri="{DAA4B4D4-2A2A-4C6A-9F4A-4D3F6A7A6F0D}"><p14:media r:embed="${mediaRid}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>
    <p:blipFill><a:blip r:embed="${pRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  </p:pic>`;

  spTree.appendChild(frag(doc, xml));
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('addMedia');
  renderPptxEditor();
  toast('تم إدراج الوسائط بنجاح');
}

/* ── النسخ واللصق والتكرار ونسخ التنسيق والحذف والترتيب والتجميع والمحاذاة ── */
function copySelectedShape(deck, slideIdx) {
  if (curPptx.selShapeIdx == null) return;
  const slPath = deck.slides[slideIdx].path;
  const sh = shapesOf(pk.doc(slPath))[curPptx.selShapeIdx];
  if (!sh) return;
  curPptx.clipboardXml = new XMLSerializer().serializeToString(sh);
  curPptx.clipboardSourcePath = slPath;
  curPptx.clipboardRels = Array.from(readRels(pk, slPath).values());
  refreshPptxRibbon(deck, slideIdx);
  toast('تم نسخ العنصر');
}

function pasteShapeToSlide(deck, slideIdx) {
  if (!curPptx.clipboardXml) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const d = new DOMParser().parseFromString(curPptx.clipboardXml, 'application/xml');
  const node = doc.importNode(d.documentElement, true);

  const newId = nextShapeId(spTree);
  const cNvPr = kid(nvOf(node), P, 'cNvPr');
  if (cNvPr) cNvPr.setAttribute('id', String(newId));

  const g = readXfrm(xfrmOf(node));
  if (g) {
    g.x += 228600;
    g.y += 228600;
    writeGeomToShape(node, g);
  }

  const relMap = new Map();
  for (const rel of curPptx.clipboardRels || []) {
    if (rel.type.endsWith('/notesSlide')) continue;
    let newRid = null;
    if (rel.type.endsWith('/chart') && rel.path && pk.has(rel.path)) {
      const chartPath = uniquePath(pk, 'ppt/charts', 'chart', 'xml');
      const chartData = pk.files.get(rel.path);
      if (typeof chartData === 'string') pk.setText(chartPath, chartData);
      else pk.setBin(chartPath, chartData);
      const sourceRelsPath = relsPathOf(rel.path), targetRelsPath = relsPathOf(chartPath);
      if (pk.has(sourceRelsPath)) pk.setText(targetRelsPath, pk.files.get(sourceRelsPath));
      const sourceOverride = kids(ctRoot(pk), null, 'Override').find(x => x.getAttribute('PartName') === '/' + rel.path);
      if (sourceOverride) setOverride(pk, '/' + chartPath, sourceOverride.getAttribute('ContentType'));
      newRid = addRel(pk, slPath, rel.type, chartPath);
    } else if (rel.mode === 'External') {
      newRid = addRel(pk, slPath, rel.type, rel.target, true);
    } else if (rel.path && pk.has(rel.path)) {
      newRid = addRel(pk, slPath, rel.type, rel.path);
    }
    if (newRid) relMap.set(rel.id, newRid);
  }
  const walker = doc.createTreeWalker(node, NodeFilter.SHOW_ELEMENT);
  const relNodes = [node]; while (walker.nextNode()) relNodes.push(walker.currentNode);
  for (const el of relNodes) {
    for (const attr of ['id', 'embed', 'link']) {
      const oldRid = el.getAttributeNS(R, attr);
      if (oldRid && relMap.has(oldRid)) el.setAttributeNS(R, 'r:' + attr, relMap.get(oldRid));
    }
  }

  spTree.appendChild(node);
  curPptx.selShapeIdx = shapesOf(doc).length - 1;
  curPptx.multiSel.clear();
  pushHistory('pasteShape');
  renderPptxEditor();
  toast('تم لصق العنصر');
}

function duplicateSelectedShape(deck, slideIdx) {
  copySelectedShape(deck, slideIdx);
  pasteShapeToSlide(deck, slideIdx);
  toast('تم تكرار العنصر');
}

function startFormatPainter(deck, slideIdx) {
  if (curPptx.copiedStyle) {
    curPptx.copiedStyle = null;
    renderPptxEditor();
    return;
  }
  if (curPptx.selShapeIdx == null) return;
  const slPath = deck.slides[slideIdx].path;
  const sh = shapesOf(pk.doc(slPath))[curPptx.selShapeIdx];
  if (!sh) return;
  const ctx = buildCtx(deck, slPath);
  curPptx.copiedStyle = readShapeAppearance(sh, ctx);
  renderPptxEditor();
  toast('انقر الآن على العنصر الهدف لتطبيق التنسيق عليه');
}

function applyFormatPainterToShape(deck, slideIdx, targetIdx) {
  const appSt = curPptx.copiedStyle;
  curPptx.copiedStyle = null;
  if (!appSt) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const sh = shapesOf(doc)[targetIdx];
  if (!sh) return;

  const spPr = kid(sh, P, 'spPr');
  if (spPr && sh.localName !== 'pic') {
    for (const f of kids(spPr, A)) if (['solidFill', 'noFill', 'gradFill'].includes(f.localName)) removeNode(f);
    let fEl;
    if (appSt.noFill) fEl = mk(doc, A, 'a:noFill');
    else if (appSt.fillType === 'grad') fEl = buildGradFillXml(doc, appSt.fillHex, appSt.gradHex2, appSt.gradAngle, appSt.opacity);
    else {
      fEl = mk(doc, A, 'a:solidFill');
      fEl.appendChild(mk(doc, A, 'a:srgbClr', { val: appSt.fillHex.replace('#', '') }));
    }
    spPr.insertBefore(fEl, kid(spPr, A, 'prstGeom') || null);
  }
  if (spPr) {
    const oldLn = kid(spPr, A, 'ln');
    if (oldLn) removeNode(oldLn);
    const lnEl = mk(doc, A, 'a:ln', { w: String(appSt.borderW * 12700) });
    if (appSt.borderW > 0) {
      const sf = mk(doc, A, 'a:solidFill');
      sf.appendChild(mk(doc, A, 'a:srgbClr', { val: appSt.borderHex.replace('#', '') }));
      lnEl.appendChild(sf);
    } else lnEl.appendChild(mk(doc, A, 'a:noFill'));
    spPr.appendChild(lnEl);
  }

  curPptx.selShapeIdx = targetIdx;
  pushHistory('formatPainter');
  renderPptxEditor();
  toast('تم نسخ التنسيق بنجاح');
}

function deleteSelectedShape(deck, slideIdx) {
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const all = shapesOf(doc);
  const toDel = curPptx.multiSel.size ? Array.from(curPptx.multiSel) : (curPptx.selShapeIdx != null ? [curPptx.selShapeIdx] : []);
  if (!toDel.length) return;
  toDel.sort((a, b) => b - a).forEach(i => { if (all[i]) removeNode(all[i]); });
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  curPptx.selCell = null;
  pushHistory('deleteShape');
  renderPptxEditor();
  toast('تم حذف العنصر');
}

function reorderShapeLayer(deck, slideIdx, dir) {
  if (curPptx.selShapeIdx == null) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const all = shapesOf(doc);
  const sh = all[curPptx.selShapeIdx];
  if (!sh) return;

  if (dir === 'front') {
    spTree.appendChild(sh);
    curPptx.selShapeIdx = shapesOf(doc).length - 1;
  } else {
    const firstSh = all[0];
    if (firstSh && firstSh !== sh) spTree.insertBefore(sh, firstSh);
    curPptx.selShapeIdx = 0;
  }
  pushHistory('reorder:' + dir);
  renderPptxEditor();
  toast(dir === 'front' ? 'تم تقديم العنصر للأمام' : 'تم إرسال العنصر للخلف');
}

function groupSelectedShapes(deck, slideIdx) {
  if (curPptx.multiSel.size < 2) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const ctx = buildCtx(deck, slPath);
  const all = shapesOf(doc);
  const indices = Array.from(curPptx.multiSel).sort((a, b) => a - b);
  const items = indices.map(i => all[i]).filter(Boolean);
  if (items.length < 2) return;

  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  items.forEach(s => {
    const g = geomOf(s, ctx);
    if (!g) return;
    minX = Math.min(minX, g.x); minY = Math.min(minY, g.y);
    maxX = Math.max(maxX, g.x + g.w); maxY = Math.max(maxY, g.y + g.h);
  });
  if (!isFinite(minX)) return;

  const grpId = nextShapeId(spTree);
  const grp = frag(doc, `<p:grpSp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvGrpSpPr><p:cNvPr id="${grpId}" name="Group ${grpId}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm>
      <a:off x="${minX}" y="${minY}"/><a:ext cx="${maxX - minX}" cy="${maxY - minY}"/>
      <a:chOff x="${minX}" y="${minY}"/><a:chExt cx="${maxX - minX}" cy="${maxY - minY}"/>
    </a:xfrm></p:grpSpPr>
  </p:grpSp>`);

  spTree.insertBefore(grp, items[0]);
  items.forEach(s => grp.appendChild(s));

  curPptx.multiSel.clear();
  curPptx.selShapeIdx = shapesOf(doc).indexOf(grp);
  pushHistory('group');
  renderPptxEditor();
  toast('تم تجميع العناصر');
}

function ungroupSelectedShape(deck, slideIdx) {
  if (curPptx.selShapeIdx == null) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const grp = shapesOf(doc)[curPptx.selShapeIdx];
  if (!grp || grp.localName !== 'grpSp') return;

  const gg = readXfrm(xfrmOf(grp));
  const children = kids(grp, P).filter(e => SHAPE_TAGS.has(e.localName));
  children.forEach(ch => {
    const cg = readXfrm(xfrmOf(ch));
    if (gg && cg && gg.chw && gg.chh) {
      cg.x = Math.round(gg.x + ((cg.x - gg.chx) / gg.chw) * gg.w);
      cg.y = Math.round(gg.y + ((cg.y - gg.chy) / gg.chh) * gg.h);
      cg.w = Math.round((cg.w / gg.chw) * gg.w);
      cg.h = Math.round((cg.h / gg.chh) * gg.h);
      writeGeomToShape(ch, cg);
    }
    spTree.insertBefore(ch, grp);
  });
  removeNode(grp);
  curPptx.selShapeIdx = null;
  curPptx.multiSel.clear();
  pushHistory('ungroup');
  renderPptxEditor();
  toast('تم فك التجميع');
}

function centerSingleShapeToSlide(axis) {
  if (!curPptx.deck || curPptx.selShapeIdx == null) return;
  const deck = curPptx.deck;
  const doc = pk.doc(deck.slides[curPptx.slideIdx].path);
  const sh = shapesOf(doc)[curPptx.selShapeIdx];
  if (!sh || isShapeLocked(sh)) return;
  const g = readXfrm(xfrmOf(sh));
  if (!g) return;
  if (axis === 'h') g.x = Math.round((deck.W - g.w) / 2);
  if (axis === 'v') g.y = Math.round((deck.H - g.h) / 2);
  writeGeomToShape(sh, g);
  pushHistory('centerSlide:' + axis);
  renderPptxEditor();
}

function alignSelectedShapes(deck, slideIdx, mode) {
  if (curPptx.multiSel.size < 2) return;
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const ctx = buildCtx(deck, slPath);
  const all = shapesOf(doc);
  const list = Array.from(curPptx.multiSel).map(i => ({ sh: all[i], g: all[i] && geomOf(all[i], ctx) })).filter(x => x.sh && x.g);
  if (list.length < 2) return;

  const minX = Math.min(...list.map(x => x.g.x));
  const maxX = Math.max(...list.map(x => x.g.x + x.g.w));
  const minY = Math.min(...list.map(x => x.g.y));
  const maxY = Math.max(...list.map(x => x.g.y + x.g.h));
  const midX = (minX + maxX) / 2;
  const midY = (minY + maxY) / 2;

  for (const item of list) {
    const g = { ...item.g };
    if (mode === 'left') g.x = minX;
    else if (mode === 'center') g.x = Math.round(midX - g.w / 2);
    else if (mode === 'right') g.x = maxX - g.w;
    else if (mode === 'top') g.y = minY;
    else if (mode === 'middle') g.y = Math.round(midY - g.h / 2);
    else if (mode === 'bottom') g.y = maxY - g.h;
    writeGeomToShape(item.sh, g);
  }
  pushHistory('align:' + mode);
  renderPptxEditor();
  toast('تمت محاذاة العناصر');
}

function distributeSelectedShapes(deck, slideIdx, axis) {
  if (curPptx.multiSel.size < 3) {
    toast('حدد 3 عناصر على الأقل للتوزيع المتساوي');
    return;
  }
  const slPath = deck.slides[slideIdx].path;
  const doc = pk.doc(slPath);
  const ctx = buildCtx(deck, slPath);
  const all = shapesOf(doc);
  const list = Array.from(curPptx.multiSel).map(i => ({ sh: all[i], g: all[i] && geomOf(all[i], ctx) })).filter(x => x.sh && x.g);
  if (list.length < 3) return;

  if (axis === 'h') {
    list.sort((a, b) => a.g.x - b.g.x);
    const first = list[0].g.x;
    const last = list[list.length - 1].g.x + list[list.length - 1].g.w;
    const totalW = list.reduce((s, x) => s + x.g.w, 0);
    const gap = (last - first - totalW) / (list.length - 1);
    let curX = first;
    for (const item of list) {
      writeGeomToShape(item.sh, { ...item.g, x: Math.round(curX) });
      curX += item.g.w + gap;
    }
  } else {
    list.sort((a, b) => a.g.y - b.g.y);
    const first = list[0].g.y;
    const last = list[list.length - 1].g.y + list[list.length - 1].g.h;
    const totalH = list.reduce((s, x) => s + x.g.h, 0);
    const gap = (last - first - totalH) / (list.length - 1);
    let curY = first;
    for (const item of list) {
      writeGeomToShape(item.sh, { ...item.g, y: Math.round(curY) });
      curY += item.g.h + gap;
    }
  }
  pushHistory('distribute:' + axis);
  renderPptxEditor();
  toast('تم توزيع المسافات بالتساوي');
}

/* ── الشرائح الذكية وتحويل PPTX إلى Word ── */
function openSmartSlidesModal(deck, slideIdx) {
  const backdrop = h('div', { class: 'md-modal-backdrop', onclick: (e) => { if (e.target === backdrop) backdrop.remove(); } });
  const cards = [
    { id: 'compare', icon: '⚖️', label: 'مقارنة عمودين' },
    { id: 'stats', icon: '📈', label: 'أرقام وإحصائيات' },
    { id: 'timeline', icon: '🗓️', label: 'خط زمني' },
    { id: 'team', icon: '👥', label: 'فريق العمل' }
  ];
  const grid = h('div', { class: 'smart-grid' }, ...cards.map(c =>
    h('button', {
      type: 'button',
      class: 'smart-card',
      onclick: () => {
        backdrop.remove();
        insertSmartSlideTemplate(deck, slideIdx, c.id);
      }
    }, h('span', null, c.icon), tr(c.label))
  ));
  backdrop.append(h('div', { class: 'md-modal' },
    h('strong', null, tr('مكتبة الشرائح الذكية')),
    grid,
    h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, btn('إلغاء', () => backdrop.remove()))
  ));
  document.body.append(backdrop);
}

function insertSmartSlideTemplate(deck, slideIdx, tplId) {
  addSlide(deck);
  const freshDeck = readDeck();
  const targetIdx = Math.min(curPptx.slideIdx, freshDeck.slides.length - 1);
  const slPath = freshDeck.slides[targetIdx].path;
  const doc = pk.doc(slPath);
  const spTree = spTreeOf(doc);
  const W0 = freshDeck.W, H0 = freshDeck.H;

  const box = (id, x, y, w, h, bgHex, txt, sz, bold, txtHex = '1F1C17') => `<p:sp xmlns:p="${P}" xmlns:a="${A}">
    <p:nvSpPr><p:cNvPr id="${id}" name="Smart ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${Math.round(W0 * x)}" y="${Math.round(H0 * y)}"/><a:ext cx="${Math.round(W0 * w)}" cy="${Math.round(H0 * h)}"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      ${bgHex ? `<a:solidFill><a:srgbClr val="${bgHex}"/></a:solidFill>` : '<a:noFill/>'}
    </p:spPr>
    <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr sz="${sz * 100}"${bold ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${txtHex}"/></a:solidFill></a:rPr><a:t>${txt}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;

  let id = nextShapeId(spTree);
  if (tplId === 'compare') {
    spTree.appendChild(frag(doc, box(id++, 0.08, 0.08, 0.84, 0.12, null, 'عنوان المقارنة', 28, true, 'B4531A')));
    spTree.appendChild(frag(doc, box(id++, 0.08, 0.25, 0.4, 0.62, 'FFF3E8', 'الخيار الأول\n• ميزة 1\n• ميزة 2', 20, false)));
    spTree.appendChild(frag(doc, box(id++, 0.52, 0.25, 0.4, 0.62, 'EFF6FF', 'الخيار الثاني\n• ميزة 1\n• ميزة 2', 20, false)));
  } else if (tplId === 'stats') {
    spTree.appendChild(frag(doc, box(id++, 0.08, 0.08, 0.84, 0.12, null, 'أهم الأرقام والإحصائيات', 28, true, 'B4531A')));
    [['+95%', 'نسبة الرضا'], ['120K', 'مستخدم نشط'], ['24/7', 'دعم متواصل']].forEach(([num, lbl], i) => {
      spTree.appendChild(frag(doc, box(id++, 0.08 + i * 0.29, 0.3, 0.26, 0.45, 'F8FAFC', `${num}\n${lbl}`, 26, true, '0F2547')));
    });
  } else if (tplId === 'timeline') {
    spTree.appendChild(frag(doc, box(id++, 0.08, 0.08, 0.84, 0.12, null, 'الخط الزمني للمشروع', 28, true, 'B4531A')));
    ['المرحلة 1: التخطيط', 'المرحلة 2: التنفيذ', 'المرحلة 3: الإطلاق'].forEach((st, i) => {
      spTree.appendChild(frag(doc, box(id++, 0.08 + i * 0.29, 0.35, 0.26, 0.32, 'B4531A', st, 18, true, 'FFFFFF')));
    });
  } else {
    spTree.appendChild(frag(doc, box(id++, 0.08, 0.08, 0.84, 0.12, null, 'فريق العمل', 28, true, 'B4531A')));
    ['الاسم الأول\nالمدير التنفيذي', 'الاسم الثاني\nمدير التقنية', 'الاسم الثالث\nمدير التصميم'].forEach((m, i) => {
      spTree.appendChild(frag(doc, box(id++, 0.08 + i * 0.29, 0.28, 0.26, 0.5, 'FAF7F1', m, 18, true)));
    });
  }

  pushHistory('smartSlide');
  renderPptxEditor();
  toast('تم إدراج الشريحة الذكية');
}

function exportPptxToDocx(deck) {
  const newPk = createNewBlankDocx('blank');
  const doc = newPk.doc('word/document.xml');
  const body = kid(doc.documentElement, W, 'body');
  const sectPr = kid(body, W, 'sectPr');
  for (const child of Array.from(kids(body, W))) if (child.localName !== 'sectPr') removeNode(child);

  deck.slides.forEach((sl, idx) => {
    const sDoc = pk.doc(sl.path);
    const titleP = frag(doc, `<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="Heading2"/><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="B4531A"/><w:sz w:val="32"/></w:rPr><w:t>${currentLanguage === 'en' ? 'Slide' : 'شريحة'} ${idx + 1}</w:t></w:r></w:p>`);
    body.insertBefore(titleP, sectPr || null);

    for (const sh of shapesOf(sDoc)) {
      const txt = getShapePlainText(sh).trim();
      if (!txt) continue;
      for (const line of txt.split(/\r?\n/)) {
        if (!line.trim()) continue;
        const p = frag(doc, `<w:p xmlns:w="${W}"><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${escXml(line.trim())}</w:t></w:r></w:p>`);
        body.insertBefore(p, sectPr || null);
      }
    }
  });

  pk = newPk;
  activeDraftId = null;
  kind = 'docx';
  fileName = fileName.replace(/\.(pptx|pptm)$/i, '') + '.docx';
  $('#fname').textContent = fileName;
  curDocx = defaultDocxState();
  resetHistory();
  pushHistory('pptxToDocx');
  renderEditor();
  toast('تم تصدير العرض إلى مستند Word');
}
