'use strict';
/* ══════════════════════════════════════════════════════════════
   brivatedocs — v2 UI Enhancements (loaded only by index-v2.html)
   1. PPTX Contextual Slide-Over Inspector (auto-opens on selection)
   2. Uniform Lucide-style SVG icon replacer for Editor Topbar,
      PPTX Ribbon, Word Toolbar, and PDF Toolbar (i18n-safe)
   ══════════════════════════════════════════════════════════════ */
(function () {
  if (!document.body || !document.body.classList.contains('v2-body')) return;

  /* ── Crisp 1.75px-stroke SVG Icon Library ── */
  const I = (d) =>
    `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.85" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;

  const ICON_MAP = {
    '⬅': I('<line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>'),
    '⬇': I('<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>'),
    '⬆': I('<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>'),
    '⬆️': I('<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>'),
    '⬇️': I('<line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>'),
    '🔎': I('<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>'),
    '🔍': I('<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>'),
    '↩': I('<polyline points="9 14 4 9 9 4"/><path d="M20 20v-7a4 4 0 0 0-4-4H4"/>'),
    '↪': I('<polyline points="15 14 20 9 15 4"/><path d="M4 20v-7a4 4 0 0 1 4-4h12"/>'),
    '➕': I('<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>'),
    '📄': I('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>'),
    '📝': I('<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>'),
    '👁': I('<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>'),
    '🗑': I('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>'),
    '📋': I('<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>'),
    '📥': I('<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>'),
    '🖌': I('<path d="M18.37 2.63L14 7l-1.59-1.59a2 2 0 0 0-2.82 0L8 7l9 9 1.59-1.59a2 2 0 0 0 0-2.82L17 10l4.37-4.37a2.12 2.12 0 1 0-3-3z"/><path d="M9 8c-2 3-4 3.5-7 4l8 10c2-1 6-5 6-7"/>'),
    '🔗': I('<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'),
    '🔓': I('<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>'),
    '🔒': I('<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>'),
    '🖼': I('<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>'),
    '🖼️': I('<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>'),
    '🖼+': I('<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>'),
    '▦': I('<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/>'),
    '📊': I('<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>'),
    '🎬': I('<rect x="2" y="2" width="20" height="20" rx="2.18"/><line x1="7" y1="2" x2="7" y2="22"/><line x1="17" y1="2" x2="17" y2="22"/><line x1="2" y1="12" x2="22" y2="12"/>'),
    '✨': I('<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>'),
    '📐': I('<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/>'),
    '🎨': I('<circle cx="13.5" cy="6.5" r=".5"/><circle cx="17.5" cy="10.5" r=".5"/><circle cx="8.5" cy="7.5" r=".5"/><circle cx="6.5" cy="12.5" r=".5"/><path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10c.93 0 1.5-.67 1.5-1.5 0-.39-.15-.74-.39-1.01-.23-.26-.38-.61-.38-.99 0-.83.67-1.5 1.5-1.5H16c2.76 0 5-2.24 5-5 0-4.42-4.03-8-9-8z"/>'),
    '🌈': I('<path d="M22 17a10 10 0 0 0-20 0"/><path d="M18 17a6 6 0 0 0-12 0"/><path d="M14 17a2 2 0 0 0-4 0"/>'),
    '🔢': I('<line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/>'),
    '▶': I('<polygon points="5 3 19 12 5 21 5 3"/>'),
    '🖥': I('<rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>'),
    '📦': I('<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>'),
    '⊞': I('<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/>'),
    '🎯': I('<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>'),
    '🕒': I('<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'),
    '🖨': I('<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>'),
    '🖨️': I('<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>'),
    '✂️': I('<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/>'),
    '✂': I('<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/>'),
    '🗜': I('<polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="14" y1="10" x2="21" y2="3"/><line x1="3" y1="21" x2="10" y2="14"/>'),
    '✍': I('<path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/>'),
    '✍️': I('<path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/>'),
    '✏️': I('<path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"/>'),
    '✏': I('<path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"/>'),
    '🖍': I('<path d="M9 11l-6 6v3h9l3-3"/><path d="M22 12l-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>'),
    '🖍️': I('<path d="M9 11l-6 6v3h9l3-3"/><path d="M22 12l-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>'),
    '🧹': I('<path d="M20 20H7L3 16c-.8-.8-.8-2 0-2.8L13.8 2.4c.8-.8 2-.8 2.8 0L21.6 7.4c.8.8.8 2 0 2.8L12 20"/>'),
    '📑': I('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>'),
    '💬': I('<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>'),
    '☑': I('<polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>'),
    '📅': I('<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>'),
    '💧': I('<path d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"/>'),
    '📌': I('<path d="M12 17v5"/><path d="M9 10.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24V17h14v-1.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V6h1a2 2 0 0 0 0-4H8a2 2 0 0 0 0 4h1v4.76z"/>'),
    '↺': I('<polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>'),
    '↻': I('<polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>'),
    '▭': I('<rect x="3" y="6" width="18" height="12" rx="2"/>'),
    '◯': I('<circle cx="12" cy="12" r="9"/>'),
    '▲': I('<polygon points="12 4 21 20 3 20"/>'),
    '◆': I('<polygon points="12 2 22 12 12 22 2 12"/>'),
    '⬡': I('<polygon points="12 2 21 7 21 17 12 22 3 17 3 7"/>'),
    '★': I('<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>'),
    '➡': I('<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>'),
    '╱': I('<line x1="5" y1="19" x2="19" y2="5"/>'),
    '🌙': I('<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>'),
    '☀️': I('<circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>'),
    '🛡': I('<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>'),
    '🛡️': I('<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>'),
    '⬛': I('<rect x="3" y="3" width="18" height="18" rx="2" fill="currentColor" fill-opacity="0.22"/><line x1="7" y1="9" x2="17" y2="9"/><line x1="7" y1="15" x2="17" y2="15"/>'),
    '💾': I('<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>'),
    '⌘': I('<path d="M18 3a3 3 0 0 0-3 3v12a3 3 0 0 0 3 3 3 3 0 0 0 3-3 3 3 0 0 0-3-3H6a3 3 0 0 0-3 3 3 3 0 0 0 3 3 3 3 0 0 0 3-3V6a3 3 0 0 0-3-3 3 3 0 0 0-3 3 3 3 0 0 0 3 3h12a3 3 0 0 0 3-3 3 3 0 0 0-3-3z"/>')
  };

  const ICON_KEYS = Object.keys(ICON_MAP).sort((a, b) => b.length - a.length);
  const TARGET_SELECTOR = '#editor .top button.btn, #app .ribbon button.btn, #app .tb2 button.btn, #app .doc-ctx-bar button.btn, #app .pdf-toolbar button.btn, #app .notes-panel summary, #themeToggleHome';

  let enhancing = false;

  function findLeadingToken(text) {
    const trimmed = text.trim();
    for (const k of ICON_KEYS) {
      if (trimmed === k) return { token: k, rest: '' };
      if (trimmed.startsWith(k + ' ')) return { token: k, rest: trimmed.slice(k.length + 1).trim() };
    }
    return null;
  }

  function enhanceButton(el) {
    if (!el || el.classList.contains('v2-no-icon')) return;
    if (el.querySelector(':scope > .v2-ic')) return;
    // Only enhance buttons with a single text node (or simple text)
    if (el.children.length > 0) return;
    const raw = (el.textContent || '').trim();
    if (!raw) return;
    const match = findLeadingToken(raw);
    if (!match) return;

    el.dataset.v2RawText = raw;
    el.textContent = '';
    const ic = document.createElement('span');
    ic.className = 'v2-ic';
    ic.setAttribute('aria-hidden', 'true');
    ic.innerHTML = ICON_MAP[match.token];
    el.appendChild(ic);
    if (match.rest) {
      const lbl = document.createElement('span');
      lbl.className = 'v2-lbl';
      lbl.textContent = match.rest;
      el.appendChild(lbl);
    }
  }

  function restoreEnhanced(root) {
    const els = (root || document).querySelectorAll('[data-v2-raw-text]');
    for (const el of els) {
      el.textContent = el.dataset.v2RawText;
      delete el.dataset.v2RawText;
    }
  }

  function enhanceAll(root) {
    if (enhancing) return;
    enhancing = true;
    try {
      const scope = root || document;
      const nodes = scope.querySelectorAll ? scope.querySelectorAll(TARGET_SELECTOR) : [];
      for (const el of nodes) enhanceButton(el);
    } finally {
      enhancing = false;
    }
  }

  /* ── Keep applyTheme & applyLanguage 100% synced with SVG icons ── */
  if (typeof window.applyTheme === 'function') {
    const origApplyTheme = window.applyTheme;
    window.applyTheme = function (theme) {
      for (const b of [document.getElementById('themeToggleHome'), document.getElementById('themeToggleEditor')]) {
        if (b && b.dataset.v2RawText) delete b.dataset.v2RawText;
      }
      const res = origApplyTheme.apply(this, arguments);
      enhanceAll(document);
      return res;
    };
  }

  if (typeof window.applyLanguage === 'function') {
    const origApplyLanguage = window.applyLanguage;
    window.applyLanguage = function () {
      enhancing = true;
      try { restoreEnhanced(document); } finally { enhancing = false; }
      const res = origApplyLanguage.apply(this, arguments);
      enhanceAll(document);
      return res;
    };
  }

  /* ── PPTX Contextual Slide-Over Inspector (Option 1) ── */
  let userClosedInspector = false;

  function ensureInspectorChrome(propsPanel) {
    if (!propsPanel) return;
    const ppt = propsPanel.closest('.ppt');
    if (!ppt) return;

    // Add close button inside inspector header once
    if (!propsPanel.querySelector(':scope > .v2-props-close')) {
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'v2-props-close';
      closeBtn.title = 'Close';
      closeBtn.setAttribute('aria-label', 'Close Inspector');
      closeBtn.innerHTML = I('<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>');
      closeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        userClosedInspector = true;
        propsPanel.classList.remove('v2-open');
        syncInspectorFab(ppt, propsPanel);
      });
      propsPanel.prepend(closeBtn);
    }

    // Add subtle floating toggle pill in stagewrap when closed
    const stageWrap = ppt.querySelector('.stagewrap');
    if (stageWrap && !stageWrap.querySelector(':scope > .v2-props-fab')) {
      const fab = document.createElement('button');
      fab.type = 'button';
      fab.className = 'v2-props-fab';
      fab.innerHTML = I('<line x1="4" y1="21" x2="4" y2="14"/><line x1="4" y1="10" x2="4" y2="3"/><line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/><line x1="20" y1="21" x2="20" y2="16"/><line x1="20" y1="12" x2="20" y2="3"/><line x1="1" y1="14" x2="7" y2="14"/><line x1="9" y1="8" x2="15" y2="8"/><line x1="17" y1="16" x2="23" y2="16"/>') + '<span>Inspector</span>';
      fab.addEventListener('click', (e) => {
        e.stopPropagation();
        userClosedInspector = false;
        propsPanel.classList.add('v2-open');
        syncInspectorFab(ppt, propsPanel);
      });
      stageWrap.appendChild(fab);
    }
  }

  function syncInspectorFab(ppt, propsPanel) {
    const fab = ppt?.querySelector('.v2-props-fab');
    if (!fab) return;
    fab.hidden = propsPanel.classList.contains('v2-open');
  }

  function syncPptxInspectorState(forceOpenOnSelect) {
    const propsPanel = document.querySelector('#app .ppt > .props');
    if (!propsPanel) return;
    const ppt = propsPanel.closest('.ppt');
    ensureInspectorChrome(propsPanel);

    const hasSelection =
      typeof curPptx !== 'undefined' &&
      curPptx &&
      (curPptx.selShapeIdx != null || (curPptx.multiSel && curPptx.multiSel.size > 0));

    if (hasSelection) {
      if (forceOpenOnSelect) userClosedInspector = false;
      propsPanel.classList.toggle('v2-open', !userClosedInspector);
    } else {
      userClosedInspector = false;
      propsPanel.classList.remove('v2-open');
    }
    syncInspectorFab(ppt, propsPanel);
  }

  if (typeof window.renderShapeProps === 'function') {
    const origRenderShapeProps = window.renderShapeProps;
    let prevSelKey = '';
    window.renderShapeProps = function () {
      const res = origRenderShapeProps.apply(this, arguments);
      const selKey =
        typeof curPptx !== 'undefined' && curPptx
          ? `${curPptx.slideIdx}:${curPptx.selShapeIdx}:${curPptx.multiSel ? curPptx.multiSel.size : 0}`
          : '';
      const selectionChanged = selKey !== prevSelKey;
      prevSelKey = selKey;
      syncPptxInspectorState(selectionChanged);
      enhanceAll(document.querySelector('#app'));
      return res;
    };
  }

  /* ── Observe dynamic editor renders (PPTX, DOCX, PDF) ── */
  const observer = new MutationObserver(() => {
    if (enhancing) return;
    enhanceAll(document.getElementById('editor'));
    if (document.querySelector('#app .ppt > .props')) {
      syncPptxInspectorState(false);
    }
  });

  const editorRoot = document.getElementById('editor');
  if (editorRoot) {
    observer.observe(editorRoot, { childList: true, subtree: true });
  }

  // Initial pass
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => enhanceAll(document));
  } else {
    enhanceAll(document);
  }
})();
