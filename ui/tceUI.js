/**
 * Teams Chat Exporter - shared injected UI helpers (panels, toasts, menus).
 *
 * This is a classic (non-module) script that is loaded in TWO worlds:
 *   1. the content-script isolated world (listed before content.js in the
 *      manifest's content_scripts), used by content.js;
 *   2. the page MAIN world (injected with <script src> by content.js), used by
 *      videoDownload/coordinator.js and videoDownload/manifestDownload.js.
 * The two copies share nothing but the DOM: both find-or-create the same
 * #tce-stack container, and all styling comes from transcriptStyles.css
 * (content-script CSS applies to every element in the document). Every class
 * is prefixed `tce-` and all CSS custom properties are declared on our own
 * root elements, so nothing leaks into the host page.
 */
(function () {
  'use strict';
  if (globalThis.__tceUI) return;

  const STACK_ID = 'tce-stack';

  // 20x20 outline icons, stroke = currentColor.
  const ICONS = {
    copy: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><rect x="6" y="6" width="10" height="12" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M4 14V4.5A1.5 1.5 0 0 1 5.5 3H12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    download: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M10 3v10m0 0l-3.5-3.5M10 13l3.5-3.5M4 16h12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    batch: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><rect x="4" y="5" width="10" height="12" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M7 3h7.5A1.5 1.5 0 0 1 16 4.5V14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><path d="M7 10h4M7 13h2" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    video: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><rect x="2" y="5" width="11" height="10" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M13 9l5-3v8l-5-3" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>',
    chevron: '<svg class="tce-chevron" viewBox="0 0 8 5" aria-hidden="true" focusable="false"><path d="M1 1l3 3 3-3" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    close: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M5 5l10 10M15 5L5 15" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    check: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M4 10.5l4 4 8-9" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    error: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><circle cx="10" cy="10" r="7.25" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 6v5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="10" cy="13.75" r="1" fill="currentColor"/></svg>',
    info: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><circle cx="10" cy="10" r="7.25" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 9v5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="10" cy="6.25" r="1" fill="currentColor"/></svg>',
    zip: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M5 2.75h6.5L15 6.25v11H5z" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M9 4.5h1.5M8 6.5h1.5M9 8.5h1.5M8 10.5h1.5" stroke="currentColor" stroke-width="1.25" stroke-linecap="round"/><rect x="8" y="12.25" width="3" height="2.5" rx=".5" fill="none" stroke="currentColor" stroke-width="1.25"/></svg>',
    stop: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><rect x="5.5" y="5.5" width="9" height="9" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>',
    play: '<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M6.5 4.5v11l9-5.5z" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>'
  };

  // ---------- theme detection ----------

  const parseRGB = (str) => {
    const m = /rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+%?))?/i.exec(str || '');
    if (!m) return null;
    let a = m[4] === undefined ? 1 : parseFloat(m[4]);
    if (m[4] && m[4].endsWith('%')) a /= 100;
    return { r: +m[1], g: +m[2], b: +m[3], a };
  };
  const luminance = ({ r, g, b }) => {
    const ch = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b);
  };

  /**
   * Returns 'dark' or 'light' for the host page.
   * 1. Teams/SharePoint theme hints on <html>/<body> (class names, data-theme...)
   * 2. computed background luminance of the page shell
   * 3. prefers-color-scheme
   */
  const detectTheme = () => {
    const roots = [document.documentElement, document.body].filter(Boolean);
    for (const node of roots) {
      const hints = [
        typeof node.className === 'string' ? node.className : '',
        node.getAttribute('data-theme') || '',
        node.getAttribute('data-app-theme') || '',
        node.getAttribute('data-color-scheme') || ''
      ].join(' ');
      if (/(^|[\s_-])(theme-)?(dark|darkv2|contrast|high-?contrast|black)(\b|[\s_-]|$)/i.test(hints)) return 'dark';
      if (/(^|[\s_-])theme-(default|light)(\b|$)|(^|\s)light(\s|$)/i.test(hints)) return 'light';
    }
    const candidates = [
      document.body,
      document.documentElement,
      document.getElementById('app'),
      document.querySelector('[role="main"]')
    ];
    for (const node of candidates) {
      if (!node) continue;
      const c = parseRGB(getComputedStyle(node).backgroundColor);
      if (c && c.a > 0.5) return luminance(c) < 0.35 ? 'dark' : 'light';
    }
    try {
      return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    } catch (e) {
      return 'light';
    }
  };

  const applyTheme = (el) => {
    if (el) el.setAttribute('data-tce-theme', detectTheme());
    return el;
  };

  // ---------- DOM helpers ----------

  let uidCounter = 0;
  const uid = (prefix) => `tce-${prefix}-${Math.random().toString(36).slice(2, 7)}${++uidCounter}`;

  /** Tiny element builder: h('div', {class: 'x', onclick: fn}, [children]) */
  const h = (tag, attrs, children) => {
    const el = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (v === undefined || v === null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k === 'html') el.innerHTML = v;
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : String(v));
      }
    }
    if (children) {
      for (const c of [].concat(children)) {
        if (c === null || c === undefined || c === false) continue;
        el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
      }
    }
    return el;
  };

  const getStack = () => {
    let stack = document.getElementById(STACK_ID);
    if (!stack) {
      stack = h('div', { id: STACK_ID, class: 'tce-root tce-stack' });
      (document.body || document.documentElement).appendChild(stack);
    }
    applyTheme(stack);
    return stack;
  };

  const FOCUSABLE = 'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), summary, [tabindex]:not([tabindex="-1"])';

  /**
   * Button factory. variant: primary | secondary | subtle | danger
   */
  const button = ({ label, variant = 'secondary', icon, onClick, title, disabled, id, ariaLabel, className } = {}) => {
    const btn = h('button', {
      type: 'button',
      id,
      class: `tce-btn tce-btn--${variant}${className ? ' ' + className : ''}`,
      title,
      'aria-label': ariaLabel,
      disabled: !!disabled
    });
    if (icon && ICONS[icon]) btn.insertAdjacentHTML('beforeend', ICONS[icon]);
    if (label) btn.appendChild(h('span', { class: 'tce-btn__label', text: label }));
    if (onClick) btn.addEventListener('click', onClick);
    return btn;
  };

  const setButtonLabel = (btn, label) => {
    const span = btn.querySelector('.tce-btn__label');
    if (span) span.textContent = label; else btn.textContent = label;
  };

  /**
   * Determinate (value/max) or indeterminate (value === null) progress bar.
   */
  const progressBar = ({ label = 'Progress', value = 0, max = 100 } = {}) => {
    const fill = h('div', { class: 'tce-progress__fill' });
    const el = h('div', {
      class: 'tce-progress',
      role: 'progressbar',
      'aria-label': label,
      'aria-valuemin': '0'
    }, [fill]);
    const set = (v, m = max) => {
      max = m || 100;
      if (v === null || v === undefined) {
        el.classList.add('tce-progress--indeterminate');
        el.removeAttribute('aria-valuenow');
        el.removeAttribute('aria-valuetext');
        fill.style.width = '';
        return;
      }
      el.classList.remove('tce-progress--indeterminate');
      const clamped = Math.max(0, Math.min(v, max));
      el.setAttribute('aria-valuemax', String(max));
      el.setAttribute('aria-valuenow', String(Math.round(clamped)));
      if (max !== 100) el.setAttribute('aria-valuetext', `${Math.round(clamped)} of ${max}`);
      fill.style.width = `${(clamped / max) * 100}%`;
    };
    set(value, max);
    return { el, set };
  };

  // ---------- clipboard ----------

  const copyText = async (text) => {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      // Fallback for frames without clipboard permission / focus.
      try {
        const ta = h('textarea', { class: 'tce-offscreen', 'aria-hidden': 'true' });
        ta.value = text;
        (document.body || document.documentElement).appendChild(ta);
        ta.select();
        const ok = document.execCommand('copy');
        ta.remove();
        return ok;
      } catch (e2) {
        return false;
      }
    }
  };

  /**
   * Disclosure with a read-only code block and a copy button
   * (used for the manual ffmpeg merge command).
   */
  const codeDisclosure = (summary, code) => {
    const copyBtn = button({ label: 'Copy', icon: 'copy', variant: 'subtle', className: 'tce-btn--small' });
    copyBtn.addEventListener('click', async () => {
      const ok = await copyText(code);
      setButtonLabel(copyBtn, ok ? 'Copied' : 'Copy failed');
      setTimeout(() => setButtonLabel(copyBtn, 'Copy'), 1500);
    });
    return h('details', { class: 'tce-disclosure' }, [
      h('summary', { class: 'tce-disclosure__summary', text: summary }),
      h('div', { class: 'tce-disclosure__body' }, [
        h('code', { class: 'tce-code', text: code }),
        copyBtn
      ])
    ]);
  };

  // ---------- panels ----------

  /**
   * Create a non-modal dialog panel in the shared top-right stack.
   * opts: { id, title, onClose, wide, focus (default true), initialFocus }
   * Returns { el, body, footer, titleEl, close, focus }.
   * Escape (while focus is inside), and the Close button, both call close().
   * Focus moves into the panel on open and returns to the previously focused
   * element on close.
   */
  const createPanel = (opts = {}) => {
    const { id, title = '', onClose, wide = false, focus = true } = opts;
    const stack = getStack();
    const titleId = uid('title');
    const titleEl = h('div', { class: 'tce-panel__title', id: titleId, text: title });
    const closeBtn = h('button', {
      type: 'button',
      class: 'tce-icon-btn tce-panel__close',
      'aria-label': 'Close',
      title: 'Close',
      html: ICONS.close
    });
    const body = h('div', { class: 'tce-panel__body' });
    const footer = h('div', { class: 'tce-panel__footer' });
    const el = h('section', {
      id,
      class: `tce-panel${wide ? ' tce-panel--wide' : ''}`,
      role: 'dialog',
      'aria-modal': 'false',
      'aria-labelledby': titleId,
      tabindex: '-1'
    }, [
      h('div', { class: 'tce-panel__header' }, [titleEl, closeBtn]),
      body,
      footer
    ]);

    const previousFocus = document.activeElement;
    let closed = false;

    const close = (reason) => {
      if (closed) return;
      if (onClose && onClose(reason) === false) return; // allow veto
      closed = true;
      const hadFocus = el.contains(document.activeElement);
      el.remove();
      if (hadFocus && previousFocus && previousFocus.isConnected && typeof previousFocus.focus === 'function') {
        try { previousFocus.focus({ preventScroll: true }); } catch (e) {}
      }
    };

    closeBtn.addEventListener('click', () => close('button'));
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !e.defaultPrevented) {
        // Let an open <details> or menu inside handle its own Escape first.
        e.stopPropagation();
        close('escape');
      }
    });

    stack.appendChild(el);

    const focusPanel = () => {
      const target = (opts.initialFocus && el.contains(opts.initialFocus) && !opts.initialFocus.disabled)
        ? opts.initialFocus
        : (body.querySelector(FOCUSABLE) || closeBtn);
      try { target.focus({ preventScroll: true }); } catch (e) {}
    };
    if (focus) requestAnimationFrame(focusPanel);

    return { el, body, footer, titleEl, closeBtn, close, focus: focusPanel, get closed() { return closed; } };
  };

  // ---------- toasts ----------

  const TOAST_ICON = { success: 'check', error: 'error', info: 'info', warning: 'error' };

  /**
   * Show a toast in the stack.
   * opts: { message, title, kind: info|success|warning|error|progress,
   *         timeout (ms; default 5000 for info/success, 0 = persistent),
   *         progress: {value, max} for a determinate bar (omit for spinner only) }
   * Errors use role=alert and stay until closed; others use role=status.
   * Returns { el, update(opts), close() }. update({kind:'error'|'success'})
   * swaps the toast for a fresh one so screen readers announce the outcome.
   */
  const toast = (opts = {}) => {
    const stack = getStack();
    let current = null;

    const build = (o) => {
      const kind = o.kind || 'info';
      const isError = kind === 'error';
      const persistent = isError || kind === 'progress' || o.timeout === 0;
      const msgEl = h('div', { class: 'tce-toast__message', text: o.message || '' });
      const content = h('div', { class: 'tce-toast__content' }, [
        o.title ? h('div', { class: 'tce-toast__title', text: o.title }) : null,
        msgEl
      ]);
      let bar = null;
      // Determinate toasts get a bar; indeterminate ones just show the spinner.
      if (kind === 'progress' && o.progress) {
        const p = o.progress;
        bar = progressBar({ label: o.title || o.message || 'Progress', value: p ? p.value : null, max: p ? p.max : 100 });
        content.appendChild(bar.el);
      }
      const iconHTML = kind === 'progress'
        ? '<span class="tce-spinner" aria-hidden="true"></span>'
        : (ICONS[TOAST_ICON[kind]] || ICONS.info);
      const el = h('div', {
        class: `tce-toast tce-toast--${kind}`,
        role: isError ? 'alert' : 'status',
        'aria-live': isError ? 'assertive' : 'polite',
        'aria-atomic': 'true'
      }, [h('span', { class: 'tce-toast__icon', html: iconHTML }), content]);
      if ((persistent && kind !== 'progress') || o.closable) {
        const closeBtn = h('button', {
          type: 'button',
          class: 'tce-icon-btn tce-toast__close',
          'aria-label': 'Close',
          title: 'Close',
          html: ICONS.close
        });
        closeBtn.addEventListener('click', () => api.close());
        el.appendChild(closeBtn);
      }
      el.addEventListener('keydown', (e) => { if (e.key === 'Escape') api.close(); });
      let timer = null;
      const timeout = o.timeout !== undefined ? o.timeout : (persistent ? 0 : 5000);
      if (timeout > 0) {
        timer = setTimeout(() => api.close(), timeout);
        // Pause auto-dismiss while hovered/focused so it can be read.
        el.addEventListener('mouseenter', () => clearTimeout(timer));
        el.addEventListener('focusin', () => clearTimeout(timer));
        el.addEventListener('mouseleave', () => { timer = setTimeout(() => api.close(), 2500); });
      }
      return { el, msgEl, bar, kind, timer: () => timer };
    };

    const api = {
      get el() { return current && current.el; },
      update(o = {}) {
        if (!current || !current.el.isConnected) return api;
        const newKind = o.kind || current.kind;
        if (newKind !== current.kind) {
          const next = build({ ...opts, ...o, kind: newKind });
          clearTimeout(current.timer());
          current.el.replaceWith(next.el);
          current = next;
          opts = { ...opts, ...o };
          return api;
        }
        if (o.message !== undefined) current.msgEl.textContent = o.message;
        if (current.bar && 'progress' in o) {
          const p = o.progress;
          current.bar.set(p ? p.value : null, p ? p.max : 100);
        }
        return api;
      },
      close() {
        if (!current) return;
        clearTimeout(current.timer());
        current.el.remove();
        current = null;
      }
    };

    current = build(opts);
    stack.appendChild(current.el);
    return api;
  };

  // ---------- menu ----------

  let openMenu = null;

  /**
   * Open a dropdown menu for `trigger`. The menu is appended to <body>
   * (never inside the trigger), positioned with position:fixed below it.
   * items: [{ label, icon, onSelect }]
   * Keyboard: Up/Down/Home/End move, Enter/Space activate, Escape/Tab close,
   * focus returns to the trigger.
   */
  const closeMenu = (restoreFocus) => {
    if (!openMenu) return;
    const { menu, trigger, cleanup } = openMenu;
    openMenu = null;
    cleanup();
    menu.remove();
    trigger.setAttribute('aria-expanded', 'false');
    if (restoreFocus) { try { trigger.focus({ preventScroll: true }); } catch (e) {} }
  };

  const showMenu = (trigger, items, { focusFirst = true, label } = {}) => {
    if (openMenu && openMenu.trigger === trigger) { closeMenu(true); return null; }
    closeMenu(false);

    const menuId = trigger.getAttribute('aria-controls') || uid('menu');
    const menu = h('div', {
      id: menuId,
      class: 'tce-menu',
      role: 'menu',
      'aria-label': label || trigger.getAttribute('aria-label') || undefined
    });
    applyTheme(menu);
    const buttons = items.map((item) => {
      const b = h('button', { type: 'button', class: 'tce-menu__item', role: 'menuitem', tabindex: '-1' });
      if (item.icon && ICONS[item.icon]) b.insertAdjacentHTML('beforeend', ICONS[item.icon]);
      b.appendChild(h('span', { text: item.label }));
      b.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        closeMenu(true);
        item.onSelect && item.onSelect();
      });
      menu.appendChild(b);
      return b;
    });

    (document.body || document.documentElement).appendChild(menu);
    trigger.setAttribute('aria-haspopup', 'menu');
    trigger.setAttribute('aria-controls', menuId);
    trigger.setAttribute('aria-expanded', 'true');

    const position = () => {
      const r = trigger.getBoundingClientRect();
      const mw = menu.offsetWidth;
      const mh = menu.offsetHeight;
      let left = r.left;
      if (left + mw > window.innerWidth - 8) left = Math.max(8, r.right - mw);
      let top = r.bottom + 4;
      if (top + mh > window.innerHeight - 8 && r.top - mh - 4 > 8) top = r.top - mh - 4;
      menu.style.left = `${Math.round(left)}px`;
      menu.style.top = `${Math.round(top)}px`;
    };
    position();

    const focusAt = (i) => {
      const n = buttons.length;
      const idx = ((i % n) + n) % n;
      buttons[idx].focus({ preventScroll: true });
    };

    const onKey = (e) => {
      const idx = buttons.indexOf(document.activeElement);
      switch (e.key) {
        case 'ArrowDown': e.preventDefault(); focusAt(idx + 1); break;
        case 'ArrowUp': e.preventDefault(); focusAt(idx < 0 ? -1 : idx - 1); break;
        case 'Home': e.preventDefault(); focusAt(0); break;
        case 'End': e.preventDefault(); focusAt(-1); break;
        case 'Escape': e.preventDefault(); e.stopPropagation(); closeMenu(true); break;
        case 'Tab': closeMenu(true); break; // focus returns to trigger, then Tab moves on
        default: break;
      }
    };
    const onOutside = (e) => {
      if (!menu.contains(e.target) && !trigger.contains(e.target)) closeMenu(false);
    };
    const onViewportChange = () => closeMenu(false);

    menu.addEventListener('keydown', onKey);
    document.addEventListener('pointerdown', onOutside, true);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('blur', onViewportChange);

    openMenu = {
      menu,
      trigger,
      cleanup: () => {
        menu.removeEventListener('keydown', onKey);
        document.removeEventListener('pointerdown', onOutside, true);
        window.removeEventListener('resize', onViewportChange);
        window.removeEventListener('blur', onViewportChange);
      }
    };

    if (focusFirst) focusAt(0);
    return { menu, close: () => closeMenu(true) };
  };

  /**
   * Wire `trigger` as a menu button: click toggles, ArrowDown/Enter/Space opens
   * with focus on the first item. getItems() is called on each open.
   */
  const bindMenuButton = (trigger, getItems, opts = {}) => {
    trigger.setAttribute('aria-haspopup', 'menu');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      // e.detail === 0 means keyboard activation (Enter/Space).
      showMenu(trigger, getItems(), { ...opts, focusFirst: true });
    });
    trigger.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        e.stopPropagation();
        if (!openMenu || openMenu.trigger !== trigger) showMenu(trigger, getItems(), opts);
        if (e.key === 'ArrowUp' && openMenu) {
          const items = openMenu.menu.querySelectorAll('[role="menuitem"]');
          items[items.length - 1]?.focus();
        }
      }
    });
  };

  // ---------- filenames ----------

  /**
   * Keep every character except those illegal in filenames on common OSes
   * (\ / : * ? " < > | and control chars). Collapses whitespace, trims
   * leading/trailing dots and spaces, caps length; falls back to `fallback`.
   */
  const sanitizeFilename = (name, fallback = 'transcript', maxLength = 120) => {
    let s = String(name || '')
      // eslint-disable-next-line no-control-regex
      .replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^[.\s]+|[.\s]+$/g, '');
    if (s.length > maxLength) s = Array.from(s).slice(0, maxLength).join('').trim();
    if (/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(s)) s = `${s}_`;
    return s || fallback;
  };

  globalThis.__tceUI = {
    ICONS,
    detectTheme,
    applyTheme,
    h,
    uid,
    getStack,
    button,
    setButtonLabel,
    progressBar,
    copyText,
    codeDisclosure,
    createPanel,
    toast,
    showMenu,
    closeMenu,
    bindMenuButton,
    sanitizeFilename
  };
})();
