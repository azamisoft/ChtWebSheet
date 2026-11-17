import { localeMessages } from './locales/messages.js';

export function normalizeLocale(value) {
  return value === 'en' || value === 'zh-CN' ? value : 'ja';
}

// Only interface copy belongs here. Workbook and conversation content are never translated.
export function localeContentSelector() {
  return [
    '[data-cws-no-translate]', '[translate="no"]', 'script', 'style', 'textarea',
    '[contenteditable="true"]', '#sheetContainer', '#sheetHost', '#sheetTabs', '#workbookWindowBar',
    '#workbookName', '#workbookNameDisplay', '#selectedCellLabel',
    '.cws-ai-message-bubble', '.cws-ai-account-label', '.cws-ai-session-title',
    '.cws-ai-session-name', '.cws-ai-session-main', '.cws-ai-attachment-name', '.cws-ai-clarification-title',
    '.cws-ai-clarification-option', '.cws-ai-clarification-status',
    '.cws-ai-op-list', '.cws-ai-op-summary', '.cws-ai-task-evidence',
    '[data-cws-ai-system-prompt]'
  ].join(',');
}

export function translateInterfaceText(source, locale, messages) {
  const key = source.trim();
  const result = locale === 'ja' ? key : messages[key]?.[locale];
  if (!result) {
    if (locale === 'ja') return source;
    const rules = [
      [/^(.+) を復元しました。ページを再読み込みしても内容は保持されます。$/, 'Restored $1. Your content is kept after reloading.', '已恢复 $1，重新加载页面后仍会保留内容。'],
      [/^(\d+) 個のワークシートを読み込みました。セルは直接編集でき、数式は自動再計算されます。$/, 'Loaded $1 worksheets. Edit cells directly; formulas recalculate automatically.', '已加载 $1 张工作表。可直接编辑单元格，公式会自动重新计算。'],
      [/^(.+) を開いています\.\.\.$/, 'Opening $1…', '正在打开 $1…'],
      [/^(.+) を保存しました。$/, 'Saved $1.', '已保存 $1。'],
      [/^データの個数[：:]\s*(.+)$/, 'Count: $1', '计数：$1'],
      [/^数値の個数[：:]\s*(.+)$/, 'Numerical count: $1', '数值计数：$1'],
      [/^平均[：:]\s*(.+)$/, 'Average: $1', '平均值：$1'],
      [/^合計[：:]\s*(.+)$/, 'Sum: $1', '求和：$1'],
      [/^最小値[：:]\s*(.+)$/, 'Minimum: $1', '最小值：$1'],
      [/^最大値[：:]\s*(.+)$/, 'Maximum: $1', '最大值：$1']
    ];
    for (const [pattern, en, zh] of rules) if (pattern.test(key)) return source.replace(key, key.replace(pattern, locale === 'en' ? en : zh));
    return source;
  }
  return source.replace(key, result);
}

export function initializeLocale(document, storage) {
  const messages = localeMessages();
  let locale = 'ja';
  try { storage ||= document.defaultView.localStorage; locale = normalizeLocale(storage.getItem('cws.ui.locale')); } catch { /* Private mode. */ }
  const excluded = localeContentSelector();
  const texts = new WeakMap();
  const attrs = new WeakMap();
  const composites = new WeakMap();
  const attributes = ['title', 'aria-label', 'placeholder', 'alt'];
  const ui = document.createElement('div');
  ui.className = 'cws-language-control';
  ui.setAttribute('data-cws-no-translate', '');
  ui.innerHTML = `<button type="button" id="cwsLanguageButton" class="cws-language-button" aria-haspopup="menu" aria-expanded="false" aria-controls="cwsLanguageMenu"><span aria-hidden="true">◎</span><span data-language-label></span><span aria-hidden="true">▾</span></button><div id="cwsLanguageMenu" class="cws-language-menu" role="menu" hidden><button type="button" role="menuitemradio" data-locale="ja">日本語</button><button type="button" role="menuitemradio" data-locale="en">English</button><button type="button" role="menuitemradio" data-locale="zh-CN">简体中文</button></div>`;
  const anchor = document.querySelector('.ribbon-display-options');
  anchor?.before(ui);
  // The same control remains available on the welcome screen.
  const welcomeControl = ui.cloneNode(true);
  welcomeControl.classList.add('is-welcome');
  welcomeControl.querySelector('button').id = 'cwsWelcomeLanguageButton';
  welcomeControl.querySelector('button').setAttribute('aria-controls', 'cwsWelcomeLanguageMenu');
  welcomeControl.querySelector('[role="menu"]').id = 'cwsWelcomeLanguageMenu';
  document.querySelector('#emptyState')?.prepend(welcomeControl);

  function blocked(element) { return Boolean(element?.closest(excluded)); }
  function text(node) {
    if (blocked(node.parentElement) || node.parentElement.closest('[data-cws-i18n]')) return;
    let entry = texts.get(node);
    if (!entry || node.data !== entry.rendered) entry = { source: node.data, rendered: node.data };
    const rendered = translateInterfaceText(entry.source, locale, messages);
    entry.rendered = rendered;
    texts.set(node, entry);
    if (node.data !== rendered) node.data = rendered;
  }
  function element(el) {
    if (blocked(el) && !(el.tagName === 'TEXTAREA' && !el.matches('[data-cws-ai-system-prompt]'))) return;
    if (el.hasAttribute('data-cws-i18n')) {
      if (!composites.has(el)) composites.set(el, el.innerHTML);
      const translated = messages[el.getAttribute('data-cws-i18n')]?.[locale];
      const desired = locale === 'ja' || !translated ? composites.get(el) : translated;
      if (el.innerHTML !== desired) el.innerHTML = desired;
    }
    let records = attrs.get(el);
    if (!records) { records = {}; attrs.set(el, records); }
    for (const name of attributes) {
      if (!el.hasAttribute(name)) { delete records[name]; continue; }
      const current = el.getAttribute(name);
      let entry = records[name];
      if (!entry || current !== entry.rendered) entry = { source: current, rendered: current };
      entry.rendered = translateInterfaceText(entry.source, locale, messages);
      records[name] = entry;
      if (current !== entry.rendered) el.setAttribute(name, entry.rendered);
    }
  }
  function visit(root) {
    if (root.nodeType === 3) { text(root); return; }
    if (root.nodeType !== 1) return;
    element(root);
    if (blocked(root)) return;
    // Walk only UI subtrees; never traverse grid cells, SVG drawings or chat text.
    for (const child of root.childNodes) visit(child);
  }
  function close(control, focus = false) {
    control.querySelector('[role="menu"]').hidden = true;
    const button = control.querySelector('button');
    button.setAttribute('aria-expanded', 'false');
    if (focus) button.focus();
  }
  function updateControls() {
    for (const control of [ui, welcomeControl]) {
      control.querySelector('[data-language-label]').textContent = { ja: '日本語', en: 'English', 'zh-CN': '中文' }[locale];
      control.querySelector('button').setAttribute('aria-label', {ja:'言語を選択',en:'Choose language','zh-CN':'选择语言'}[locale]);
      control.querySelectorAll('[data-locale]').forEach(button => button.setAttribute('aria-checked', String(button.dataset.locale === locale)));
    }
  }
  function setLocale(value) {
    locale = normalizeLocale(value);
    document.documentElement.lang = locale;
    try { storage.setItem('cws.ui.locale', locale); } catch { /* Still switch this page. */ }
    visit(document.body);
    updateControls();
    document.dispatchEvent(new document.defaultView.CustomEvent('cws:localechange', { detail: { locale } }));
  }
  for (const control of [ui, welcomeControl]) {
    const button = control.querySelector('button');
    const menu = control.querySelector('[role="menu"]');
    button.addEventListener('click', () => {
      const opening = menu.hidden;
      [ui, welcomeControl].forEach(other => close(other));
      menu.hidden = !opening;
      button.setAttribute('aria-expanded', String(opening));
      if (opening) menu.querySelector('[aria-checked="true"]').focus();
    });
    control.addEventListener('click', event => {
      const option = event.target.closest('[data-locale]');
      if (!option) return;
      setLocale(option.dataset.locale);
      close(control, true);
    });
    control.addEventListener('keydown', event => {
      if (event.key === 'Escape') { close(control, true); event.preventDefault(); }
      if (menu.hidden || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
      const options = [...menu.querySelectorAll('[data-locale]')];
      const current = options.indexOf(document.activeElement);
      const next = event.key === 'Home' ? 0 : event.key === 'End' ? options.length - 1 : (current + (event.key === 'ArrowDown' ? 1 : -1) + options.length) % options.length;
      options[next].focus();
      event.preventDefault();
    });
  }
  document.addEventListener('pointerdown', event => {
    for (const control of [ui, welcomeControl]) if (!control.contains(event.target)) close(control);
  });
  const observer = new document.defaultView.MutationObserver(records => {
    for (const record of records) {
      if (blocked(record.target.nodeType === 1 ? record.target : record.target.parentElement) && !(record.type === 'attributes' && record.target.tagName === 'TEXTAREA')) continue;
      if (record.type === 'characterData') text(record.target);
      else if (record.type === 'attributes') element(record.target);
      else for (const added of record.addedNodes) visit(added);
    }
  });
  setLocale(locale);
  observer.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: attributes });
  return { setLocale, getLocale: () => locale, destroy: () => observer.disconnect() };
}
