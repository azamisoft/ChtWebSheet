import test from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { initializeLocale } from '../../src/modules/ui/locale.js';
import { localeMessages } from '../../src/modules/ui/locales/messages.js';
import { readFileSync } from 'node:fs';
const settle = () => new Promise(resolve => setTimeout(resolve, 0));
function fixture(saved = 'ja') {
  const dom = new JSDOM(`<html><body><div id="emptyState"></div><div class="ribbon-display-options"></div><button id="save" title="保存">保存</button><span data-cws-i18n="リレーションシップ">リレーション<br>シップ</span><div id="sheetHost"><span>保存</span><svg><text>ホーム</text></svg></div><div id="sheetTabs">設定</div><div class="cws-ai-message-bubble">保存</div><div class="cws-ai-session-main">設定</div><div class="cws-ai-account-label">更新</div><textarea placeholder="直接入力">保存</textarea></body></html>`, { url: 'https://example.test' });
  const storage = dom.window.localStorage;
  storage.setItem('cws.ui.locale', saved);
  return { dom, document: dom.window.document, storage, locale: initializeLocale(dom.window.document, storage) };
}
test('switches and restores UI without changing workbook, conversation, names or input', async () => {
  const { dom, document: d, locale, storage } = fixture();
  for (const lang of ['en','zh-CN','ja']) {
    locale.setLocale(lang); await settle();
    assert.equal(d.querySelector('#save').textContent, {en:'Save','zh-CN':'保存',ja:'保存'}[lang]);
    assert.equal(d.querySelector('textarea').placeholder, {en:'Enter directly','zh-CN':'直接输入',ja:'直接入力'}[lang]);
    assert.equal(d.querySelector('textarea').value,'保存');
    assert.equal(d.querySelector('#sheetHost').textContent,'保存ホーム');
    const dynamic=d.createElement('span'); dynamic.textContent='設定'; d.querySelector('#sheetHost').append(dynamic); await settle(); assert.equal(dynamic.textContent,'設定'); dynamic.remove();
    assert.equal(d.querySelector('.cws-ai-message-bubble').textContent,'保存');
    assert.equal(d.querySelector('.cws-ai-session-main').textContent,'設定');
    assert.equal(d.querySelector('.cws-ai-account-label').textContent,'更新');
    assert.equal(storage.getItem('cws.ui.locale'),lang);
  }
  assert.equal(d.querySelector('[data-cws-i18n]').innerHTML,'リレーション<br>シップ');
  locale.destroy(); dom.window.close();
});
test('persisted locale and app-rendered controls update without an observer loop', async () => {
  const {dom,document:d,locale} = fixture('en');
  assert.equal(d.documentElement.lang,'en');
  const button=d.createElement('button'); button.textContent='設定'; d.body.append(button);
  await settle(); assert.equal(button.textContent,'Settings');
  button.textContent='閉じる'; await settle(); assert.equal(button.textContent,'Close');
  locale.setLocale('ja'); await settle(); assert.equal(button.textContent,'閉じる');
  locale.destroy(); dom.window.close();
});
test('language menu supports selection, Escape and focus restoration', async () => {
  const {dom,document:d,locale} = fixture();
  const trigger=d.querySelector('#cwsLanguageButton'); trigger.click();
  assert.equal(trigger.getAttribute('aria-expanded'),'true');
  d.querySelector('#cwsLanguageMenu [data-locale="zh-CN"]').click(); await settle();
  assert.equal(d.documentElement.lang,'zh-CN'); assert.equal(d.activeElement,trigger);
  trigger.click(); trigger.dispatchEvent(new dom.window.KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
  assert.equal(trigger.getAttribute('aria-expanded'),'false');
  locale.destroy(); dom.window.close();
});
test('static interface labels and composite labels have both translations', () => {
  const dom=new JSDOM(readFileSync(new URL('../../index.html',import.meta.url),'utf8'));
  const d=dom.window.document, messages=localeMessages(); const missing=[];
  const walker=d.createTreeWalker(d.body,4);
  for(let node; (node=walker.nextNode());) {
    const key=node.textContent.trim();
    if (!/[ぁ-んァ-ン一-龯]/.test(key) || node.parentElement.closest('script,style,[data-cws-i18n]') || ['MS Pゴシック','MS P明朝','ＭＳ Ｐゴシック'].includes(key)) continue;
    if(!messages[key]?.en || !messages[key]?.['zh-CN']) missing.push(key);
  }
  for(const el of d.querySelectorAll('[data-cws-i18n]')) assert.ok(messages[el.dataset.cwsI18n]?.en,el.dataset.cwsI18n);
  assert.deepEqual(missing,[]); dom.window.close();
});
