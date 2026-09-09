import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { en, zhCN, translate, resolveLocale, languagePreference, errorText } from '../src/taskpane/i18n';
import type { MessageKey } from '../src/taskpane/i18n';
test('manual language overrides Excel; auto prefers Excel then browser and has a fallback', () => {
  assert.equal(resolveLocale('auto', 'zh-CN', 'en-US'), 'zh-CN');
  assert.equal(resolveLocale('auto', 'en-GB', 'zh-CN'), 'en');
  assert.equal(resolveLocale('auto', undefined, 'zh-SG'), 'zh-CN');
  assert.equal(resolveLocale('auto', 'fr-FR'), 'en');
  assert.equal(resolveLocale('en', 'zh-CN'), 'en');
  assert.equal(resolveLocale('zh-CN', 'en-US'), 'zh-CN');
  assert.equal(languagePreference('invalid'), 'auto');
});
test('translations preserve all interpolation parameters and DOM keys exist', async () => {
  const placeholders = (value: string) => [...value.matchAll(/\{\w+\}/g)].map(match => match[0]).sort();
  for (const key of Object.keys(en) as MessageKey[]) {
    assert.ok(zhCN[key].length > 0, key);
    assert.deepEqual(placeholders(en[key]), placeholders(zhCN[key]), key);
  }
  const html = await readFile('src/taskpane/taskpane.html', 'utf8');
  for (const match of html.matchAll(/data-i18n(?:-label)?="([^"]+)"/g)) assert.ok(match[1]! in en, match[1]);
  assert.equal(translate('zh-CN', 'scanProgress', { sheet: '<img src=x>', count: 49 }),
    '正在扫描 <img src=x> · 已检查 49 个单元格…'); // Returned as text, never as markup.
});
test('localized errors preserve important cleanup uncertainty and English original details', () => {
  assert.match(errorText('zh-CN', 'REFRESH_FAILED'), /部分新预览可能仍需清理/);
  assert.match(errorText('zh-CN', 'CLEANUP_FAILED'), /部分预览可能已经发生变化/);
  assert.equal(errorText('en', 'OFFICE_FILE', 'Host file code 42'), 'Host file code 42');
  assert.match(errorText('zh-CN', 'UNKNOWN'), /原始公式未修改/);
});
