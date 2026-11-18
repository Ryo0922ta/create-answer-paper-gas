const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const html = fs.readFileSync(require('node:path').join(__dirname, '../src/PartEditor.html'), 'utf8');
const start = html.indexOf('  function submitAnalyzedDocument(payload) {');
const end = html.indexOf('\n  loadParts();\n  loadBasicTemplates();', start);
assert.ok(start >= 0 && end > start);
let clock = 0, scheduled, success, failure, calls = 0, shown = '', chosen = 0;
const elements = {};
const ctx = vm.createContext({
  performance: {now: () => clock}, console: {info() {}},
  setTimeout(fn, delay) { assert.equal(delay, 10000); scheduled = fn; return 1; },
  clearTimeout() { scheduled = null; },
  byId(id) { return elements[id] ||= {disabled:false, innerHTML:'', classList:{add(){},remove(){}}}; },
  message(value) { shown = value; }, failure(e) { shown = e.message; },
  showPlacementChoice() { chosen++; }, loadParts() {}, confirm() { return false; },
  google: {script:{run:{withSuccessHandler(fn) { success = fn; return this; },
    withFailureHandler(fn) { failure = fn; return this; }, importAnalyzedDocument() { calls++; }}}},
});
vm.runInContext('let placementPending = false, deletePending = false, placementRevision = 0, analyzed = [{}];\n' + html.slice(start, end), ctx);
ctx.submitAnalyzedDocument({previewOnly:true});
ctx.submitAnalyzedDocument({previewOnly:true});
assert.equal(calls, 1);
clock = 10000; scheduled();
assert.match(shown, /10秒以上/);
ctx.submitAnalyzedDocument({previewOnly:true});
assert.equal(calls, 1);
success({placementPreview:true});
assert.equal(chosen, 1);

ctx.submitAnalyzedDocument({previewOnly:false});
clock += 10000; scheduled();
assert.match(shown, /処理が続いて/);
failure(new Error('late failure'));
assert.match(shown, /late failure/);

ctx.submitAnalyzedDocument({previewOnly:true});
vm.runInContext('placementRevision++', ctx);
clock += 120; success({placementPreview:true});
assert.equal(chosen, 1);
assert.equal(scheduled, null);
ctx.submitAnalyzedDocument({previewOnly:false});
clock += 10000; scheduled();
success('配置しました。');
assert.match(shown, /配置しました。/);
assert.match(shown, /3500ms/);
console.log('配置候補・配置処理の10秒タイムアウト: OK');
