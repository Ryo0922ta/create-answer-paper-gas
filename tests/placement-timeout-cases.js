const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const html = fs.readFileSync(require('node:path').join(__dirname, '../src/PartEditor.html'), 'utf8');
const start = html.indexOf('  function submitPlacement(payload) {');
const choiceStart = html.indexOf('  function showPlacementChoice(payload, result) {');
const end = html.indexOf('\n  loadParts();\n  loadBasicTemplates();', start);
assert.ok(start >= 0 && end > start);
let clock = 0, scheduled, success, failure, calls = 0, shown = '', chosen = 0, basicArgs, selectedCell, editArgs;
const elements = {};
const ctx = vm.createContext({
  performance: {now: () => clock}, console: {info() {}},
  Event,
  setTimeout(fn, delay) { assert.equal(delay, 10000); scheduled = fn; return 1; },
  clearTimeout() { scheduled = null; },
  byId(id) { return elements[id] ||= {disabled:false, innerHTML:'', classList:{add(){},remove(){}}}; },
  message(value) { shown = value; }, failure(e) { shown = e.message; },
  showPlacementChoice() { chosen++; }, loadParts() {}, confirm() { return false; },
  google: {script:{run:{withSuccessHandler(fn) { success = fn; return this; },
    withFailureHandler(fn) { failure = fn; return this; }, importAnalyzedDocument() { calls++; },
    addBasicPart(...args) { basicArgs = args; }, getSelectedOutputCell() { selectedCell = success; },
    updateEditablePart(data) { editArgs = data; }}}},
});
vm.runInContext('let placementPending = false, deletePending = false, placementRevision = 0, analyzed = [{}];\n' + html.slice(start, end), ctx);
ctx.submitPlacement({previewOnly:true});
ctx.submitPlacement({previewOnly:true});
assert.equal(calls, 1);
clock = 10000; scheduled();
assert.match(shown, /10秒以上/);
ctx.submitPlacement({previewOnly:true});
assert.equal(calls, 1);
success({placementPreview:true});
assert.equal(chosen, 1);

ctx.submitPlacement({previewOnly:false});
clock += 10000; scheduled();
assert.match(shown, /処理が続いて/);
failure(new Error('late failure'));
assert.match(shown, /late failure/);

ctx.submitPlacement({previewOnly:true});
vm.runInContext('placementRevision++', ctx);
clock += 120; success({placementPreview:true});
assert.equal(chosen, 1);
assert.equal(scheduled, null);
ctx.submitPlacement({previewOnly:false});
clock += 10000; scheduled();
success('配置しました。');
assert.match(shown, /配置しました。/);
assert.match(shown, /3500ms/);

function createElement(tag) {
  const classes = new Set();
  return {tag, children:[], dataset:{}, isConnected:true,
    classList:{add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name)},
    appendChild(child) { this.children.push(child); },
    dispatchEvent(event) { if (event.type === 'input') this.oninput?.(); },
    click() { this.onclick(); },
    set innerHTML(_value) { this.children = []; },
    querySelectorAll(name) { return this.children.flatMap(child => child.tag === name ? [child] : child.querySelectorAll(name)); },
  };
}
ctx.document = {createElement};
const panel = elements.basicPlacementChoice = createElement('div');
const pickStart = html.indexOf('  function pickPlacementCell(');
const pickEnd = html.indexOf('  function partAction(', pickStart);
vm.runInContext(html.slice(pickStart, pickEnd) + html.slice(choiceStart, start), ctx);
const basic = {basicType:'得点枠', previewOnly:true};
ctx.submitPlacement(basic);
assert.deepEqual(basicArgs, ['得点枠', true, null, null]);
clock += 10000; scheduled();
success({placementPreview:true, range:'AZ44:BT47', signature:'auto'});
assert.ok(panel.children.some(child => child.textContent === '得点枠：AZ44:BT47'));
panel.children.find(child => child.textContent === '位置を変更').click();
const form = panel.children.at(-1);
const input = form.querySelectorAll('input')[0];
assert.equal(input.value, 'AZ44');
form.children.find(child => child.textContent === 'シートで選択したセルを取得').click();
selectedCell('D12');
assert.equal(input.value, 'D12');
input.value = 'D13';
input.oninput();
form.children.find(child => child.textContent === '指定位置を確認').click();
assert.deepEqual(basicArgs, ['得点枠', true, null, 'D13']);
success({placementPreview:true, range:'D13:AJ16', signature:'manual'});
panel.children.find(child => child.textContent === 'この位置に配置').click();
assert.deepEqual(basicArgs, ['得点枠', false, 'manual', 'D13']);
clock += 10000; scheduled();
success('得点枠を出力に配置しました。');
assert.match(shown, /得点枠を出力に配置しました。/);
const editPanel = elements.editPlacementChoice = createElement('div');
ctx.submitPlacement({editPart:{id:'heading', placementCell:'C12'}, editKind:'basic', editLabel:'考査見出し', previewOnly:true});
assert.equal(editArgs.previewOnly, true);
success({placementPreview:true, sourceRange:'C3:AI4', range:'C12:AI13', signature:'move', revision:'revision'});
assert.ok(editPanel.children.some(child => /C3:AI4 → C12:AI13/.test(child.textContent)));
editPanel.children.find(child => child.textContent === '変更を出力に反映').click();
assert.equal(editArgs.previewOnly, false);
assert.equal(editArgs.revision, 'revision');
assert.equal(editArgs.placementSignature, 'move');
clock += 1400;
success('基本パーツを更新しました。');
assert.match(shown, /1400ms/);
console.log('配置候補・配置処理の10秒タイムアウト: OK');
