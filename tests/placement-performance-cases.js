const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {performance} = require('node:perf_hooks');
const ctx = vm.createContext({});
ctx.withOperationLog_ = (_name, fn) => fn();
for (const name of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', name), 'utf8'), ctx);
}
const samples = [];
for (let i = 0; i < 20; i++) {
  const question = {id:'q', kind:'question', number:'1', order:1, gap:1, width:33, height:7,
    rows:Array.from({length:3}, (_, n) => ({rowLabel:String(n+1), height:2, pattern:'小問複合枠',
      frameCount:5, itemsPerFrame:1, labels:['a','b','c','d','e'], weights:[1,1,1,1,1]}))};
  const start = performance.now();
  const plan = ctx.outputPlan_({parts:[]}, {parts:[question]}, {rows:[], merges:[]});
  const requests = ctx.outputDrawRequests_(1, plan);
  samples.push(performance.now() - start);
  assert.ok(requests.length > 0);
}
const max = Math.max(...samples);
assert.ok(max < 3000, 'ローカルの配置計算と要求生成が3000msを超えました');
console.log('配置計算・要求生成（通信なし、20回）最大: ' + max.toFixed(2) + 'ms');
console.log('注意: 実シートの通信と描画は含まれません。');
let reads = 0, alerts = 0;
ctx.SpreadsheetApp = {getActiveSpreadsheet: () => ({getSheetByName: () => ({})}),
  getUi: () => ({alert: text => { alerts++; assert.match(text, /配置や画面への反映は含みません/); }})};
ctx.draftSheet_ = () => ({});
ctx.readDraft_ = () => ({});
ctx.outputReadGrid_ = () => { reads++; return {rows:[], merges:[]}; };
ctx.console = {info() {}};
vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/PerformanceBenchmark.js'), 'utf8'), ctx);
const report = ctx.benchmarkPlacementRead();
assert.equal(reads, 5);
assert.equal(alerts, 1);
assert.equal(report.samplesMs.length, 5);
ctx.draftSheet_ = () => null;
assert.throws(() => ctx.benchmarkPlacementRead(), /作業データ/);
