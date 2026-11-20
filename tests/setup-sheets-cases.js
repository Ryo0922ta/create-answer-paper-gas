const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({});
for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8'), context);
}
let basicHeading = '順番', headingWrites = 0;
let templateRows = [], templateReads = 0;
const sheets = { '基本パーツ': {getRange(address, column, rows, columns) {
  if (typeof address === 'number') {
    assert.deepEqual([address, column, rows, columns], [4, 1, 10, 7]);
    templateReads++;
    return {getValues: () => templateRows};
  }
  assert.equal(address, 'B3');
  return {
    getDisplayValue: () => basicHeading,
    setValue(value) { basicHeading = value; headingWrites++; return this; },
    setNote(note) { assert.match(note, /数値は自動配置/); },
  };
}} };
const alerts = [];
let draft = { parts: [] }, committed, plan, outputConfigured = false, outputFlushed = false;
context.SpreadsheetApp = { getActiveSpreadsheet: () => ({
  getSheetByName: name => sheets[name],
  insertSheet: name => (sheets[name] = { name }),
  deleteSheet: sheet => { delete sheets[sheet.name]; },
}), getUi: () => ({ alert: message => alerts.push(message) }), flush: () => {
  assert.equal(outputConfigured, true);
  outputFlushed = true;
} };
context.withDraftLock_ = callback => callback();
context.readDraft_ = () => draft;
context.configureOutputSheet_ = () => { outputConfigured = true; outputFlushed = false; };
context.logEvent_ = () => {};
context.Utilities = { getUuid: (() => { let n = 0; return () => String(++n); })() };
context.commitDraft_ = (next, signature, previous) => {
  assert.equal(outputFlushed, true, '配置前に出力シートの初期設定を反映する');
  assert.equal(signature, null);
  assert.equal(previous, draft);
  context.validateDraft_(next);
  plan = context.outputPlan_(previous, next, { rows: [], merges: [] });
  committed = next;
  draft = next;
};
context.setupSheetsLogged_();
assert.equal(basicHeading, '配置の左上セル');
assert.ok(sheets['出力']);
assert.equal(committed.parts.length, 3);
assert.deepEqual(Array.from(committed.parts, part => part.type), ['考査見出し', '得点枠', '年組氏名欄']);
for (const [type, row] of [['得点枠', 44], ['年組氏名欄', 48]]) {
  const rect = plan.additions.find(item => item.part.type === type).rect;
  assert.equal(rect.row, row);
  assert.equal(rect.column, 52);
}
const original = sheets['出力'];
committed = null;
context.setupSheetsLogged_();
assert.equal(sheets['出力'], original);
assert.equal(committed, null);
assert.equal(headingWrites, 1);
assert.equal(templateReads, 1, '既存の出力は再配置しない');
delete sheets['出力'];
assert.throws(() => context.setupSheetsLogged_(), /以前の作業データ/);
assert.equal(sheets['出力'], undefined);

draft = {parts: []};
templateRows = [
  [true, ' d8 ', '考査見出し', '期末考査', 20, 3, 0],
  [true, 9, '年組氏名欄', '', 16, 2, 2],
  [true, 'AM12', '得点枠', '', 24, 3, 1],
];
context.setupSheetsLogged_();
assert.deepEqual(Array.from(plan.additions, entry => context.outputA1_(entry.rect)),
  ['D8:W10', 'AM12:BJ14', 'AZ48:BO49']);
assert.deepEqual(Array.from(committed.parts, part => [part.sourceRow, part.text, part.width, part.height, part.gap]),
  [[4, '期末考査', 20, 3, 0], [6, '', 24, 3, 1], [5, '', 16, 2, 2]]);

delete sheets['出力'];
draft = {parts: []};
templateRows[2][1] = 'D8';
assert.throws(() => context.setupSheetsLogged_(), /考査見出し.*D8:W10.*重なっています/);
assert.equal(sheets['出力'], undefined);
assert.equal(draft.parts.length, 0);

templateRows[2][1] = 'BT51';
assert.throws(() => context.setupSheetsLogged_(), /得点枠.*BT51:CQ53.*用紙外/);
assert.equal(sheets['出力'], undefined);
assert.equal(draft.parts.length, 0);

templateRows[2][1] = 'A52';
assert.throws(() => context.setupSheetsLogged_(), /得点枠.*A52.*範囲内/);
assert.equal(sheets['出力'], undefined);
assert.equal(draft.parts.length, 0);

draft = {parts: []};
context.commitDraft_ = () => { throw new Error('配置失敗'); };
assert.throws(() => context.setupSheetsLogged_(), /配置失敗/);
assert.equal(sheets['出力'], undefined);
console.log('setup sheets cases passed');
