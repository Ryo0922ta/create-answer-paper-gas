const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({});
for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8'), context);
}
const sheets = { '基本パーツ': {} };
const alerts = [];
let draft = { parts: [] }, committed;
context.SpreadsheetApp = { getActiveSpreadsheet: () => ({
  getSheetByName: name => sheets[name],
  insertSheet: name => (sheets[name] = { name }),
  deleteSheet: sheet => { delete sheets[sheet.name]; },
}), getUi: () => ({ alert: message => alerts.push(message) }) };
context.withDraftLock_ = callback => callback();
context.readDraft_ = () => draft;
context.configureOutputSheet_ = () => {};
context.logEvent_ = () => {};
context.Utilities = { getUuid: (() => { let n = 0; return () => String(++n); })() };
context.commitDraft_ = next => {
  committed = next;
  context.validateDraft_(next);
  const plan = context.outputPlan_(draft, next, { rows: [], merges: [] });
  draft = next;
  return plan;
};
context.setupSheetsLogged_();
assert.ok(sheets['出力']);
assert.equal(committed.parts.length, 3);
assert.deepEqual(Array.from(committed.parts, part => part.type), ['考査見出し', '得点枠', '年組氏名欄']);
const plan = context.outputPlan_({parts: []}, committed, {rows: [], merges: []});
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
delete sheets['出力'];
assert.throws(() => context.setupSheetsLogged_(), /以前の作業データ/);
assert.equal(sheets['出力'], undefined);
draft = {parts: []};
context.commitDraft_ = () => { throw new Error('配置失敗'); };
assert.throws(() => context.setupSheetsLogged_(), /配置失敗/);
assert.equal(sheets['出力'], undefined);
console.log('setup sheets cases passed');
