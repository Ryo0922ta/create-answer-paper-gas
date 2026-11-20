/** 作業状態の再開、配置、設定保存をローカルのシート模型で確認する。 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function columnNumber(label) {
  return [...label].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
}
class FakeRange {
  constructor(sheet, row, column, height, width) {
    Object.assign(this, { sheet, row, column, height, width });
  }
  getValues() {
    return Array.from({ length: this.height }, (_, r) =>
      Array.from({ length: this.width }, (_, c) =>
        this.sheet.cells.get((this.row + r) + ':' + (this.column + c)) ?? ''));
  }
  getDisplayValues() {
    return this.getValues().map(row => row.map(value => String(value)));
  }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return String(this.getValue()); }
  setValues(values) {
    assert.equal(values.length, this.height);
    values.forEach((row, r) => {
      assert.equal(row.length, this.width);
      row.forEach((value, c) => this.sheet.cells.set((this.row + r) + ':' + (this.column + c), value));
    });
    return this;
  }
  setValue(value) { return this.setValues([[value]]); }
  clearContent() {
    for (let r = 0; r < this.height; r++)
      for (let c = 0; c < this.width; c++)
        this.sheet.cells.delete((this.row + r) + ':' + (this.column + c));
    return this;
  }
  clearDataValidations() { return this; }
  setNumberFormat() { return this; }
}
class FakeSheet {
  constructor(name, book) {
    this.name = name; this.book = book; this.cells = new Map();
    this.rows = 200; this.columns = 30; this.hidden = false;
  }
  getParent() { return this.book; }
  getRange(a, b, c = 1, d = 1) {
    if (typeof a === 'number') return new FakeRange(this, a, b, c, d);
    const [first, last = first] = a.split(':');
    const parse = address => {
      const match = /^([A-Z]+)(\d+)$/.exec(address);
      return [Number(match[2]), columnNumber(match[1])];
    };
    const [row, column] = parse(first), [endRow, endColumn] = parse(last);
    return new FakeRange(this, row, column, endRow - row + 1, endColumn - column + 1);
  }
  getMaxRows() { return this.rows; }
  getMaxColumns() { return this.columns; }
  insertRowsAfter(_, count) { this.rows += count; }
  insertColumnsAfter(_, count) { this.columns += count; }
  hideSheet() { this.hidden = true; }
  showSheet() { this.hidden = false; }
  setName(name) {
    this.book.sheets.delete(this.name);
    this.name = name;
    this.book.sheets.set(name, this);
    return this;
  }
  isSheetHidden() { return this.hidden; }
}
class FakeBook {
  constructor() { this.sheets = new Map(); }
  getSheetByName(name) { return this.sheets.get(name) || null; }
  insertSheet(name) {
    const sheet = new FakeSheet(name, this);
    this.sheets.set(name, sheet);
    return sheet;
  }
}
const book = new FakeBook();
const settings = book.insertSheet('設定');
const basic = book.insertSheet('基本パーツ');
book.insertSheet('出力');
settings.getRange('A20').setValue('使用');
settings.getRange('J5').setValue('対象の大問番号');
basic.getRange('A3').setValue('使用');
let uuid = 0, rendered = 0;
const context = vm.createContext({
  SpreadsheetApp: {
    getActiveSpreadsheet: () => book,
    flush() {},
    getUi: () => ({ alert() {} }),
  },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  Utilities: { getUuid: () => 'part-' + (++uuid) },
  console,
});
const root = path.resolve(__dirname, '..');
for (const file of ['Logging.js', 'AnswerSheetGenerator.js', 'DocumentAnalysis.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context, { filename: file });
}
// 設定シートがなくても空の作業状態を開始できる。
book.sheets.delete('設定');
assert.equal(context.readDraft_().parts.length, 0);
assert.equal(book.getSheetByName('配置データ').isSheetHidden(), false);
assert.equal(context.readDraft_().parts.length, 0);
const old = book.getSheetByName('配置データ');
old.setName('作業データ');
old.hideSheet();
assert.equal(context.readDraft_().parts.length, 0);
assert.equal(book.getSheetByName('配置データ'), old);
assert.equal(old.isSheetHidden(), false);
book.insertSheet('作業データ');
assert.throws(() => context.readDraft_(), /両方あります/);
book.sheets.delete('作業データ');
vm.runInContext(fs.readFileSync(path.join(root, 'src', 'ActionButtons.js'), 'utf8'), context);
const menus = [];
context.removeLegacySheetButtons_ = () => {};
context.SpreadsheetApp.getUi = () => ({ createMenu(name) {
  menus.push(name);
  return { addItem() { return this; }, addToUi() { return this; } };
} });
context.onOpen();
assert.deepEqual(menus, ['パーツ編集']);
console.log('設定シートなしの作業開始とメニュー: OK');
