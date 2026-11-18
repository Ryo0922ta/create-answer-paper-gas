/** T-02: 解析候補のプレビューと出力描画が同じ枠配置を使うことを確認する。 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const sample = JSON.parse(fs.readFileSync(
  path.join(root, 'サンプル問題/IT分野_問題構造_詳細版.json'), 'utf8'));
const source = Object.entries(sample.questions).map(function ([major, rows]) {
  return [major].concat(rows.map(function (row) {
    return row['問題番号'] + ' ' + row['設問文'];
  })).join('\n');
}).join('\n');

const context = vm.createContext({});
['Logging.js', 'AnswerSheetGenerator.js', 'DocumentAnalysis.js'].forEach(function (file) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context, { filename: file });
});

const candidates = context.analyzeDocumentText(source);
const preview = context.previewAnalyzedDocument(candidates);
assert.equal(JSON.stringify(preview.map(q => q.number)), JSON.stringify(['1', '2', '3']));
assert.equal(JSON.stringify(preview.map(q => q.totalFrames)), JSON.stringify([7, 8, 5]));

const drawn = [];
context.drawBox_ = function (_sheet, _row, _column, _height, width, value) {
  drawn.push({ width, value });
};
preview.forEach(function (question, qi) {
  question.groups.forEach(function (group, gi) {
    assert.equal(group.error, undefined);
    assert.equal(group.boxes.reduce((sum, box) => sum + box.width, 0), question.width);
    assert.equal(group.boxes.filter(box => box.kind === 'answer').length, group.frameCount);
    drawn.length = 0;
    context.drawAnswerRow_({}, 1, 1, question.width, candidates[qi].groups[gi]);
    assert.equal(JSON.stringify(drawn), JSON.stringify(group.boxes.map(box => ({
      width: box.width, value: box.value,
    }))));
  });
});
assert.doesNotThrow(() => context.validateAnswerWidths_({ width: 33 }, {
  labels: ['①', '②'], weights: [1, 2], sheetRow: 1,
}));

const changed = context.analyzeDocumentText(source);
changed[0].order = 4;
changed[0].width = 25;
changed[0].groups[0].weights = [3, 1, 1, 1, 1];
changed[1].groups[0].labels[0] = 'X';
const edited = context.previewAnalyzedDocument(changed);
assert.equal(JSON.stringify(edited.map(q => q.number)), JSON.stringify(['2', '3', '1']));
assert.equal(edited[0].groups[0].boxes[1].value, 'X');
assert.equal(edited[2].groups[0].boxes.reduce((sum, box) => sum + box.width, 0), 25);
assert.notEqual(JSON.stringify(edited[2].groups[0].boxes), JSON.stringify(preview[0].groups[0].boxes));

changed[1].groups[0].weights = [0, 1, 1, 1];
assert.match(context.previewAnalyzedDocument(changed)[0].groups[0].error, /幅比/);
const essay = context.previewAnalyzedDocument([{ number: '4', order: 1, width: 20,
  groups: [{ path: '4-(1)', rowLabel: '1', pattern: '記述欄', frameCount: 2, height: 2 }] }]);
assert.equal(essay[0].groups[0].boxes.filter(box => box.kind === 'answer').length, 2);
assert.equal(essay[0].groups[0].boxes.reduce((sum, box) => sum + box.width, 0), 20);

console.log('T-02 の枠数・ラベル・順番・幅と出力描画の一致: OK');
