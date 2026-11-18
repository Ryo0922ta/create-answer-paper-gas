/** T-01: 実際の設問文と期待する解答欄を、現在の解析結果と比較する。 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const sample = JSON.parse(fs.readFileSync(
  path.join(root, 'サンプル問題/IT分野_問題構造_詳細版.json'), 'utf8'));
const expected = [
  { number: '1', groups: [
    { path: '1-(1)', labels: ['(a)', '(b)', '(c)', '(d)', '(e)'], frames: 5, pattern: '小問複合枠', status: '自動' },
    { path: '1-(2)', labels: ['①', '②'], frames: 2, pattern: '小問複合枠', status: '自動' },
  ] },
  { number: '2', groups: [
    { path: '2-(1)', labels: ['①', '②', '③', '④'], frames: 4, pattern: '小問複合枠', status: '自動' },
    { path: '2-(2)', labels: ['(A)', '(B)', '(C)', '(D)'], frames: 4, pattern: '小問複合枠', status: '自動' },
  ] },
  { number: '3', groups: [
    { path: '3-(1)', labels: ['(ア)', '(イ)', '(ウ)'], frames: 3, pattern: '小問複合枠', status: '自動' },
    { path: '3-(2)', labels: ['①', '②'], frames: 2, pattern: '小問複合枠', status: '自動' },
  ] },
];

// JSONにない表記・問題を補い、T-01の五種類を一通り確認できるようにする。
const additional = [
  { name: '丸数字の範囲表記', text: '大問6\n(1) 空欄①〜④に当てはまる語句を書きなさい。',
    groups: [{ path: '6-(1)', labels: ['①', '②', '③', '④'], frames: 4, pattern: '小問複合枠', status: '自動' }] },
  { name: '記号のない記述問題', text: '大問4\n(1) 認証が重要な理由を説明しなさい。',
    groups: [{ path: '4-(1)', labels: [], frames: 1, pattern: '記述欄', status: '自動' }] },
  { name: '文章から二つの枠を推定', text: '大問5\n(1) 安全なパスワードの条件について、2つの例を挙げなさい。',
    groups: [{ path: '5-(1)', labels: [], frames: 2, pattern: '記述欄', status: '要確認' }] },
  { name: '選択肢だけでは解答記号を確定しない', text: '大問7\n(1) 正しいものを一つ選びなさい。\n【選択肢】\nア. A\nイ. B',
    groups: [{ path: '7-(1)', labels: [], frames: 1, pattern: '記述欄', status: '要確認' }] },
];

function questionText(includeChoices) {
  return Object.entries(sample.questions).map(function ([major, rows]) {
    return [major].concat(rows.map(function (row) {
      return [row['問題番号'] + ' ' + row['設問文'],
        includeChoices ? '【選択肢】\n' + row['選択肢群'] : ''].filter(Boolean).join('\n');
    })).join('\n');
  }).join('\n');
}

function normalizeLabel(label) {
  return String(label).replace(/[（）Ａ-Ｚａ-ｚ０-９]/g, function (character) {
    return character.normalize('NFKC');
  }).replace(/[\s　]+/g, '');
}

function compactGroup(group) {
  return {
    path: group.path,
    labels: group.labels.map(normalizeLabel),
    frames: group.frameCount,
    pattern: group.pattern,
    status: group.status,
  };
}

// Apps Script の解析関数だけをローカルの隔離された文脈で実行する。
const context = vm.createContext({});
['Logging.js', 'AnswerSheetGenerator.js', 'DocumentAnalysis.js'].forEach(function (file) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context, { filename: file });
});

const differences = [];
const totals = new Map(sample.pattern_analysis.map(function (item) {
  return [item['大問'], item['記述解答箇所数']];
}));
expected.forEach(function (major) {
  assert.equal(major.groups.reduce((sum, group) => sum + group.frames, 0),
    totals.get('大問' + major.number), 'サンプル集計と期待枠数が一致しません: 大問' + major.number);
});

function compare(name, actual, wanted) {
  if (JSON.stringify(actual) !== JSON.stringify(wanted)) {
    differences.push({ name, expected: wanted, actual });
  }
}

[false, true].forEach(function (includeChoices) {
  const mode = includeChoices ? '選択肢を含む' : '設問文のみ';
  const result = context.analyzeDocumentText(questionText(includeChoices));
  compare(mode + '／大問数', result.length, expected.length);
  expected.forEach(function (major, index) {
    const actual = result[index];
    if (!actual) return;
    compare(mode + '／大問' + major.number + 'の番号', actual.number, major.number);
    compare(mode + '／大問' + major.number + 'の枠数', actual.totalFrames,
      major.groups.reduce((sum, group) => sum + group.frames, 0));
    compare(mode + '／大問' + major.number + 'の小問',
      actual.groups.map(compactGroup), major.groups);
  });
});

additional.forEach(function (item) {
  const result = context.analyzeDocumentText(item.text);
  compare(item.name, result[0].groups.map(compactGroup), item.groups);
});

if (differences.length) {
  differences.forEach(function (item) {
    console.log('差分: ' + item.name);
    console.log('  期待: ' + JSON.stringify(item.expected));
    console.log('  現行: ' + JSON.stringify(item.actual));
  });
  console.log('合計 ' + differences.length + ' 件の差分。判定ルールの修正対象です。');
  if (process.argv.includes('--strict')) process.exitCode = 1;
} else {
  console.log('T-01 の期待結果と現行解析が一致しました。');
}
