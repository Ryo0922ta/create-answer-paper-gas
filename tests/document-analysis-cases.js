/** T-01: 実際の設問文から解析した解答欄とプレビューを、期待値と比較する。 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
// 設計文書「基本設計/GAS/テストケース.md」の全小問と解答欄を転記したデータ。
const cases = require('./fixtures/document-analysis-cases.json');

// テストケース表にない範囲表記や枠数の推定も確認する。
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
  return cases.map(function (major) {
    return ['大問' + major.number].concat(major.groups.map(function (row) {
      const question = '(' + row.number + ') ' + row.question;
      return includeChoices && row.choices ? question + '\n【選択肢】\n' + row.choices : question;
    })).join('\n');
  }).join('\n');
}

function normalizeLabel(label) {
  // 資料は a、解析結果は（ a ）のように記す。丸数字はそのまま比較する。
  return String(label).replace(/[（）Ａ-Ｚａ-ｚ０-９]/g, function (character) {
    return character.normalize('NFKC');
  }).replace(/\s+/g, '').replace(/^\(([^()]+)\)$/, '$1');
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
const allRows = cases.flatMap(major => major.groups);
assert.equal(cases.length, 12, '資料の大問数');
assert.equal(allRows.length, 26, '資料の小問数');
assert.equal(allRows.reduce((sum, row) => sum + row.frames, 0), 84, '資料の総枠数');

function compare(name, actual, wanted) {
  if (JSON.stringify(actual) !== JSON.stringify(wanted)) {
    differences.push({ name, expected: wanted, actual });
  }
}

[false, true].forEach(function (includeChoices) {
  const mode = includeChoices ? '選択肢を含む' : '設問文のみ';
  const result = context.analyzeDocumentText(questionText(includeChoices));
  const preview = context.previewAnalyzedDocument(result);
  const numbers = cases.map(major => major.number);
  compare(mode + '／解析の大問番号と順序', result.map(major => major.number), numbers);
  compare(mode + '／プレビューの大問番号と順序', preview.map(major => major.number), numbers);
  compare(mode + '／解析の総枠数', result.reduce((sum, major) => sum + major.totalFrames, 0), 84);
  compare(mode + '／プレビューの総枠数', preview.reduce((sum, major) => sum + major.totalFrames, 0), 84);

  cases.forEach(function (major, index) {
    const actual = result[index];
    const rendered = preview[index];
    const location = mode + '／大問' + major.number;
    const totalFrames = major.groups.reduce((sum, row) => sum + row.frames, 0);
    compare(location + '／解析の枠数', actual?.totalFrames, totalFrames);
    compare(location + '／解析の小問数', actual?.groups.length, major.groups.length);
    compare(location + '／プレビューの枠数', rendered?.totalFrames, totalFrames);
    compare(location + '／プレビューの小問数', rendered?.groups.length, major.groups.length);
    compare(location + '／プレビューのエラー', rendered?.error, undefined);

    major.groups.forEach(function (row, rowIndex) {
      const rowPath = major.number + '-(' + row.number + ')';
      const rowLocation = location + ' 小問(' + row.number + ')';
      const labels = row.labels.map(normalizeLabel);
      const group = actual?.groups[rowIndex];
      const previewGroup = rendered?.groups[rowIndex];
      compare(rowLocation + '／解析', group && compactGroup(group), {
        path: rowPath,
        labels,
        frames: row.frames,
        pattern: labels.length ? '小問複合枠' : '記述欄',
        status: includeChoices && row.choices && !labels.length ? '要確認' : '自動',
      });
      compare(rowLocation + '／解析の小問番号', group?.rowLabel, row.number);
      compare(rowLocation + '／プレビューのエラー', previewGroup?.error, undefined);
      const boxes = previewGroup?.boxes;
      // 括弧付きの小問番号だけを持つ行は、番号の隣に記号なしの解答枠を表示する。
      compare(rowLocation + '／プレビューの解答欄', boxes && {
        path: previewGroup.path,
        numbers: boxes.filter(box => box.kind === 'number').map(box => box.value),
        labels: boxes.filter(box => box.kind === 'label').map(box => normalizeLabel(box.value)),
        answers: boxes.filter(box => box.kind === 'answer').map(box => box.value),
      }, {
        path: rowPath,
        numbers: [row.number],
        labels,
        answers: Array(row.frames).fill(''),
      });
    });
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
  console.log('T-01: 大問12件・小問26件・84枠の解析とプレビューが期待値と一致しました。');
}
