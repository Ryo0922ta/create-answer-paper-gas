const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const logs = [];
const context = vm.createContext({
  console: Object.fromEntries(['info', 'warn', 'error'].map(level =>
    [level, message => logs.push({ level, ...JSON.parse(message) })])),
});
const root = path.resolve(__dirname, '..');
for (const file of ['Logging.js', 'AnswerSheetGenerator.js', 'DocumentAnalysis.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context, { filename: file });
}
const sensitive = '秘密の問題文';
const candidate = context.analyzeDocumentText('大問1\n(1) ' + sensitive);
assert.equal(candidate.length, 1);
assert.ok(logs.some(x => x.operation === 'analyze_document' && x.event === 'start'));
assert.ok(logs.some(x => x.operation === 'analyze_document' && x.event === 'complete' && x.questionCount === 1));
assert.throws(() => context.analyzeDocumentText(''));
assert.ok(logs.some(x => x.operation === 'analyze_document' && x.event === 'failed'));

context.readDraft_ = () => ({ parts: [] });
context.validateDraft_ = () => [];
context.updateOutputParts_ = () => { throw new Error(sensitive); };
assert.throws(() => context.commitDraft_({ parts: [] }));
assert.ok(logs.some(x => x.operation === 'update_output' && x.event === 'failed'));

context.LockService = { getDocumentLock: () => ({ tryLock: () => false }) };
assert.throws(() => context.addBasicPart('考査見出し'));
assert.ok(logs.some(x => x.operation === 'draft_lock' && x.event === 'timeout'));
assert.ok(logs.some(x => x.operation === 'add_basic' && x.event === 'failed'));
assert.equal(JSON.stringify(logs).includes(sensitive), false);
assert.ok(logs.filter(x => x.event === 'complete' || x.event === 'failed')
  .every(x => typeof x.durationMs === 'number'));

console.log('操作ログ・入力値の除外・出力復旧・ロック失敗: OK');
