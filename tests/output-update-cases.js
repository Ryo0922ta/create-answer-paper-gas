const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({});
context.withOperationLog_ = (_operation, action) => action();
const root = path.resolve(__dirname, '..');
for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
}
const blank = { rows: [], merges: [] };
{
  const app = vm.createContext({withOperationLog_: (_name, action) => action(), logEvent_() {}});
  for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js']) {
    vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), app);
  }
  const types = ['考査見出し', '得点枠', '年組氏名欄'];
  let draftReads = 0, templateReads = 0, wholeReads = 0, cacheReads = 0, writes = 0, saved;
  let grid = blank, cache = blank, ranges = [], defaultCell;
  app.LockService = {getDocumentLock: () => ({tryLock: () => true, releaseLock() {}})};
  app.Utilities = {getUuid: () => 'new-basic'};
  app.readDraft_ = () => { draftReads++; return {parts:[]}; };
  app.listBasicTemplates = () => {
    templateReads++;
    return types.map(type => ({type, width:33, height:type === '考査見出し' ? 2 : 4, gap:1, text:'', placementCell:defaultCell}));
  };
  const sheet = {getName: () => '出力', getSheetId: () => 7};
  app.SpreadsheetApp = {getActiveSpreadsheet: () => ({getSheetByName: () => sheet, getId: () => 'book'}), flush() {}};
  app.outputReadRanges_ = (_sheet, requested) => { ranges.push(...requested.map(app.outputA1_)); return grid; };
  app.outputReadGrid_ = () => { wholeReads++; return grid; };
  app.outputCachedGrid_ = () => { cacheReads++; return cache; };
  app.outputCache_ = () => {};
  app.writeDraft_ = draft => { saved = draft; };
  app.Sheets = {Spreadsheets:{batchUpdate(request) {
    writes++;
    assert.equal(request.responseRanges.length, 1);
    assert.ok(request.includeSpreadsheetInResponse);
    const [, , rect] = JSON.parse(signature)[0];
    return {updatedSpreadsheet:{sheets:[{properties:{sheetId:7,gridProperties:{rowCount:51,columnCount:72}},
      data:[{startRow:rect.row - 1,startColumn:rect.column - 1}]}]}};
  }}};
  let signature;
  for (const [i, expected] of ['C3:AI4', 'AZ44:BT47', 'AZ48:BT51'].entries()) {
    const preview = app.addBasicPart(types[i], true);
    assert.equal(preview.range, expected);
    assert.equal(preview.placementPreview, true);
    signature = preview.signature;
    ranges = [];
    app.addBasicPart(types[i], false, signature);
    assert.deepEqual(ranges, [expected]);
    assert.equal(writes, i + 1);
    assert.equal(saved.parts[0].width, i === 0 ? 33 : 21);
  }
  assert.equal(draftReads, 6);
  assert.equal(templateReads, 6);
  assert.equal(wholeReads, 0);
  assert.equal(cacheReads, 1);
  cache = null;
  assert.equal(app.addBasicPart(types[0], true).range, 'C3:AI4');
  assert.equal(wholeReads, 1);

  ranges = [];
  const manual = app.addBasicPart(types[1], true, null, 'D12');
  assert.deepEqual(ranges, ['D12:AJ15']);
  signature = manual.signature;
  ranges = [];
  app.addBasicPart(types[1], false, signature, 'D12');
  assert.deepEqual(ranges, ['D12:AJ15']);
  assert.equal(writes, 4);
  assert.equal(saved.parts[0].width, 33);
  assert.equal(saved.parts[0].placementCell, undefined);
  assert.throws(() => app.addBasicPart(types[1], false, signature, 'D13'), /候補を再確認/);
  assert.throws(() => app.addBasicPart(types[1], true, null, 'BT51'), /用紙外/);
  grid = {rows:[], merges:[]};
  for (const [row, column] of [[12, 4], [44, 52]]) {
    grid.rows[row - 1] = {values:[]};
    grid.rows[row - 1].values[column - 1] = {note:'先に書き込まれたセル'};
  }
  assert.equal(app.addBasicPart(types[1], true).manualRequired, true);
  assert.throws(() => app.addBasicPart(types[1], true, null, 'D12'), /重なっています/);
  assert.throws(() => app.addBasicPart(types[1], false, signature, 'D12'), /配置先が変更/);
  assert.equal(writes, 4);
  assert.equal(wholeReads, 1);
  defaultCell = 'D16';
  const fromTemplate = app.addBasicPart(types[1], true);
  assert.equal(fromTemplate.range, 'D16:AJ19');
  defaultCell = 'D17';
  assert.throws(() => app.addBasicPart(types[1], false, fromTemplate.signature), /候補を再確認/);
  assert.equal(app.addBasicPart(types[1], true, null, 'E17').range, 'E17:AK20');
}
{
  const app = vm.createContext({withOperationLog_: (_name, action) => action(), logEvent_() {}});
  for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js', 'DraftState.js', 'DocumentAnalysis.js']) {
    vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), app);
  }
  const sheet = {getName: () => '出力', getSheetId: () => 7};
  const basicRows = [
    [true, ' c12 ', '考査見出し', '確認', 10, 2, 0],
    [true, 2, '年組氏名欄', '', 20, 4, 1],
    [true, '', '得点枠', '', 20, 4, 1],
  ];
  app.getBasicSheet_ = () => ({getRange: () => ({getValues: () => basicRows})});
  assert.deepEqual(Array.from(app.listBasicTemplates(), part => part.placementCell), ['C12', '', '']);
  app.SpreadsheetApp = {getActiveSpreadsheet: () => ({getSheetByName: () => sheet, getId: () => 'book'}), flush() {}};
  app.withDraftLock_ = action => action();
  app.deletePartRevision_ = part => JSON.stringify(part);
  app.Utilities = {getUuid: (() => { let id = 0; return () => String(++id); })()};
  let stored = {parts:[]}, grid = blank, rangeReads = [], writes = 0;
  app.readDraft_ = () => JSON.parse(JSON.stringify(stored));
  app.writeDraft_ = draft => { stored = draft; };
  app.outputReadRanges_ = (_sheet, rects) => { rangeReads.push(Array.from(rects, app.outputA1_)); return grid; };
  app.outputReadGrid_ = () => grid;
  app.outputCachedGrid_ = () => grid;
  app.outputCache_ = () => {};
  app.Sheets = {Spreadsheets:{batchUpdate(request) {
    writes++;
    return {updatedSpreadsheet:{sheets:[{properties:{sheetId:7,gridProperties:{rowCount:51,columnCount:72}},
      data:request.responseRanges.map(range => {
        const start = app.outputCellPosition_(range.split('!')[1].split(':')[0]);
        return {startRow:start.row - 1, startColumn:start.column - 1};
      })}]}};
  }}};
  const questions = app.analyzeDocumentText('大問3\n(1) 用語を答えなさい。\n大問4\n(1) 用語を答えなさい。');
  questions.forEach(q => { delete q.order; q.groups[0].height = 4; });
  questions[0].placementCell = 'C12';
  const rangesOf = preview => Array.from(preview.candidates, candidate => candidate.range);
  for (const gap of [0, 1, 2]) {
    questions[0].gap = gap;
    const result = app.importAnalyzedDocument({questions,previewOnly:true});
    assert.deepEqual(rangesOf(result), ['C12:AI16', 'C' + (17 + gap) + ':AI' + (21 + gap)]);
  }
  const pinned = app.importAnalyzedDocument({questions,positions:{'4':'D24'},previewOnly:true});
  assert.deepEqual(rangesOf(pinned), ['C12:AI16', 'D24:AJ28']);
  assert.deepEqual(rangeReads.at(-1), ['C12:AI16', 'D24:AJ28']);
  assert.throws(() => app.importAnalyzedDocument({questions,positions:{'4':'D16'},previewOnly:true}), /大問[34].*重なっています/);
  const automatic = [ {...questions[0], placementCell:''}, questions[1] ];
  const reserved = app.importAnalyzedDocument({questions:automatic,positions:{'4':'C3'},previewOnly:true});
  assert.deepEqual(rangesOf(reserved), ['C8:AI12', 'C3:AI7']);
  grid = {rows:[], merges:[]};
  const candidate = app.importAnalyzedDocument({questions,previewOnly:true});
  grid.rows[18] = {values:[,,{note:'先に書き込まれたセル'}]};
  assert.throws(() => app.importAnalyzedDocument({questions,placementSignature:candidate.signature}), /大問4.*C19:AI23/);
  assert.equal(writes, 0);
  const below = app.importAnalyzedDocument({questions,previewOnly:true});
  assert.deepEqual(rangesOf(below), ['C12:AI16', 'C20:AI24']);
  app.importAnalyzedDocument({questions,placementSignature:below.signature});
  assert.equal(writes, 1);
  assert.equal(stored.placementBatch, undefined);
  assert.equal(stored.parts[0].placementCell, undefined);
  assert.deepEqual(rangeReads.at(-1), ['C12:AI16', 'C20:AI24']);

  const heading = {id:'heading',kind:'basic',type:'考査見出し',width:10,height:2,gap:1,order:1,text:'見出し'};
  const source = {row:30,column:3,width:10,height:2};
  grid = {rows:[], merges:[{startRowIndex:29,endRowIndex:31,startColumnIndex:2,endColumnIndex:12}]};
  grid.rows[29] = {values:[,,{userEnteredValue:{stringValue:'見出し'}}]};
  heading.outputState = {rect:source,fingerprint:app.outputFingerprint_(grid,source)};
  stored = {parts:[heading]};
  const edit = {id:heading.id,originalRange:'C30:L31',placementCell:'C31',width:10,height:2,gap:1,text:heading.text};
  const move = app.updateDraftBasicPart_({...edit,previewOnly:true});
  assert.equal(move.range, 'C31:L32');
  assert.deepEqual(rangeReads.at(-1), ['C31:L32', 'C30:L31']);
  const confirmed = {...edit,placementSignature:move.signature,revision:move.revision};
  grid.rows[31] = {values:[,,{note:'既存内容'}]};
  assert.throws(() => app.updateDraftBasicPart_(confirmed), /C31:L32.*重なっています/);
  delete grid.rows[31];
  grid.rows[29].values[2].userEnteredValue.stringValue = '直接編集';
  assert.throws(() => app.updateDraftBasicPart_(confirmed), /直接編集/);
  assert.equal(writes, 1);
  grid.rows[29].values[2].userEnteredValue.stringValue = '見出し';
  app.updateDraftBasicPart_(confirmed);
  assert.equal(writes, 2);
  assert.equal(stored.parts[0].outputState.rect.row, 31);

  stored = {parts:[heading]};
  app.Sheets.Spreadsheets.batchUpdate = () => { throw new Error('書き込みの応答なし'); };
  app.restoreOutputRanges_ = () => assert.fail('実セルが変わっていない場合は復旧の書き込みをしない');
  const readsBeforeFailure = rangeReads.length;
  assert.throws(() => app.updateDraftBasicPart_(confirmed), /書き込みの応答なし/);
  assert.equal(rangeReads.length - readsBeforeFailure, 2);
  assert.equal(stored.parts[0].outputState.rect.row, 30);
}
{
  const oldRect = {row:12, column:3, height:2, width:3};
  const newRect = {...oldRect, row:13};
  const search = {row:11, column:2, height:6, width:6};
  function layout(rect, label) {
    const rows = [];
    rows[rect.row - 1] = {values:[]};
    rows[rect.row - 1].values[rect.column - 1] = {userEnteredValue:{stringValue:label}};
    return {rows, merges:[{startRowIndex:rect.row - 1, endRowIndex:rect.row,
      startColumnIndex:rect.column - 1, endColumnIndex:rect.column + rect.width - 1}]};
  }
  for (const [kind, label] of [['question', '大問 2'], ['basic', '得点欄']]) {
    const part = {id:kind, kind, outputState:{rect:oldRect,
      fingerprint:context.outputFingerprint_(layout(oldRect, label), oldRect)}};
    const signatures = context.outputDeleteSignatures_([part]);
    const moved = layout(newRect, label);
    const selected = {row:14, column:4};
    const candidate = context.outputFindDeleteCandidate_(moved, signatures, selected, search);
    assert.equal(candidate.signature.part.id, kind);
    assert.equal(context.outputA1_(candidate.rect), 'C13:E14');
    const selection = {partId:kind, rect:candidate.rect};
    const previous = {parts:[part]};
    const next = {parts:[]};
    assert.equal(context.outputSelectedRemovalPlan_(previous, next, selection, moved).removals[0].rect.row, 13);
    assert.equal(part.outputState.rect.row, 12);
    assert.throws(() => context.outputSelectedRemovalPlan_(previous, next, selection, layout(newRect, '直接編集')), /確認後/);
    assert.throws(() => context.outputFindDeleteCandidate_(moved, signatures, {row:15,column:4}, search), /特定できません/);
    assert.throws(() => context.outputFindDeleteCandidate_(moved, signatures.concat(signatures), selected, search), /複数/);
    const copyRect = {...oldRect, row:20};
    const copied = layout(copyRect, label);
    const original = layout(oldRect, label);
    copied.rows[oldRect.row - 1] = original.rows[oldRect.row - 1];
    copied.merges.push(...original.merges);
    assert.throws(() => context.outputSelectedRemovalPlan_(previous, next, {partId:kind, rect:copyRect}, copied), /登録位置にも/);

    const updateContext = vm.createContext({
      withOperationLog_: (_operation, action) => action(),
      logEvent_() {},
    });
    for (const file of ['AnswerSheetGenerator.js', 'OutputUpdate.js']) {
      vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), updateContext);
    }
    const requests = [];
    updateContext.SpreadsheetApp = {getActiveSpreadsheet:() => ({
      getSheetByName:() => ({getSheetId:() => 7}), getId:() => 'book',
    }), flush() {}};
    updateContext.Sheets = {Spreadsheets:{batchUpdate: body => requests.push(...body.requests)}};
    updateContext.outputReadRanges_ = (_sheet, ranges) => {
      assert.equal(ranges[0].row, newRect.row);
      assert.equal(ranges[1].row, oldRect.row);
      return moved;
    };
    updateContext.writeDraft_ = draft => assert.equal(draft.parts.length, 0);
    updateContext.updateOutputParts_(previous, next, null, selection);
    assert.ok(requests.length > 0);
    assert.ok(requests.every(item => Object.values(item)[0].range.startRowIndex === newRect.row - 1));
  }
}
{
  const operations = new Map();
  const id = '12345678-1234-1234-1234-123456789abc';
  context.CacheService = {getUserCache: () => ({
    put: (key, value) => operations.set(key, value),
    get: key => operations.get(key),
  })};
  context.LockService = {getUserLock: () => ({waitLock(){}, releaseLock(){}})};
  const originalDelete = context.deleteDraftPart_;
  let deletes = 0;
  const rect = {row:4,column:3,height:2,width:8};
  context.deleteDraftPart_ = (partId, revision, currentRect) => {
    assert.equal(partId, 'p');
    assert.equal(revision, 'revision');
    assert.equal(currentRect.row, rect.row);
    deletes++;
    return '削除しました。';
  };
  const pending = {state:'pending', partId:'p', revision:'revision', rect};
  context.putDeleteOperation_(id, pending, 120);
  assert.equal(context.cancelDeletePart(id).state, 'cancelled');
  assert.throws(() => context.confirmDeletePart(id), /終了済み/);
  context.putDeleteOperation_(id, pending, 120);
  assert.equal(context.confirmDeletePart(id), '削除しました。');
  assert.equal(context.getDeletePartStatus(id).state, 'complete');
  assert.throws(() => context.confirmDeletePart(id), /終了済み/);
  assert.equal(deletes, 1);
  context.deleteDraftPart_ = originalDelete;
  delete context.CacheService;
  delete context.LockService;
}
context.listBasicTemplates = () => [{ type: '年組氏名欄', height: 4 }];
const namePart = { id: 'name', kind: 'basic', type: '年組氏名欄', width: 20, height: 4, order: 1, gap: 1 };
const scorePart = { id: 'score', kind: 'basic', type: '得点枠', width: 15, height: 4, order: 2, gap: 1 };
const firstQuestion = { id: 'first-q', kind: 'question', number: '1', order: 3,
  gap: 1, width: 33, rows: [{ height: 2, pattern: '記述欄', frameCount: 1,
    itemsPerFrame: 1, labels: [], weights: [] }] };
const newFileDraft = { parts: [namePart, scorePart, firstQuestion] };
context.validateDraft_(newFileDraft);
const newFilePlan = context.outputPlan_({parts: []}, newFileDraft, blank);
assert.equal(newFilePlan.removals.length, 0);
assert.equal(newFilePlan.additions.length, 3);
assert.equal(newFilePlan.additions.find(entry => entry.part.id === 'first-q').rect.row, 3);
const manualQuestion = { ...firstQuestion, placementCell: 'd12' };
const manual = context.outputPlan_({parts: []}, {parts: [manualQuestion]}, blank);
assert.equal(context.outputA1_(manual.additions[0].rect), 'D12:AJ14');
assert.throws(() => context.outputPlan_({parts: []}, {parts: [{...manualQuestion, placementCell:'BT51'}]}, blank), /用紙外/);
assert.throws(() => context.outputPlan_({parts: []}, {parts: [{...manualQuestion, placementCell:'C0'}]}, blank), /左上のセル/);
const manualBlocked = {rows: Array.from({length: 12}, () => ({})), merges: []};
manualBlocked.rows[11] = {values: [{}, {}, {}, {userEnteredValue:{stringValue:'既存'}}]};
assert.throws(() => context.outputPlan_({parts: []}, {parts: [manualQuestion]}, manualBlocked), /重なっています/);
assert.throws(() => context.outputPlan_({parts: []}, {parts: [manualQuestion,
  {...manualQuestion, id:'second', number:'2', placementCell:'D12'}]}, blank), /重なっています/);
for (const pair of [[namePart, scorePart], [scorePart, namePart]]) {
  const plan = context.outputPlan_({parts: []}, {parts: pair}, blank);
  const nameRect = plan.additions.find(entry => entry.part.type === '年組氏名欄').rect;
  const scoreRect = plan.additions.find(entry => entry.part.type === '得点枠').rect;
  assert.equal(nameRect.row + nameRect.height - 1, 51);
  assert.equal(scoreRect.row + scoreRect.height, nameRect.row);
  assert.equal(nameRect.column, 52);
  assert.equal(scoreRect.column, 52);
  assert.equal(scoreRect.row, 44);
  assert.equal(nameRect.row, 48);
}
for (const pair of [[namePart, scorePart], [scorePart, namePart]]) {
  const first = context.outputPlan_({parts: []}, {parts: [pair[0]]}, blank).additions[0];
  const existing = { ...pair[0], outputState: { rect: first.rect,
    fingerprint: context.outputFingerprint_(blank, first.rect) } };
  const next = context.outputPlan_({parts: [existing]}, {parts: [existing, pair[1]]}, blank);
  assert.equal(next.removals.length, 0);
  const nameRect = pair[1].type === '年組氏名欄' ? next.additions[0].rect : first.rect;
  const scoreRect = pair[1].type === '得点枠' ? next.additions[0].rect : first.rect;
  assert.equal(scoreRect.row + scoreRect.height, nameRect.row);
}
const wideName = { ...namePart, id: 'wide', width: 33 };
const widePlan = context.outputPlan_({parts: []}, {parts: [wideName]}, blank);
assert.equal(widePlan.additions[0].rect.width, 21);
assert.equal(wideName.width, 21);
const occupiedBottom = {rows: Array.from({length: 51}, () => ({})), merges: []};
occupiedBottom.rows[50] = {values: Array.from({length: 71}, (_, i) =>
  i === 70 ? {userEnteredValue: {stringValue: '既存'}} : {})};
assert.throws(() => context.outputPlan_({parts: []}, {parts: [namePart]}, occupiedBottom), /右下の配置場所/);
const rect = { row: 3, column: 3, width: 8, height: 2 };
const part = { id: 'a', kind: 'basic', type: '考査見出し', order: 1, gap: 1,
  width: 8, height: 2, text: '考査',
  outputState: { rect, fingerprint: context.outputFingerprint_(blank, rect) } };
const other = { id: 'b', kind: 'basic', type: '考査見出し', order: 2, gap: 1,
  width: 8, height: 2, text: '追加' };
const direct = { rows: Array.from({length: 5}, () => ({})), merges: [] };
direct.rows[2] = { values: [{}, {}, { userEnteredValue: { stringValue: '直接編集' } }] };
assert.throws(() => context.outputPlan_({parts: [part]}, {parts: []}, direct), /直接編集/);
const appended = context.outputPlan_({parts: [part]}, {parts: [part, other]}, direct);
assert.equal(appended.removals.length, 0);
assert.equal(appended.additions.length, 1);
assert.equal(context.outputOverlap_(rect, appended.additions[0].rect), false);
const third = { ...other, id: 'c', order: 3 };
const multiple = context.outputPlan_({parts: []}, {parts: [other, third]}, blank);
assert.equal(multiple.additions.length, 2);
assert.equal(context.outputOverlap_(multiple.additions[0].rect, multiple.additions[1].rect), false);
const drawPlan = context.outputPlan_({parts:[]}, {parts:[firstQuestion]}, blank);
const requests = context.outputDrawRequests_(7, drawPlan);
assert.ok(requests.some(request => request.mergeCells));
assert.ok(requests.some(request => request.updateBorders));
assert.ok(requests.some(request => request.updateCells && request.updateCells.rows &&
  request.updateCells.rows[0].values[0].userEnteredValue.stringValue === '大問 1'));
assert.ok(requests.every(request => !request.repeatCell || request.repeatCell.range.sheetId === 7));
const confirmedSig = context.outputPlanSignature_(drawPlan);
const confirmed = context.outputConfirmedAdditions_({parts:[]}, {parts:[firstQuestion]}, confirmedSig);
assert.equal(confirmed.additions[0].rect.row, 3);
assert.throws(() => context.outputConfirmedAdditions_({parts:[]}, {parts:[firstQuestion]},
  JSON.stringify([['question','1',{row:50,column:3,width:33,height:3}]])), /変更/);
assert.throws(() => context.outputConfirmedAdditions_({parts:[]}, {parts:[firstQuestion]},
  JSON.stringify([['question','2',drawPlan.additions[0].rect]])), /変更/);
assert.equal(context.outputConfirmedAdditions_({parts:[firstQuestion]},
  {parts:[{...firstQuestion, gap:2}]}, confirmedSig), null);
// 確定済みの新規大問は事前に配置先を読み、更新結果は一括応答で取得する。
{
  const saved = { readRanges:context.outputReadRanges_, readGrid:context.outputReadGrid_,
    spreadsheet:context.SpreadsheetApp, sheets:context.Sheets, write:context.writeDraft_, log:context.logEvent_ };
  let reads = 0, writes = 0, recorded;
  const outputSheet = {getSheetId:()=>7,getName:()=> '出力'};
  context.SpreadsheetApp = {getActiveSpreadsheet:()=>({getSheetByName:()=>outputSheet,getId:()=> 'book'}),flush(){}};
  context.Sheets = {Spreadsheets:{batchUpdate:(request,_id,options)=>{
    writes++;
    assert.equal(request.includeSpreadsheetInResponse,true);
    assert.ok(options.fields.includes('updatedSpreadsheet'));
    return {updatedSpreadsheet:{sheets:[{properties:{sheetId:7,gridProperties:{rowCount:51,columnCount:72}},
      data:[{startRow:2,startColumn:2,rowData:[{values:[{userEnteredValue:{stringValue:'大問 1'}}]}]}]}]}};
  }}};
  context.outputReadGrid_ = () => { throw new Error('範囲全体を読みました'); };
  context.outputReadRanges_ = (_sheet, ranges) => { reads++; assert.equal(ranges.length, 1); return blank; };
  context.logEvent_ = () => {};
  context.writeDraft_ = draft => { recorded = draft; };
  const next = {parts:[JSON.parse(JSON.stringify(firstQuestion))]};
  context.updateOutputParts_({parts:[]}, next, confirmedSig);
  assert.equal(reads, 1); assert.equal(writes, 1);
  assert.equal(recorded.parts[0].outputState.rect.row, 3);
  const oldQuestion = JSON.parse(JSON.stringify(firstQuestion));
  oldQuestion.outputState = {rect:drawPlan.additions[0].rect,
    fingerprint:context.outputFingerprint_(blank, drawPlan.additions[0].rect)};
  const changedQuestion = JSON.parse(JSON.stringify(oldQuestion));
  changedQuestion.rows[0].rowLabel = '変更';
  const replaceSig = context.outputPlanSignature_({additions:[{part:changedQuestion, rect:context.outputRect_(oldQuestion)}]});
  assert.ok(context.outputConfirmedReplacements_({parts:[oldQuestion]}, {parts:[changedQuestion]}, replaceSig));
  context.updateOutputParts_({parts:[oldQuestion]}, {parts:[changedQuestion]}, replaceSig);
  assert.equal(reads, 2); assert.equal(writes, 2);
  context.outputReadRanges_ = () => ({rows:[{},{},{values:[{},{},{userEnteredValue:{stringValue:'直接編集'}}]}],merges:[]});
  assert.throws(() => context.updateOutputParts_({parts:[oldQuestion]}, {parts:[changedQuestion]}, replaceSig), /直接編集/);
  context.outputReadRanges_ = () => ({rows:[{},{},{values:[{},{},{effectiveValue:{numberValue:42}}]}],merges:[]});
  assert.throws(() => context.updateOutputParts_({parts:[]}, {parts:[firstQuestion]}, confirmedSig), /配置先が変更/);
  assert.equal(writes, 2);
  context.outputReadRanges_=saved.readRanges; context.outputReadGrid_=saved.readGrid;
  context.SpreadsheetApp=saved.spreadsheet; context.Sheets=saved.sheets;
  context.writeDraft_=saved.write; context.logEvent_=saved.log;
}
const altered = JSON.parse(JSON.stringify(part));
// 削除は対象範囲を一度だけ読み、直接編集や範囲外にまたがる結合を検出する。
{
  const saved = { readRanges:context.outputReadRanges_, readGrid:context.outputReadGrid_,
    spreadsheet:context.SpreadsheetApp, sheets:context.Sheets, write:context.writeDraft_, log:context.logEvent_ };
  let reads = 0, writes = 0, savedDraft;
  const sheet = {getSheetId:()=>7, getName:()=> '出力'};
  context.SpreadsheetApp = {getActiveSpreadsheet:()=>({getSheetByName:()=>sheet,getId:()=> 'book'}),flush(){}};
  context.Sheets = {Spreadsheets:{batchUpdate:()=>{ writes++; return {}; }}};
  context.outputReadGrid_ = () => { throw new Error('全範囲の読み取り'); };
  context.outputReadRanges_ = (_sheet, ranges) => { reads++; assert.equal(ranges.length, 1); return blank; };
  context.writeDraft_ = draft => { savedDraft = draft; };
  context.logEvent_ = () => {};
  context.updateOutputParts_({parts:[part]}, {parts:[]});
  assert.equal(reads, 1); assert.equal(writes, 1); assert.equal(savedDraft.parts.length, 0);
  context.outputReadRanges_ = () => direct;
  assert.throws(() => context.updateOutputParts_({parts:[part]}, {parts:[]}), /直接編集/);
  context.outputReadRanges_ = () => ({rows:[], merges:[{startRowIndex:1,endRowIndex:4,startColumnIndex:2,endColumnIndex:3}]});
  assert.throws(() => context.updateOutputParts_({parts:[part]}, {parts:[]}), /またぐ結合/);
  assert.equal(writes, 1);
  const overlapping = {...part, id:'overlap'};
  assert.throws(() => context.outputOnlyRemovals_({parts:[part, overlapping]}, {parts:[overlapping]}), /安全に確認/);
  context.outputReadRanges_ = () => blank;
  context.Sheets.Spreadsheets.batchUpdate = (request) => {
    writes++;
    if (writes === 3) assert.ok(request.requests.some(item => item.updateCells?.rows));
  };
  context.writeDraft_ = draft => {
    if (!draft.parts.length) throw new Error('保存失敗');
    savedDraft = draft;
  };
  assert.throws(() => context.updateOutputParts_({parts:[part]}, {parts:[]}), /保存失敗/);
  assert.equal(writes, 3);
  assert.equal(savedDraft.parts[0].id, part.id);
  context.outputReadRanges_=saved.readRanges; context.outputReadGrid_=saved.readGrid;
  context.SpreadsheetApp=saved.spreadsheet; context.Sheets=saved.sheets;
  context.writeDraft_=saved.write; context.logEvent_=saved.log;
}
altered.text = '変更';
const replacement = context.outputPlan_({parts: [part]}, {parts: [altered]}, blank);
assert.equal(replacement.removals.length, 1);
assert.equal(replacement.additions[0].rect.row, rect.row);
const legacy = { ...part };
delete legacy.outputState;
assert.throws(() => context.outputPlan_({parts: [legacy]}, {parts: []}, blank), /確認できない/);
// 古い大問に高さが保存されていなくても、検証時に算出した高さだけを変更とみなさない。
const legacyQuestion = { id: 'legacy-q', kind: 'question', number: '1', order: 1,
  gap: 1, width: 20, rows: [{ height: 2 }] };
const normalizedQuestion = { ...legacyQuestion, height: 3 };
const newQuestion = { id: 'new-q', kind: 'question', number: '2', order: 2,
  gap: 1, width: 20, height: 3, rows: [{ height: 2 }] };
const legacyAddition = context.outputPlan_({parts: [legacyQuestion]},
  {parts: [normalizedQuestion, newQuestion]}, blank);
assert.equal(legacyAddition.removals.length, 0);
assert.equal(legacyAddition.additions.length, 1);
assert.equal(legacyAddition.additions[0].part.id, 'new-q');
const occupancyGrid = { rows: [{ values: [{userEnteredValue:{stringValue:'x'}},
  {userEnteredFormat:{borders:{bottom:{style:'SOLID'}}}}] }],
  merges: [{startRowIndex:2,endRowIndex:4,startColumnIndex:3,endColumnIndex:5}] };
const occupancy = context.outputSnapshot_(occupancyGrid).occupied;
let cachedValue;
context.CacheService = {getDocumentCache: () => ({put: (_key, value) => {cachedValue=value;}, get: () => cachedValue})};
const cacheSheet = {getSheetId: () => 42};
context.outputCache_(cacheSheet, occupancyGrid);
const cachedGrid = context.outputCachedGrid_(cacheSheet);
assert.equal(context.outputVacant_(cachedGrid, {row:1,column:1,height:1,width:1}, []), false);
assert.equal(context.outputVacant_(cachedGrid, {row:3,column:4,height:1,width:1}, []), false);
assert.equal(context.outputVacant_(cachedGrid, {row:5,column:6,height:1,width:1}, []), true);
cachedValue = '["broken"]';
assert.equal(context.outputCachedGrid_(cacheSheet), null);
delete context.CacheService;
for (let row = 1; row <= 5; row++) for (let col = 1; col <= 6; col++) {
  assert.equal(occupancy[row-1][col-1] === '0',
    context.outputVacant_(occupancyGrid, {row, column:col, width:1, height:1}, []));
}
const bordered = { rows: [{ values: [{ userEnteredFormat: { borders: { top: {style: 'SOLID'} } } }] }], merges: [] };
assert.equal(context.outputVacant_(bordered, {row:1,column:1,width:1,height:1}, []), false);
assert.equal(context.outputVacant_({ rows: [], merges: [{ endRowIndex: 2, endColumnIndex: 2 }] },
  {row:1,column:1,width:1,height:1}, []), false);
const gridSheet = { getParent: () => ({getId: () => 'book'}), getName: () => '出力',
  getSheetId: () => 7, getImages: () => [], getDrawings: () => [],
  getRange: () => { throw new Error('余分な範囲読み取り'); } };
context.Sheets = { Spreadsheets: { get: () => ({sheets: [{}]}) } };
assert.throws(() => context.outputReadGrid_(gridSheet), /完全に読み取れません/);
context.Sheets.Spreadsheets.get = () => ({sheets: [{ properties: {sheetId:7,
  gridProperties:{rowCount:51,columnCount:72}}, data:[{rowData:[{}]}] }]});
assert.equal(context.outputReadGrid_(gridSheet).rows.length, 1);
context.Sheets.Spreadsheets.get = () => ({sheets: [{properties:{sheetId:7,
  gridProperties:{rowCount:51,columnCount:72}}, data:[{rowData:[{values:[
    {userEnteredValue:{formulaValue:'=1'}}, {effectiveValue:{numberValue:1}},
    {userEnteredFormat:{borders:{bottom:{style:'SOLID'}}}}]}]}],
  merges:[{startRowIndex:2,endRowIndex:3,startColumnIndex:0,endColumnIndex:2}]}]});
const apiGrid = context.outputReadGrid_(gridSheet);
for (const col of [1,2,3]) assert.equal(context.outputVacant_(apiGrid, {row:1,column:col,height:1,width:1}, []), false);
assert.equal(context.outputVacant_(apiGrid, {row:3,column:2,height:1,width:1}, []), false);
assert.equal(context.outputVacant_(apiGrid, {row:2,column:2,height:1,width:1}, []), true);
assert.equal(context.outputFingerprint_(apiGrid,{row:1,column:2,width:1,height:1}),
  context.outputFingerprint_({rows:[{values:[{},{}]}],merges:[]},{row:1,column:2,width:1,height:1}));
let requested;
context.Sheets.Spreadsheets.get = (_id, options) => { requested = options.ranges; return {sheets:[{
  properties:{sheetId:7,gridProperties:{rowCount:51,columnCount:72}},
  data:[{startRow:9,startColumn:3,rowData:[{values:[{userEnteredValue:{stringValue:'既存'}}]}]},
    {startRow:20,startColumn:5,rowData:[]}],
  merges:[{startRowIndex:9,endRowIndex:11,startColumnIndex:4,endColumnIndex:6}]}]}; };
const targets = [{row:10,column:4,width:3,height:2},{row:21,column:6,width:2,height:2}];
const targeted = context.outputReadRanges_(gridSheet, targets);
assert.equal(requested.length, 2);
assert.match(requested[0], /D10:F11/);
assert.equal(context.outputVacant_(targeted, {row:10,column:4,width:1,height:1}, []), false);
assert.equal(context.outputVacant_(targeted, {row:11,column:5,width:1,height:1}, []), false);
assert.equal(context.outputVacant_(targeted, {row:21,column:6,width:2,height:2}, []), true);
context.Sheets.Spreadsheets.get = () => ({sheets:[{properties:{sheetId:7,
  gridProperties:{rowCount:51,columnCount:72}},data:[]}]});
assert.throws(() => context.outputReadRanges_(gridSheet, targets), /完全に読み取れません/);
context.Sheets.Spreadsheets.get = () => ({sheets: [{properties:{sheetId:7,
  gridProperties:{rowCount:51,columnCount:72}}, data:[{rowData:[{}]}]}]});
gridSheet.getImages = () => [{}];
assert.throws(() => context.outputReadGrid_(gridSheet), /画像または図形/);
let restoreRequest;
context.Sheets.Spreadsheets.batchUpdate = (body, id) => { restoreRequest = {body, id}; };
const restoreGrid = { rows: [{values: [{userEnteredValue:{stringValue:'保持'},
  userEnteredFormat:{borders:{top:{style:'SOLID'}}}}]}],
  merges: [{endRowIndex:1,endColumnIndex:2}] };
context.restoreOutputRanges_({getId:()=>'book'}, {getSheetId:()=>7}, restoreGrid,
  [{row:1,column:1,width:2,height:1}]);
assert.equal(restoreRequest.id, 'book');
assert.equal(restoreRequest.body.requests[1].updateCells.rows[0].values[0].userEnteredValue.stringValue, '保持');
assert.equal(restoreRequest.body.requests[2].mergeCells.range.endColumnIndex, 2);
// 追加時にシート全体をコピーせず、失敗時は変更範囲だけを復旧する。
let restoredOutput = false, draftWrites = 0;
const output = { getSheetId() { return 7; }, getRange() { return {}; } };
const book = { getSheetByName() { return output; }, getId() { return 'book'; } };
context.SpreadsheetApp = { getActiveSpreadsheet: () => book, flush() {} };
context.outputReadGrid_ = () => blank;
context.outputPlan_ = () => ({ removals: [], additions: [{part: other, rect}] });
context.drawBasic_ = () => {};
context.restoreOutputRanges_ = (_ss, _sheet, _grid, rects) => {
  restoredOutput = true;
  assert.equal(rects.length, 1);
  assert.equal(rects[0].row, rect.row);
};
context.writeDraft_ = () => { if (++draftWrites === 1) throw new Error('write failed'); };
context.logEvent_ = () => {};
assert.throws(() => context.updateOutputParts_({parts: []}, {parts: [other]}, '古い候補'), /再確認/);
assert.equal(restoredOutput, false);
assert.throws(() => context.updateOutputParts_({parts: []}, {parts: [other]}), /write failed/);
assert.equal(restoredOutput, true);
assert.equal(draftWrites, 2);
context.restoreOutputRanges_ = () => { throw new Error('restore failed'); };
draftWrites = 0;
assert.throws(() => context.updateOutputParts_({parts: []}, {parts: [other]}), /変更履歴/);

// 新規ファイルで基本パーツを保持しながら、問題文の解析から大問を追加する。
vm.runInContext(fs.readFileSync(path.join(root, 'src', 'DocumentAnalysis.js'), 'utf8'), context);
const baseName = { ...namePart, outputState: { rect: {row:48,column:52,width:20,height:4},
  fingerprint: context.outputFingerprint_(blank, {row:48,column:52,width:20,height:4}) } };
const baseScore = { ...scorePart, outputState: { rect: {row:44,column:52,width:15,height:4},
  fingerprint: context.outputFingerprint_(blank, {row:44,column:52,width:15,height:4}) } };
let stored = { parts: [baseName, baseScore] };
context.readDraft_ = () => JSON.parse(JSON.stringify(stored));
context.withDraftLock_ = action => action();
context.Utilities = { getUuid: (() => { let i = 0; return () => 'new-' + (++i); })() };
context.outputReadGrid_ = () => blank;
// 上の復旧テスト用モックを元の配置ロジックに戻す。
vm.runInContext(fs.readFileSync(path.join(root, 'src', 'OutputUpdate.js'), 'utf8'), context);
context.outputReadGrid_ = () => blank;
context.commitDraft_ = (draft, signature) => {
  const plan = context.outputPlan_(stored, draft, blank);
  assert.equal(context.outputPlanSignature_(plan), signature);
  assert.equal(plan.removals.length, 0);
  stored = draft;
};
const questions = context.analyzeDocumentText('大問1\n(1) 用語を答えなさい。');
const candidate = context.importAnalyzedDocument({ questions, previewOnly: true });
assert.equal(candidate.placementPreview, true);
assert.equal(candidate.candidates.length, 1);
assert.match(candidate.candidates[0].range, /^C3:/);
const savedCachedGrid = context.outputCachedGrid_;
const staleGrid = {rows:[], merges:[], stale:true};
context.outputCachedGrid_ = () => staleGrid;
const originalFindSpace = context.outputFindSpace_;
context.outputFindSpace_ = (grid, ...args) => {
  if (grid.stale) throw new Error('新しいパーツを配置できる空きがありません。出力シートを確認してください。');
  return originalFindSpace(grid, ...args);
};
let fullReads = 0;
context.outputReadGrid_ = () => { fullReads++; return blank; };
assert.equal(context.importAnalyzedDocument({questions,previewOnly:true}).manualRequired, true);
assert.equal(fullReads, 0, 'キャッシュに空きがない場合は全範囲を再取得しない');
context.outputFindSpace_ = originalFindSpace;
context.outputCachedGrid_ = savedCachedGrid;
const findSpace = context.outputFindSpace_;
context.outputFindSpace_ = () => { throw new Error('新しいパーツを配置できる空きがありません。出力シートを確認してください。'); };
const noAutoSpace = context.importAnalyzedDocument({ questions, previewOnly: true });
assert.equal(noAutoSpace.manualRequired, true);
assert.equal(noAutoSpace.numbers[0], '1');
context.outputFindSpace_ = findSpace;
let manualReads = 0;
context.outputReadRanges_ = (_sheet, rects) => { manualReads++; assert.equal(rects.length, 1); return blank; };
const customCandidate = context.importAnalyzedDocument({ questions, previewOnly: true, positions: {'1':'D12'} });
assert.match(customCandidate.candidates[0].range, /^D12:/);
assert.equal(manualReads, 1);
context.outputReadRanges_ = () => manualBlocked;
assert.throws(() => context.importAnalyzedDocument({questions, previewOnly:true, positions:{'1':'D12'}}), /重なっています/);
context.outputReadRanges_ = () => blank;
assert.throws(() => context.importAnalyzedDocument({ questions, previewOnly: true, positions: {'1':'BT51'} }), /用紙外/);
context.importAnalyzedDocument({ questions, positions: {'1':'D12'}, placementSignature: customCandidate.signature });
assert.equal(stored.parts.length, 3);
assert.equal(stored.parts.find(item => item.id === 'name').outputState.rect.row, 48);
assert.equal(stored.parts.find(item => item.id === 'score').outputState.rect.row, 44);
{
  const cell = {getNumRows:()=>1, getNumColumns:()=>1, getA1Notation:()=> 'D12'};
  const sheet = {getName:()=> '出力', getActiveRangeList:()=>({getRanges:()=>[cell]})};
  context.SpreadsheetApp = {getActiveSpreadsheet:()=>({getActiveSheet:()=>sheet})};
  assert.equal(context.getSelectedOutputCell(), 'D12');
  sheet.getName = () => '基本パーツ';
  assert.throws(() => context.getSelectedOutputCell(), /出力シート/);
  sheet.getName = () => '出力'; cell.getNumRows = () => 2;
  assert.throws(() => context.getSelectedOutputCell(), /１つ/);
}
{
  let written;
  const savedPart = {...part};
  context.readDraft_ = () => ({parts:[savedPart]});
  context.writeDraft_ = draft => { written = draft; };
  context.SpreadsheetApp = {getActiveSpreadsheet:()=>({getSheetByName:()=>({})})};
  context.outputReadRanges_ = () => blank;
  assert.throws(() => context.unregisterChangedPart('a'), /一致している/);
  assert.equal(written, undefined);
  context.outputReadRanges_ = () => direct;
  context.unregisterChangedPart('a');
  assert.equal(written.parts.length, 0);
}
console.log('直接編集の競合、空き領域、旧パーツ保護: OK');
