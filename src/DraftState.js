/** 出力シートで編集中の内容。設定シートとは独立して保持する。 */
const DRAFT = {
  sheetName: '作業データ',
  version: 'answer-paper-draft-v1',
  chunkSize: 20000,
  basicHeaderRow: 64,
  basicFirstRow: 65,
};

function withDraftLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    logEvent_('warn', 'draft_lock', 'timeout');
    throw new Error('別の更新中です。少し待ってください。');
  }
  try { return callback(); } finally { lock.releaseLock(); }
}

function draftSheet_(ss, create) {
  let sheet = ss.getSheetByName(DRAFT.sheetName);
  if (!sheet && create) {
    sheet = ss.insertSheet(DRAFT.sheetName);
    sheet.hideSheet();
  }
  return sheet;
}

function readDraft_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = draftSheet_(ss, false);
  if (!sheet) {
    const draft = initialDraft_(ss);
    writeDraft_(draft);
    logEvent_('info', 'draft_read', 'initialized', draftLogCounts_(draft));
    return draft;
  }
  const header = sheet.getRange('A1:B1').getValues()[0];
  if (header[0] !== DRAFT.version) {
    throw new Error('「作業データ」シートが別の形式です。内容を確認してください。');
  }
  const count = Number(header[1]);
  if (!Number.isInteger(count) || count < 1 || count > 100) {
    throw new Error('作業データの形式を確認してください。');
  }
  const raw = sheet.getRange(2, 1, count, 1).getValues()
    .map(function (row) { return String(row[0] || ''); }).join('');
  let draft;
  try { draft = JSON.parse(raw); }
  catch (_) { throw new Error('作業データを読み取れません。'); }
  if (!draft || !Array.isArray(draft.parts)) throw new Error('作業データの形式を確認してください。');
  logEvent_('info', 'draft_read', 'loaded', draftLogCounts_(draft));
  return draft;
}

function initialDraft_(ss) {
  return { parts: [] };
}

function writeDraft_(draft) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = draftSheet_(ss, true);
  const raw = JSON.stringify(draft);
  const chunks = raw.match(new RegExp('[\\s\\S]{1,' + DRAFT.chunkSize + '}', 'g')) || [''];
  if (chunks.length > 100) throw new Error('作業データが大きすぎます。');
  ensureSize_(sheet, chunks.length + 1, 2);
  const previous = Number(sheet.getRange('B1').getValue()) || 0;
  sheet.getRange(2, 1, Math.max(previous, chunks.length, 1), 1).clearContent();
  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks.map(function (chunk) { return [chunk]; }));
  sheet.getRange('A1:B1').setValues([[DRAFT.version, chunks.length]]);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
}

function validateDraft_(draft) {
  if (!draft || !Array.isArray(draft.parts)) throw new Error('作業データの形式を確認してください。');
  const ids = new Set(), numbers = new Set();
  let basicCount = 0, questionCount = 0, detailCount = 0;
  draft.parts.forEach(function (part) {
    if (!part.id || ids.has(part.id)) throw new Error('パーツの識別情報が重複しています。');
    ids.add(part.id);
    requiredInteger_(part.order, '順番', 1, 999);
    requiredInteger_(part.gap, '後の空き', 0, 10);
    if (part.kind === 'basic') {
      basicCount++;
      if (!['考査見出し', '年組氏名欄', '得点枠'].includes(part.type)) {
        throw new Error('基本パーツの種類を確認してください。');
      }
      requiredInteger_(part.width, '基本パーツの幅', part.type === '年組氏名欄' ? 8 : 3, 33);
      requiredInteger_(part.height, '基本パーツの高さ', 2, 20);
    } else if (part.kind === 'question') {
      questionCount++;
      const number = String(part.number || '').trim();
      if (!number || /^[=+@-]/.test(number) || numbers.has(number)) {
        throw new Error('大問番号が空欄または重複しています。');
      }
      numbers.add(number);
      requiredInteger_(part.width, '大問の幅', 6, 33);
      if (!Array.isArray(part.rows) || !part.rows.length) throw new Error('大問には小問を1件以上入力してください。');
      detailCount += part.rows.length;
      part.rows.forEach(function (row, index) {
        requiredInteger_(row.height, '小問の高さ', 1, 10);
        requiredInteger_(row.frameCount, '小問の枠数', 1, 30);
        requiredInteger_(row.itemsPerFrame, '1枠の記入項目数', 1, 10);
        const labels = row.labels || [];
        const weights = row.weights || [];
        if (!Array.isArray(labels) || !Array.isArray(weights) ||
            labels.some(function (label) { return !label || String(label).includes(','); }) ||
            (row.pattern === '小問複合枠' && labels.length !== row.frameCount)) {
          throw new Error('小問' + (index + 1) + 'のラベルと枠数を確認してください。');
        }
        if (row.pattern === '記述欄' && row.frameCount > part.width - 2) {
          throw new Error('小問' + (index + 1) + 'の枠数に対して幅が足りません。');
        }
        answerRowLayout_(part.width, row);
      });
      part.height = 1 + part.rows.reduce(function (sum, row) { return sum + row.height; }, 0);
    } else {
      throw new Error('パーツの種類を確認してください。');
    }
  });
  if (basicCount > EXAM.basicRowCount || questionCount > EXAM.questionRowCount ||
      detailCount > EXAM.detailRowCount) throw new Error('パーツの登録可能件数を超えています。');
  return calculatePlacements_(draft.parts.slice().sort(function (a, b) {
    return a.order - b.order || a.sheetRow - b.sheetRow;
  }));
}

function commitDraft_(draft, expectedSignature, previousDraft, selectedRemoval) {
  return withOperationLog_('update_output', function () {
    const previous = previousDraft || readDraft_();
    validateDraft_(draft);
    return updateOutputParts_(previous, draft, expectedSignature, selectedRemoval);
  }, function () { return draftLogCounts_(draft); });
}

function getDraftEditableParts_() {
  return readDraft_().parts.slice().sort(function (a, b) { return a.order - b.order; })
    .map(function (part) {
      if (part.kind === 'basic') return {
        id: part.id, kind: 'basic', order: part.order, type: part.type,
        label: part.type + (part.text ? '：' + part.text : ''), text: part.text,
        width: part.width, height: part.height, gap: part.gap, range: outputRect_(part) ? outputA1_(outputRect_(part)) : '',
      };
      return {
        id: part.id, kind: 'question', order: part.order, number: part.number,
        label: '大問 ' + part.number, title: part.title, rubric: part.rubric,
        points: part.points, width: part.width, height: part.height, gap: part.gap, range: outputRect_(part) ? outputA1_(outputRect_(part)) : '',
        detailCount: part.rows.length,
        details: part.rows.map(function (row) { return {
          rowLabel: row.rowLabel, labels: row.labels.join(','), weights: row.weights.join(','),
          height: row.height, pattern: row.pattern, path: row.path,
          frameCount: row.frameCount, itemsPerFrame: row.itemsPerFrame, evidence: row.evidence,
        }; }),
      };
    });
}

function listBasicTemplates() {
  const basic = getBasicSheet_();
  const rows = basic.getRange(EXAM.basicStartRow, 1, EXAM.basicRowCount, 7).getValues();
  return ['考査見出し', '年組氏名欄', '得点枠'].map(function (type) {
    const index = rows.findIndex(function (row) { return String(row[2] || '').trim() === type; });
    const values = index >= 0 ? rows[index] : [];
    return { type: type, sourceRow: index >= 0 ? EXAM.basicStartRow + index : '',
      text: String(values[3] || ''), width: Number(values[4]) || 33,
      height: Number(values[5]) || (type === '考査見出し' ? 2 : 4),
      gap: values[6] === '' || values[6] === undefined ? 1 : Number(values[6]) };
  });
}

function addBasicPart(type, previewOnly, expectedSignature) {
  return withOperationLog_('add_basic', function () {
    return addBasicPartLogged_(type, previewOnly, expectedSignature);
  });
}

function addBasicPartLogged_(type, previewOnly, expectedSignature) {
  return withDraftLock_(function () {
    const template = listBasicTemplates().find(function (item) { return item.type === type; });
    if (!template) throw new Error('基本パーツの種類を選んでください。');
    const draft = readDraft_();
    const nextOrder = draft.parts.reduce(function (max, part) { return Math.max(max, part.order); }, 0) + 1;
    draft.parts.push({ id: Utilities.getUuid(), kind: 'basic', sheetRow: draft.parts.length + 1,
      sourceRow: template.sourceRow, order: nextOrder, type: template.type,
      text: template.text, width: template.width, height: template.height, gap: template.gap });
    if (previewOnly === true) {
      const output = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAM.outputSheetName);
      if (!output) throw new Error('「出力」シートがありません。');
      validateDraft_(draft);
      const plan = outputPlan_(readDraft_(), draft, outputReadGrid_(output));
      logEvent_('info', 'add_basic', 'candidate', { range: outputA1_(plan.additions[0].rect) });
      return { range: outputA1_(plan.additions[0].rect), signature: outputPlanSignature_(plan) };
    }
    if (!expectedSignature) throw new Error('配置候補を確認してください。');
    commitDraft_(draft, expectedSignature);
    return template.type + 'を出力に配置しました。';
  });
}

function updateDraftBasicPart_(data) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const part = draft.parts.find(function (item) { return item.id === data.id && item.kind === 'basic'; });
    if (!part) throw new Error('編集する基本パーツが見つかりません。');
    part.order = requiredInteger_(data.order, '順番', 1, 999);
    part.width = requiredInteger_(data.width, '基本パーツの幅', part.type === '年組氏名欄' ? 8 : 3, 33);
    part.height = requiredInteger_(data.height, '基本パーツの高さ', 2, 20);
    part.gap = requiredInteger_(data.gap, '後の空き', 0, 10);
    part.text = String(data.text || '');
    commitDraft_(draft);
    return '基本パーツを更新しました。';
  });
}

function updateDraftQuestion_(data) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const part = draft.parts.find(function (item) { return item.id === data.id && item.kind === 'question'; });
    if (!part || part.number !== String(data.originalNumber || '')) {
      throw new Error('編集する大問が変わりました。再読み込みしてください。');
    }
    const number = String(data.number || '').trim();
    if (!number || /^[=+@-]/.test(number)) throw new Error('大問番号を確認してください。');
    part.number = number;
    part.order = requiredInteger_(data.order, '順番', 1, 999);
    part.width = requiredInteger_(data.width, '大問の幅', 6, 33);
    part.gap = requiredInteger_(data.gap, '後の空き', 0, 10);
    part.title = String(data.title || '');
    part.rubric = String(data.rubric || '');
    part.points = String(data.points || '');
    if (!Array.isArray(data.details) || !data.details.length) throw new Error('小問を1件以上入力してください。');
    part.rows = data.details.map(function (detail, index) {
      const labels = (detail.labels || []).map(function (label) { return String(label).trim(); });
      const weights = (detail.weights || []).map(Number);
      const pattern = answerPattern_(detail.pattern, labels.join(','), index + 1);
      return { sheetRow: index + 1, index: index + 1, rowLabel: String(detail.rowLabel || ''),
        labels: labels, weights: weights, height: requiredInteger_(detail.height, '小問の高さ', 1, 10),
        pattern: pattern, path: String(detail.path || number + '-' + (index + 1)),
        frameCount: requiredInteger_(detail.frameCount, '小問の枠数', 1, 30),
        itemsPerFrame: requiredInteger_(detail.itemsPerFrame, '記入項目数', 1, 10),
        evidence: String(detail.evidence || '') };
    });
    commitDraft_(draft);
    return '大問 ' + number + ' を更新しました。';
  });
}

function duplicateDraftPart_(id) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const source = draft.parts.find(function (part) { return part.id === id; });
    if (!source) throw new Error('複製するパーツが見つかりません。');
    const copy = JSON.parse(JSON.stringify(source));
    copy.id = Utilities.getUuid();
    copy.order = source.order + 1;
    if (copy.kind === 'question') {
      let number = 1;
      const used = new Set(draft.parts.filter(function (part) { return part.kind === 'question'; })
        .map(function (part) { return String(part.number); }));
      while (used.has(String(number))) number++;
      copy.number = String(number);
      copy.rows.forEach(function (row, index) { row.path = copy.number + '-' + (index + 1); });
    }
    draft.parts.push(copy);
    commitDraft_(draft);
    return 'パーツを複製しました。';
  });
}

function deleteDraftPart_(id, expectedRevision, rect) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const index = draft.parts.findIndex(function (part) { return part.id === id; });
    if (index < 0) throw new Error('削除するパーツが見つかりません。');
    if (expectedRevision && deletePartRevision_(draft.parts[index]) !== expectedRevision) {
      throw new Error('確認後に削除対象が変更されました。パーツ一覧を再確認してください。');
    }
    const previous = JSON.parse(JSON.stringify(draft));
    draft.parts.splice(index, 1);
    commitDraft_(draft, null, previous, rect ? {partId:id, rect:rect} : undefined);
    return 'パーツを出力から削除しました。';
  });
}

function moveDraftPart_(id, direction) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const part = draft.parts.find(function (item) { return item.id === id; });
    if (!part) throw new Error('移動するパーツが見つかりません。');
    const ordered = draft.parts.slice().sort(function (a, b) { return a.order - b.order; });
    const index = ordered.indexOf(part);
    const target = index + (direction === 'up' ? -1 : 1);
    if (target < 0 || target >= ordered.length) return 'これ以上移動できません。';
    ordered.splice(index, 1);
    ordered.splice(target, 0, part);
    ordered.forEach(function (item, i) { item.order = i + 1; });
    commitDraft_(draft);
    return '順番を変更しました。';
  });
}

