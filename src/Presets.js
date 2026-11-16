/**
 * パーツプリセット。A:J がパーツ本体、L:W が大問の小問行。
 * B列の名前を変更すると、設定側H列の選択は新しい名前へ選び直す必要がある。
 */
const PRESET = { sheetName: 'パーツプリセット', firstRow: 4, maxChoices: 1000 };

/** 既存の設定値を消さずにプリセットシートと選択欄を追加する。 */
function setupPartPresets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensurePartPresetUi_(ss);
  SpreadsheetApp.getUi().alert('プリセット欄を用意しました。パーツ編集パネルからパーツを保存し、対象行のH列で選んでください。');
}

function ensurePartPresetUi_(ss) {
  const settings = ss.getSheetByName(EXAM.settingsSheetName);
  const basic = ss.getSheetByName(EXAM.basicSheetName);
  if (!settings || !basic || settings.getRange('A20').getDisplayValue() !== '使用' ||
      basic.getRange('A3').getDisplayValue() !== '使用') {
    throw new Error('「設定」と「基本パーツ」シートが必要です。');
  }

  let sheet = ss.getSheetByName(PRESET.sheetName);
  const newlyCreated = !sheet;
  sheet = sheet || ss.insertSheet(PRESET.sheetName);
  ensureSize_(sheet, PRESET.maxChoices, 23);
  if (newlyCreated) {
    sheet.getRange('A1:W2').merge().setValue('パーツプリセット：上の表にパーツ、右側に大問の小問行を保存します。B列の保存名を設定シートのH列で選びます。')
      .setBackground('#17375e').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
    sheet.getRange('A3:J3').setValues([[
      'ID', '保存名', '分類', 'パーツ種類', '表示文字・観点', '配点', '幅（列）', '高さ（行）', '後の空き（行）', '大問の題名'
    ]]);
    sheet.getRange('L3:W3').setValues([[
      'プリセットID', '行番号', '小問グループ', '小問ラベル', '幅比', '高さ（行）', 'メモ', '解答欄パターン', 'グループ経路', '枠数', '1枠の記入項目数', '判定根拠'
    ]]);
    sheet.getRange('A3:J3').setBackground('#d9eaf7').setFontWeight('bold');
    sheet.getRange('L3:W3').setBackground('#d9eaf7').setFontWeight('bold');
    sheet.setFrozenRows(3);
    sheet.setColumnWidths(1, 23, 110);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 190);
    sheet.setColumnWidth(5, 240);
    sheet.setColumnWidth(12, 300);
    sheet.setColumnWidth(15, 220);
    sheet.getRange('A4:B1000').setNumberFormat('@');
    sheet.getRange('L4:P1000').setNumberFormat('@');
  } else if (sheet.getRange('A3').getDisplayValue() !== 'ID' ||
             sheet.getRange('L3').getDisplayValue() !== 'プリセットID') {
    throw new Error('「パーツプリセット」という名前のシートが別の形式です。内容を確認してください。');
  }
  const patternHeading = String(sheet.getRange('S3').getDisplayValue()).trim();
  if (patternHeading && patternHeading !== '解答欄パターン') throw new Error('プリセットのS3に別の見出しがあります。');
  if (!patternHeading) sheet.getRange('S3').setValue('解答欄パターン').setBackground('#d9eaf7').setFontWeight('bold');
  ['グループ経路', '枠数', '1枠の記入項目数', '判定根拠'].forEach(function (label, i) {
    const cell = sheet.getRange(3, 20 + i);
    const previous = String(cell.getDisplayValue()).trim();
    if (previous && previous !== label) throw new Error('プリセットのT～W列に別の見出しがあります。');
    if (!previous) cell.setValue(label).setBackground('#d9eaf7').setFontWeight('bold');
  });

  ensureDetailPatternColumn_(settings);
  ensureSize_(basic, EXAM.basicStartRow + EXAM.basicRowCount - 1, 8);
  settings.getRange('H20').setValue('プリセット').setBackground('#d9eaf7')
    .setFontWeight('bold').setNote('保存済みの名前を選択後、メニュー「選択したプリセットを適用」を実行します。');
  basic.getRange('H3').setValue('プリセット').setBackground('#d9eaf7')
    .setFontWeight('bold').setNote('保存済みの名前を選択後、メニュー「選択したプリセットを適用」を実行します。');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('B4:B1000'), true)
    .setAllowInvalid(false)
    .setHelpText('パーツプリセットの保存名を選んでください。')
    .build();
  settings.getRange(EXAM.questionStartRow, 8, EXAM.questionRowCount, 1)
    .setBackground('#e9f4ea').setDataValidation(rule);
  basic.getRange(EXAM.basicStartRow, 8, EXAM.basicRowCount, 1)
    .setBackground('#e9f4ea').setDataValidation(rule);
  settings.setColumnWidth(8, 190);
  basic.setColumnWidth(8, 190);
  return sheet;
}

function withPresetLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新中です。少し待ってから再実行してください。');
  try { return callback(); } finally { lock.releaseLock(); }
}

/** パネルで選んだ現在のパーツと、大問なら小問行を保存する。 */
function savePartPreset(id, name) {
  return withPresetLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const presets = ensurePartPresetUi_(ss);
    name = String(name || '').trim();
    if (!name || name.length > 80 || /^[=+@-]/.test(name)) {
      throw new Error('保存名は1～80文字で入力してください。先頭に =、+、-、@ は使えません。');
    }
    const parentRows = presetRows_(presets, 1, 10);
    if (parentRows.some(function (r) { return String(r.values[1]).trim() === name; })) {
      throw new Error('保存名「' + name + '」は使用済みです。別の名前にしてください。');
    }

    const part = findEditablePart_(id);
    if (!part) throw new Error('保存するパーツが見つかりません。');
    const source = part.kind === 'basic' ? ss.getSheetByName(EXAM.basicSheetName) :
      ss.getSheetByName(EXAM.settingsSheetName);
    const values = source.getRange(part.row, 1, 1, 7).getValues()[0];
    if (values[0] !== true) throw new Error('保存するパーツの「使用」にチェックを付けてください。');
    const presetId = Utilities.getUuid();
    let parent, children = [];
    if (part.kind === 'basic') {
      const type = String(values[2] || '').trim();
      if (!['考査見出し', '年組氏名欄', '得点枠'].includes(type)) throw new Error('パーツ種類を選んでください。');
      parent = [presetId, name, '基本', type, String(values[3] || ''), '',
        optionalInteger_(values[4], 33, '幅', type === '年組氏名欄' ? 8 : 3, 33),
        optionalInteger_(values[5], type === '考査見出し' ? 2 : 4, '高さ', 2, 20),
        optionalInteger_(values[6], 1, '後の空き', 0, 10), ''];
    } else {
      const number = String(values[2] || '').trim();
      if (!number) throw new Error('大問番号を入力してください。');
      const width = optionalInteger_(values[5], 33, '幅', 6, 33);
      const details = source.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount).getDisplayValues()
        .filter(function (r) { return String(r[0] || '').trim() === number; });
      if (!details.length) throw new Error('大問「' + number + '」の小問行を1行以上入力してください。');
      const indexes = new Set();
      children = details.map(function (r) {
        const index = requiredInteger_(r[1], '小問行の行番号', 1, 999);
        if (indexes.has(index)) throw new Error('小問行の行番号 ' + index + ' が重複しています。');
        indexes.add(index);
        const labels = String(r[3] || '').trim();
        const weights = String(r[4] || '').trim();
        const labelList = labels ? labels.split(',').map(function (x) { return x.trim(); }) : [''];
        const ratios = weights ? weights.split(',').map(Number) : labelList.map(function () { return 1; });
        if (labelList.some(function (x) { return labels && !x; }) || ratios.length !== labelList.length ||
            ratios.some(function (x) { return !Number.isFinite(x) || x <= 0; })) {
          throw new Error('小問行 ' + index + ' のラベルと幅比を確認してください。');
        }
        const height = optionalInteger_(r[5], 2, '小問行の高さ', 1, 10);
        const pattern = answerPattern_(r[7], labels, index);
        if (pattern === '小問複合枠') validateAnswerWidths_({ width: width }, { labels: labelList, weights: ratios, sheetRow: index });
        const frames = optionalInteger_(r[9], pattern === '小問複合枠' ? labelList.length : 1, '枠数', 1, 30);
        if (pattern === '小問複合枠' && frames !== labelList.length) throw new Error('ラベル数と枠数が違います。');
        return [presetId, index, String(r[2] || ''), labels, weights || '1', height, String(r[6] || ''), pattern,
          String(r[8] || ''), frames, optionalInteger_(r[10], 1, '記入項目数', 1, 10), String(r[11] || '')];
      }).sort(function (a, b) { return a[1] - b[1]; });
      parent = [presetId, name, '大問', '大問', String(values[3] || ''),
        values[4] === '' ? '' : String(values[4]), width, '',
        optionalInteger_(values[6], 1, '後の空き', 0, 10),
        String(source.getRange(part.row, 9).getDisplayValue() || '')];
    }

    const pRow = presetEmptyRow_(presets, 1);
    const childStart = children.length ? presetEmptyBlock_(presets, 12, children.length) : 0;
    try {
      presets.getRange(pRow, 1, 1, 10).setValues([parent]);
      if (children.length) {
        presets.getRange(childStart, 12, children.length, EXAM.detailColumnCount).setValues(children);
      }
    } catch (e) {
      presets.getRange(pRow, 1, 1, 10).clearContent();
      if (children.length) presets.getRange(childStart, 12, children.length, EXAM.detailColumnCount).clearContent();
      throw e;
    }
    return '「' + name + '」をパーツプリセットに保存しました。';
  });
}

function presetRows_(sheet, column, width) {
  const count = Math.max(0, sheet.getLastRow() - PRESET.firstRow + 1);
  return count ? sheet.getRange(PRESET.firstRow, column, count, width).getValues()
    .map(function (values, i) { return { row: PRESET.firstRow + i, values: values }; }) : [];
}

function presetEmptyRow_(sheet, column) {
  const rows = presetRows_(sheet, column, 10);
  const empty = rows.find(function (r) {
    return r.values.every(function (value) { return value === ''; });
  });
  const row = empty ? empty.row : PRESET.firstRow + rows.length;
  ensureSize_(sheet, row, 23);
  return row;
}

function presetEmptyBlock_(sheet, column, count) {
  const rows = presetRows_(sheet, column, 1);
  const row = PRESET.firstRow + rows.length;
  ensureSize_(sheet, row + count - 1, 23);
  return row;
}

/** H列で選択したプリセットを、現在選択中の1行へ適用する。 */
function applySelectedPartPreset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensurePartPresetUi_(ss);
  const active = ss.getActiveSheet();
  const sheetName = active.getName();
  const row = active.getActiveCell().getRow();
  const kind = sheetName === EXAM.settingsSheetName ? '大問' :
    sheetName === EXAM.basicSheetName ? '基本' : '';
  if (!kind || row < (kind === '大問' ? EXAM.questionStartRow : EXAM.basicStartRow) ||
      row >= (kind === '大問' ? EXAM.questionStartRow + EXAM.questionRowCount :
        EXAM.basicStartRow + EXAM.basicRowCount)) {
    throw new Error('「設定」の大問行、または「基本パーツ」の入力行を選択してください。');
  }
  const name = String(active.getRange(row, 8).getDisplayValue()).trim();
  if (!name) throw new Error('選択した行のH列でプリセットを選んでください。');
  const before = presetTargetSnapshot_(active, row, kind);
  if (kind === '大問' && before.details.length > 0) {
    const response = SpreadsheetApp.getUi().alert(
      '小問行の置き換え',
      '大問「' + before.number + '」の既存の小問' + before.details.length +
        '行を「' + name + '」の内容に置き換えますか？',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (response !== SpreadsheetApp.getUi().Button.YES) return;
  }
  const result = withPresetLock_(function () {
    const current = presetTargetSnapshot_(active, row, kind);
    if (JSON.stringify(current) !== JSON.stringify(before) ||
        String(active.getRange(row, 8).getDisplayValue()).trim() !== name) {
      throw new Error('確認中に設定が変更されました。選択し直してください。');
    }
    return applyPartPresetLocked_(ss, active, row, kind, name, current);
  });
  generateAnswerSheet();
  return result;
}

function presetTargetSnapshot_(sheet, row, kind) {
  const values = sheet.getRange(row, 1, 1, 7).getValues()[0];
  const number = kind === '大問' ? String(values[2] || '').trim() : '';
  const details = kind === '大問' ? sheet.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount)
    .getValues().map(function (r, i) { return { index: i, values: r }; })
    .filter(function (x) { return String(x.values[0] || '').trim() === number && number !== ''; }) : [];
  return { values: values, number: number, details: details,
    title: kind === '大問' ? sheet.getRange(row,9).getValue() : '' };
}

function applyPartPresetLocked_(ss, target, row, kind, name, before) {
  const presets = ss.getSheetByName(PRESET.sheetName);
  const matches = presetRows_(presets, 1, 10).filter(function (r) {
    return String(r.values[1] || '').trim() === name;
  });
  if (matches.length !== 1) throw new Error('プリセット「' + name + '」が見つからないか、保存名が重複しています。');
  const p = matches[0].values;
  if (p[2] !== kind) throw new Error('「' + name + '」は' + p[2] + '用です。' + kind + 'の行には適用できません。');
  const order = requiredInteger_(before.values[1], '適用先の順番', 1, 999);
  let newValues, newDetails = [], changed = [];
  if (kind === '基本') {
    const type = String(p[3] || '').trim();
    if (!['考査見出し', '年組氏名欄', '得点枠'].includes(type)) throw new Error('プリセットの種類が正しくありません。');
    newValues = [true, order, type, String(p[4] || ''),
      requiredInteger_(p[6], 'プリセットの幅', type === '年組氏名欄' ? 8 : 3, 33),
      requiredInteger_(p[7], 'プリセットの高さ', 2, 20),
      requiredInteger_(p[8], 'プリセットの後の空き', 0, 10)];
  } else {
    if (!before.number) throw new Error('適用先の大問番号を先に入力してください。');
    const currentQuestions = target.getRange(EXAM.questionStartRow, 3, EXAM.questionRowCount, 1)
      .getDisplayValues().filter(function (r) { return String(r[0]).trim() === before.number; });
    if (currentQuestions.length !== 1) throw new Error('適用先の大問番号が重複しています。');
    const width = requiredInteger_(p[6], 'プリセットの幅', 6, 33);
    newValues = [true, order, before.number, String(p[4] || ''),
      p[5] === '' ? '' : String(p[5]), width,
      requiredInteger_(p[8], 'プリセットの後の空き', 0, 10)];
    const childRows = presetRows_(presets, 12, EXAM.detailColumnCount)
      .filter(function (r) { return String(r.values[0]) === String(p[0]); })
      .sort(function (a, b) { return Number(a.values[1]) - Number(b.values[1]); });
    if (!childRows.length) throw new Error('大問プリセットに小問行がありません。');
    const used = new Set();
    newDetails = childRows.map(function (r) {
      const d = r.values;
      const index = requiredInteger_(d[1], 'プリセット小問行の行番号', 1, 999);
      if (used.has(index)) throw new Error('プリセットの小問行番号が重複しています。');
      used.add(index);
      const labels = String(d[3] || '').trim();
      const weights = String(d[4] || '').trim();
      const labelsArray = labels ? labels.split(',').map(function (x) { return x.trim(); }) : [''];
      const ratios = weights ? weights.split(',').map(Number) : labelsArray.map(function () { return 1; });
      if (labelsArray.some(function (x) { return labels && !x; }) || ratios.length !== labelsArray.length ||
          ratios.some(function (x) { return !Number.isFinite(x) || x <= 0; })) {
        throw new Error('プリセットの小問行 ' + index + ' のラベルまたは幅比を確認してください。');
      }
      const pattern = answerPattern_(d[7], labels, r.row);
      if (pattern === '小問複合枠') validateAnswerWidths_({ width: width }, { labels: labelsArray, weights: ratios, sheetRow: r.row });
      const frames = optionalInteger_(d[9], pattern === '小問複合枠' ? labelsArray.length : 1, '枠数', 1, 30);
      if (pattern === '小問複合枠' && frames !== labelsArray.length) throw new Error('プリセットの枠数とラベル数が違います。');
      return [before.number, index, String(d[2] || ''), labels, weights || '1',
        requiredInteger_(d[5], 'プリセット小問行の高さ', 1, 10), String(d[6] || ''), pattern,
        String(d[8] || ''), frames, optionalInteger_(d[10], 1, '記入項目数', 1, 10), String(d[11] || '')];
    });
    const allDetails = target.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount).getValues();
    const empty = allDetails.map(function (r, i) { return { index: i, values: r }; })
      .filter(function (x) { return x.values.every(function (v) { return v === '' || v === false; }); });
    const slots = before.details.concat(empty);
    if (slots.length < newDetails.length) throw new Error('小問行の空きが足りません。');
    const targetIndexes = slots.slice(0, newDetails.length).map(function (x) { return x.index; });
    changed = Array.from(new Set(before.details.map(function (x) { return x.index; }).concat(targetIndexes)))
      .map(function (index) { return { index: index, values: allDetails[index].slice() }; });
    newDetails = newDetails.map(function (values, i) { return { index: targetIndexes[i], values: values }; });
  }

  try {
    target.getRange(row, 1, 1, 7).setValues([newValues]);
    if (kind === '大問') target.getRange(row,9).setValue(String(p[9] || ''));
    if (kind === '大問') {
      before.details.forEach(function (x) {
        target.getRange(EXAM.detailStartRow + x.index, 10, 1, EXAM.detailColumnCount).clearContent();
      });
      newDetails.forEach(function (x) {
        target.getRange(EXAM.detailStartRow + x.index, 10, 1, EXAM.detailColumnCount)
          .setNumberFormat('@').setValues([x.values]);
      });
    }
    SpreadsheetApp.flush();
    calculatePlacements_(readParts_(ss.getSheetByName(EXAM.settingsSheetName)));
  } catch (e) {
    target.getRange(row, 1, 1, 7).setValues([before.values]);
    if (kind === '大問') target.getRange(row,9).setValue(before.title);
    changed.forEach(function (x) {
      target.getRange(EXAM.detailStartRow + x.index, 10, 1, EXAM.detailColumnCount).setValues([x.values]);
    });
    throw new Error('プリセットを適用できませんでした。設定を戻しました。詳細: ' + (e.message || e));
  }
  return '「' + name + '」を' + (kind === '大問' ? '大問 ' + before.number : '基本パーツ') + 'に適用しました。';
}