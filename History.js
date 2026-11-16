/** 「設定_旧_...」から大問・小問・基本パーツの履歴を復元する。 */
function listSettingsHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(function (sheet) { return /^設定_旧_/.test(sheet.getName()); })
    .map(function (sheet) {
      const name = sheet.getName();
      const hasQuestions = sheet.getRange('A20').getDisplayValue() === '使用' &&
        sheet.getRange('J5').getDisplayValue() === '対象の大問番号';
      const embeddedBasics = sheet.getRange('A5').getDisplayValue() === '使用';
      const paired = ss.getSheetByName('基本パーツ_旧_' + name.slice('設定_旧_'.length));
      return {
        name: name,
        supported: hasQuestions,
        basicSource: paired ? '同時保存された基本パーツ' :
          embeddedBasics ? '旧設定シート内の基本パーツ' : '基本パーツは現在の値を維持'
      };
    }).sort(function (a, b) { return b.name.localeCompare(a.name); });
}

function previewSettingsHistory(name) {
  const source = historySource_(name);
  const data = historyData_(source);
  return {
    name: name,
    basicSource: data.basicSource,
    basics: data.basics ? data.basics.filter(function (r) { return r[0] === true; })
      .map(function (r) { return String(r[2] || ''); }) : null,
    questions: data.questions.filter(function (r) { return r[0] === true; })
      .map(function (r) { return String(r[2] || ''); }),
    detailCount: data.details.filter(function (r) { return String(r[0] || '').trim(); }).length
  };
}

function restoreSettingsHistory(name) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新中です。少し待ってください。');
  try {
    return restoreSettingsHistoryLocked_(name);
  } finally {
    lock.releaseLock();
  }
}

function restoreSettingsHistoryLocked_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = historySource_(name);
  const data = historyData_(source);
  const settings = ss.getSheetByName(EXAM.settingsSheetName);
  const basic = ss.getSheetByName(EXAM.basicSheetName);
  if (!settings || !basic) throw new Error('先に新形式の設定シートを作成してください。');
  if (settings.getRange('A20').getDisplayValue() !== '使用' ||
      basic.getRange('A3').getDisplayValue() !== '使用') {
    throw new Error('現在の設定表が新形式ではありません。');
  }
  if (typeof ensurePartPresetUi_ === 'function') ensurePartPresetUi_(ss);

  // 復元直前の状態も履歴として保存する。
  const backupName = makeBackupName_(ss);
  settings.copyTo(ss).setName(backupName);
  basic.copyTo(ss).setName('基本パーツ_旧_' + backupName.slice('設定_旧_'.length));

  ensureDetailPatternColumn_(settings);
  const qRange = settings.getRange(EXAM.questionStartRow, 1, EXAM.questionRowCount, 9);
  const dRange = settings.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount);
  const bRange = basic.getRange(EXAM.basicStartRow, 1, EXAM.basicRowCount, 8);
  const before = { questions: qRange.getValues(), details: dRange.getValues(), basics: bRange.getValues() };
  try {
    qRange.setValues(data.questions);
    dRange.setValues(data.details);
    if (data.basics) bRange.setValues(data.basics);
    SpreadsheetApp.flush();
    calculatePlacements_(readParts_(settings));
  } catch (e) {
    qRange.setValues(before.questions);
    dRange.setValues(before.details);
    bRange.setValues(before.basics);
    throw new Error('復元できませんでした。現在の設定は維持しました。詳細: ' + (e.message || e));
  }
  generateAnswerSheet();
  return name + ' を再利用しました。復元前の設定: ' + backupName;
}

function historySource_(name) {
  if (typeof name !== 'string' || !/^設定_旧_[0-9]{4}_[0-9]{6}(?:_[0-9]+)?$/.test(name)) {
    throw new Error('履歴の名前が正しくありません。');
  }
  const source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!source) throw new Error('選択した履歴が見つかりません。');
  return source;
}

function historyData_(source) {
  const ss = source.getParent();
  if (source.getRange('A20').getDisplayValue() !== '使用' ||
      source.getRange('J5').getDisplayValue() !== '対象の大問番号') {
    throw new Error('この履歴は現在の大問・小問表と形式が異なります。');
  }
  const questions = historyRows_(source, EXAM.questionStartRow, 1, EXAM.questionRowCount, 9);
  const details = historyRows_(source, EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount);
  const paired = ss.getSheetByName('基本パーツ_旧_' + source.getName().slice('設定_旧_'.length));
  let basics = null;
  let basicSource = '現在の基本パーツを維持';
  if (paired && paired.getRange('A3').getDisplayValue() === '使用') {
    basics = historyRows_(paired, EXAM.basicStartRow, 1, EXAM.basicRowCount, 8);
    basicSource = '同時保存された基本パーツ';
  } else if (source.getRange('A5').getDisplayValue() === '使用') {
    basics = historyRows_(source, 6, 1, EXAM.basicRowCount, 7)
      .map(function (r) { return r.concat(''); });
    basicSource = '旧設定シート内の基本パーツ';
  }
  if (basics) basics = basics.map(function (r) {
    return [r[0] === true, r[1], r[2], r[3], r[4], r[5], r[6], r[7]];
  });
  return { questions: questions, details: details, basics: basics, basicSource: basicSource };
}

function historyRows_(sheet, startRow, startColumn, count, width) {
  const available = Math.max(0, Math.min(count, sheet.getMaxRows() - startRow + 1));
  const readable = Math.max(0, Math.min(width, sheet.getMaxColumns() - startColumn + 1));
  const rows = available && readable ? sheet.getRange(startRow, startColumn, available, readable)
    .getValues().map(function (r) {
      while (r.length < width) r.push('');
      return r;
    }) : [];
  while (rows.length < count) rows.push(Array(width).fill(''));
  return rows;
}