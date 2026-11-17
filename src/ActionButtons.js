/**
 * ひな形に操作ボタンを配置する。既存の同名ボタンは再利用する。
 * 初回は clasp push 後、Apps Script エディタから一度実行する。
 */
function installActionButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureActionButtons_(ss);
  SpreadsheetApp.getUi().alert('操作ボタンを配置しました。シートを開き直してご利用ください。');
}

function ensureActionButtons_(ss) {
  const definitions = [
    { sheet: EXAM.settingsSheetName, row: 7, column: 1, key: 'update_output',
      label: '出力を更新', action: 'confirmGenerateAnswerSheet' },
    { sheet: EXAM.settingsSheetName, row: 7, column: 4, key: 'setup_presets',
      label: 'プリセット欄を追加・更新', action: 'confirmSetupPartPresets' },
    { sheet: EXAM.settingsSheetName, row: 10, column: 1, key: 'apply_preset',
      label: '選択したプリセットを適用', action: 'confirmApplySelectedPartPreset' },
    { sheet: EXAM.basicSheetName, row: 15, column: 1, key: 'apply_preset',
      label: '選択したプリセットを適用', action: 'confirmApplySelectedPartPreset' },
    { sheet: EXAM.outputSheetName, row: 2, column: 73, key: 'edit_parts',
      label: 'パーツ編集パネルを開く', action: 'confirmShowPartEditor' },
  ];
  const missingSheets = [...new Set(definitions.map(function (item) { return item.sheet; }))]
    .filter(function (name) { return !ss.getSheetByName(name); });
  if (missingSheets.length) {
    throw new Error('操作ボタンを配置できません。「' + missingSheets.join('」「') + '」シートがありません。');
  }

  definitions.forEach(function (item) {
    const sheet = ss.getSheetByName(item.sheet);
    ensureSize_(sheet, item.row, item.column + 13);
    const anchor = sheet.getRange(item.row, item.column);
    const marker = '解答用紙操作:' + item.key;
    let button = sheet.getImages().find(function (image) {
      return image.getAltTextTitle() === marker;
    });
    if (!button) {
      const bytes = Utilities.base64Decode(ACTION_BUTTON_IMAGES[item.key]);
      button = sheet.insertImage(Utilities.newBlob(bytes, 'image/png', item.key + '.png'),
        anchor.getColumn(), anchor.getRow());
    }
    button.setAnchorCell(anchor)
      .setAnchorCellXOffset(0)
      .setAnchorCellYOffset(0)
      .setWidth(260)
      .setHeight(44)
      .setAltTextTitle(marker)
      .setAltTextDescription(item.label)
      .assignScript(item.action);
  });
}

function runConfirmedAction_(label, action) {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(label, '実行しますか？', ui.ButtonSet.YES_NO);
  if (answer === ui.Button.YES) return action();
}

function confirmGenerateAnswerSheet() {
  return runConfirmedAction_('出力を更新', generateAnswerSheet);
}

function confirmSetupPartPresets() {
  return runConfirmedAction_('プリセット欄を追加・更新', setupPartPresets);
}

function confirmApplySelectedPartPreset() {
  return runConfirmedAction_('選択したプリセットを適用', applySelectedPartPreset);
}

function confirmShowPartEditor() {
  return runConfirmedAction_('パーツ編集パネルを開く', showPartEditor);
}
