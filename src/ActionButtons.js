/** 出力を中心とした操作メニュー。 */
function onOpen() {
  return withOperationLog_('open', function () {
    return onOpenLogged_();
  });
}

function onOpenLogged_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('パーツ編集')
    .addItem('パネルを開く', 'showPartEditor')
    .addToUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeLegacySheetButtons_(ss);
}

/** 旧画像ボタンのみを除き、シート上の他の画像には触れない。 */
function removeLegacySheetButtons_(ss) {
  [EXAM.settingsSheetName, EXAM.basicSheetName, EXAM.outputSheetName].forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.getImages().filter(function (image) {
      return String(image.getAltTextTitle() || '').indexOf('解答用紙操作:') === 0;
    }).forEach(function (image) { image.remove(); });
    sheet.getDrawings().filter(function (drawing) {
      return ['generateAnswerSheet', 'confirmSaveCurrentSettings',
        'confirmSetupPartPresets', 'confirmApplySelectedPartPreset'].includes(drawing.getOnAction());
    }).forEach(function (drawing) { drawing.remove(); });
  });
}

/** 旧ボタン配置関数を呼んでも、新しい画像は配置しない。 */
function installActionButtons() {
  removeLegacySheetButtons_(SpreadsheetApp.getActiveSpreadsheet());
}
