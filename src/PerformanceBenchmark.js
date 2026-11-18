/** 読み取りと空き領域の集計時間を計測する。出力や作業データは変更しない。 */
function benchmarkPlacementRead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXAM.outputSheetName);
  if (!sheet) throw new Error('「出力」シートがありません。');
  if (!draftSheet_(ss, false)) throw new Error('「作業データ」がありません。パーツ編集を一度開いてから実行してください。');
  const samples = [];
  for (let i = 0; i < 5; i++) {
    const started = Date.now();
    readDraft_();
    outputSnapshot_(outputReadGrid_(sheet));
    samples.push(Date.now() - started);
  }
  const sorted = samples.slice().sort(function (a, b) { return a - b; });
  const report = { operation:'benchmark_placement_read', samplesMs:samples,
    medianMs:sorted[2], maxMs:sorted[4] };
  console.info(JSON.stringify(report));
  SpreadsheetApp.getUi().alert('読み取りと空き判定の計測（5回）\n' +
    '各回：' + samples.join('、') + 'ms\n中央値：' + sorted[2] + 'ms\n' +
    '※配置や画面への反映は含みません。配置処理の応答時間は配置後に画面へ表示されます。');
  return report;
}
