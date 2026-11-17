/**
 * 解答用紙ジェネレーター
 *
 * 「設定」シートの内容から、B4横・左右2領域の「出力」シートを生成します。
 *
 * 配置ルール:
 * - 使用にチェックを付けた基本パーツと大問を「順番」で並べる
 * - 左側の上から配置する
 * - パーツ全体が左側に収まらなければ、次のパーツから右側へ配置する
 * - 一つの大問とその小問行は途中で分割しない
 *
 * ひな形の初期作成時だけ、保守担当者が setupSheets() で入力表を用意する。
 * 試験ごとにスプレッドシート全体をコピーして使用する。
 * 既存の設定シートがある場合、setupSheets() は内容を変更しない。
 */

const EXAM = {
  settingsSheetName: "設定",
  basicSheetName: "基本パーツ",
  outputSheetName: "出力",

  // B4横を想定した出力領域。印刷時の物理寸法は印刷設定にも依存します。
  outputColumns: 72,
  outputRows: 51,
  cellSizePx: 19,

  leftColumn: 3,
  rightColumn: 39,
  areaWidth: 33,
  firstRow: 3,

  basicStartRow: 4,
  basicRowCount: 10,

  questionStartRow: 21,
  questionRowCount: 40,

  detailStartRow: 6,
  detailRowCount: 150,
  detailColumnCount: 12, // J:U。Q パターン、R 経路、S 枠数、T 記入項目数、U 根拠
};

/**
 * 入力しやすい設定シートを初回だけ作る。
 * 既存の設定・基本パーツ・出力シートがあれば何も変更しない。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existing = [EXAM.settingsSheetName, EXAM.basicSheetName, EXAM.outputSheetName]
    .filter(function (name) { return ss.getSheetByName(name); });
  if (existing.length) {
    SpreadsheetApp.getUi().alert(
      "既存の「" + existing.join("」「") + "」シートがあります。設定は変更していません。\n" +
      "新しい試験を作る場合は、スプレッドシート全体をコピーしてください。"
    );
    return;
  }

  const settings = ss.insertSheet(EXAM.settingsSheetName);
  const basicSheet = ss.insertSheet(EXAM.basicSheetName);
  const output = ss.insertSheet(EXAM.outputSheetName);

  ensureSize_(settings, 155, 21);
  ensureSize_(basicSheet, 13, 8);
  ensureSize_(output, EXAM.outputRows, EXAM.outputColumns);

  resetSheet_(settings);
  resetSheet_(basicSheet);
  resetSheet_(output);

  configureBasicSheet_(basicSheet, []);

  // 全体説明
  settings.getRange("A1:P2")
    .merge()
    .setValue(
      "B4横・解答用紙の設定\n" +
      "基本パーツ・大問・小問を入力 → 使用にチェック → このシートの「出力を更新」ボタン"
    )
    .setBackground("#17375e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(13)
    .setWrap(true)
    .setVerticalAlignment("middle");

  settings.setRowHeights(1, 2, 28);

  settings.getRange("A3:H3")
    .merge()
    .setValue(
      "配置：順番どおりに左上から並びます。左に収まらないパーツは、右上から続けて並びます。"
    )
    .setBackground("#eaf2f8")
    .setWrap(true);

  settings.getRange("J3:P3")
    .merge()
    .setValue(
      "例：小問ラベル「①,②,③」／幅比「1,1,2」→ 3つ目の欄が広くなります。"
    )
    .setBackground("#eaf2f8")
    .setWrap(true);

  settings.getRange("A4:G5").merge()
    .setValue("① 考査見出し・年組氏名欄・得点枠は「基本パーツ」シートで編集します。")
    .setBackground("#d9eaf7").setFontWeight("bold").setWrap(true);

  // 大問
  settings.getRange("A18:G18")
    .merge()
    .setValue("② 大問：使用する大問を登録します。大問番号は③小問行との対応に使います。")
    .setBackground("#d9eaf7")
    .setFontWeight("bold");

  settings.getRange("A19:G19")
    .merge()
    .setValue(
      "例：使用 ✓ ／順番 2／大問番号 1／観点 知・技／配点 10／幅 33／後の空き 1"
    )
    .setBackground("#f4f8fc")
    .setWrap(true);

  settings.getRange("I20").setValue("大問の題名").setBackground("#d9eaf7").setFontWeight("bold");
  settings.getRange("A20:G20").setValues([[
    "使用",
    "順番",
    "大問番号",
    "観点",
    "配点",
    "幅（列）",
    "後の空き（行）",
  ]]);

  // 小問行
  settings.getRange("J4:U4")
    .merge()
    .setValue(
      "③ 小問行：大問番号で②の大問と対応します。大問の中では行番号順に並びます。"
    )
    .setBackground("#d9eaf7")
    .setFontWeight("bold");

  settings.getRange("J5:U5").setValues([[
    "対象の大問番号",
    "行番号",
    "小問グループの番号",
    "小問ラベル",
    "幅比",
    "高さ（行）",
    "メモ",
    "解答欄パターン",
    "グループ経路",
    "枠数",
    "1枠の記入項目数",
    "判定根拠",
  ]]);

  // 空欄から入力を開始できるよう、実際のデータ行には見本を入れない。
  settings.getRange("A21:A60").insertCheckboxes();
  // 大問番号や (1) のような文字を、数式や日付に変換させない。
  settings.getRange("C21:C60").setNumberFormat("@");
  settings.getRange("J6:M155").setNumberFormat("@");
  settings.getRange("A20:G20").setNotes([[
    "出力する大問だけチェック",
    "基本パーツと共通の順番。例：2",
    "③小問行の「対象の大問番号」と同じ文字を入力",
    "見出しに表示。空欄でも可",
    "見出しに表示。空欄なら点数表示なし",
    "6～33。空欄は33",
    "0～10。空欄は1",
  ]]);

  settings.getRange("J5:U5").setNotes([[
    "②大問の「大問番号」と一致させる",
    "大問内の並び順。例：1、2、3",
    "行の左端に表示。例：1、(1)",
    "カンマ区切り。例：①,②,③。空欄なら一つの空欄",
    "ラベルと同じ個数。例：1,1,2。ラベル空欄なら空欄または1",
    "1～10。空欄なら2",
    "出力には反映されません",
    "小問複合枠／記述欄。空欄ならラベルの有無から自動判定",
    "例：2-(2)①。空欄なら大問番号と行番号から補います",
    "物理的な解答枠数。空欄ならラベル数または1",
    "1枠内に記入する項目数。空欄なら1",
    "解析で枠数を決めた理由。自動出力には反映されません",
  ]]);

  // 設定表の見た目
  settings.getRange("A21:G60").setBackground("#fff9e6");
  settings.getRange("J6:U155").setBackground("#fff9e6");

  ["A20:G20", "J5:U5"].forEach(function (address) {
    settings.getRange(address)
      .setBackground("#d9eaf7")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setWrap(true);
  });

  settings.getRange("A20:G60")
    .setBorder(true, true, true, true, true, true);

  settings.getRange("J5:U25")
    .setBorder(true, true, true, true, true, true);

  settings.setColumnWidth(1, 55);
  settings.setColumnWidth(2, 70);
  settings.setColumnWidth(3, 125);
  settings.setColumnWidth(4, 320);
  settings.setColumnWidths(5, 3, 110);
  settings.setColumnWidth(8, 20);
  settings.setColumnWidth(9, 220);
  settings.setColumnWidth(10, 125);
  settings.setColumnWidth(11, 75);
  settings.setColumnWidth(12, 100);
  settings.setColumnWidth(13, 240);
  settings.setColumnWidth(14, 150);
  settings.setColumnWidth(15, 105);
  settings.setColumnWidth(16, 150);
  settings.setColumnWidth(17, 140);
  settings.setColumnWidth(18, 170);
  settings.setColumnWidths(19, 2, 105);
  settings.setColumnWidth(21, 260);
  ensureDetailPatternColumn_(settings);

  settings.setRowHeight(5, 42);
  settings.setRowHeight(20, 42);
  settings.setFrozenRows(5);

  // プリセット管理表は設定の初期化対象に含めない。
  ensurePartPresetUi_(ss);
  configureOutputSheet_(output);
  ensureActionButtons_(ss);

  SpreadsheetApp.getUi().alert(
    "空の設定シートを作成しました。\n" +
    "次の試験では、このスプレッドシート全体をコピーして使用してください。"
  );
}

/** 考査見出し・氏名欄・得点枠の入力専用シート。 */
function configureBasicSheet_(sheet, previousRows) {
  sheet.getRange("A1:G2").merge()
    .setValue("基本パーツ設定\n使用する行にチェック。順番は「設定」シートの大問と共通です。")
    .setBackground("#17375e").setFontColor("#ffffff")
    .setFontWeight("bold").setWrap(true);
  sheet.getRange("A3:G3").setValues([[
    "使用", "順番", "パーツ種類", "表示文字", "幅（列）", "高さ（行）", "後の空き（行）"
  ]]).setBackground("#d9eaf7").setFontWeight("bold").setWrap(true);
  sheet.getRange(EXAM.basicStartRow, 1, EXAM.basicRowCount, 7)
    .setBackground("#fff9e6");
  sheet.getRange("A4:A13").insertCheckboxes();
  sheet.getRange("C4:C13").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["考査見出し", "年組氏名欄", "得点枠"], true)
      .setAllowInvalid(false).build()
  );
  if (previousRows.length) {
    // checkbox の false を維持しつつ、旧設定の入力を引き継ぐ。
    sheet.getRange(EXAM.basicStartRow, 1, Math.min(previousRows.length, EXAM.basicRowCount), 7)
      .setValues(previousRows.slice(0, EXAM.basicRowCount).map(function (r) {
        return [r[0] === true, r[1], r[2], r[3], r[4], r[5], r[6]];
      }));
  }
  sheet.getRange("A3:G3").setNotes([[
    "出力する行だけチェック", "大問と共通の順番", "プルダウンから選択",
    "考査見出しに表示する文字", "3～33。年組氏名欄は8～33。空欄は33",
    "2～20。空欄は既定値", "0～10。空欄は1"
  ]]);
  sheet.setColumnWidths(1, 2, 75);
  sheet.setColumnWidth(3, 145);
  sheet.setColumnWidth(4, 360);
  sheet.setColumnWidths(5, 3, 115);
  sheet.setFrozenRows(3);
}

/**
 * 設定を読み、問題がないことを確認してから出力を更新する。
 */
function generateAnswerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(EXAM.settingsSheetName);

  if (!settings) {
    SpreadsheetApp.getUi().alert(
      "「設定」シートがありません。先に setupSheets() を実行してください。"
    );
    return;
  }

  let placements;

  try {
    ensureDetailPatternColumn_(settings);
    const parts = readParts_(settings);
    placements = calculatePlacements_(parts);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "設定を確認してください。\n\n" + String(error.message || error)
    );
    return;
  }

  const output =
    ss.getSheetByName(EXAM.outputSheetName) ||
    ss.insertSheet(EXAM.outputSheetName);

  ensureSize_(output, EXAM.outputRows, EXAM.outputColumns);
  resetSheet_(output);
  configureOutputSheet_(output);

  placements.forEach(function (placement) {
    if (placement.part.kind === "basic") {
      drawBasic_(output, placement);
    } else {
      drawQuestion_(output, placement);
    }
  });

  SpreadsheetApp.getUi().alert(
    "設定から出力シートを更新しました。\n" +
      "印刷するときは、出力シートのA1:BT51を選択してB4・横向き・1ページにしてください。"
  );
}

/** 既存の設定を消さずに Q 列のパターン設定だけを追加する。 */
function ensureDetailPatternColumn_(sheet) {
  const titleCell = sheet.getRange('I20');
  const titleHeading = String(titleCell.getDisplayValue()).trim();
  if (titleHeading && titleHeading !== '大問の題名') throw new Error('設定シートのI20に別の見出しがあります。');
  if (!titleHeading) titleCell.setValue('大問の題名').setBackground('#d9eaf7').setFontWeight('bold');
  ensureSize_(sheet, EXAM.detailStartRow + EXAM.detailRowCount - 1, 21);
  const headings = ['解答欄パターン', 'グループ経路', '枠数', '1枠の記入項目数', '判定根拠'];
  headings.forEach((heading, i) => {
    const cell = sheet.getRange(5, 17 + i);
    const actual = String(cell.getDisplayValue()).trim();
    if (actual && actual !== heading) throw new Error(`設定シートの${'QRSTU'[i]}5に別の見出しがあります。`);
    if (!actual) {
      cell.setValue(heading).setBackground('#d9eaf7').setFontWeight('bold');
      sheet.getRange(6, 17 + i, EXAM.detailRowCount, 1).setBackground('#fff9e6');
    }
  });
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['小問複合枠', '記述欄'], true).setAllowInvalid(false).build();
  sheet.getRange('Q6:Q155').setDataValidation(rule);
}

function effectiveFrameCount_(values, labels, pattern, sheetRow) {
  const fallback = pattern === '小問複合枠' ? labels.length : 1;
  const count = optionalInteger_(values[9], fallback, `小問行 ${sheetRow} 行目の枠数`, 1, 30);
  if (pattern === '小問複合枠' && count !== labels.length) {
    throw new Error(`小問行 ${sheetRow} 行目：枠数と小問ラベル数を一致させてください。`);
  }
  return count;
}

function answerPattern_(raw, labelsText, sheetRow) {
  const value = String(raw || '').trim();
  if (!value) return String(labelsText || '').trim() ? '小問複合枠' : '記述欄';
  if (value !== '小問複合枠' && value !== '記述欄') {
    throw new Error(`小問行 ${sheetRow} 行目：解答欄パターンを選んでください。`);
  }
  return value;
}

/**
 * 設定シートから使用中のパーツを読み取る。
 */
function readParts_(sheet) {
  ensureDetailPatternColumn_(sheet);
  const basicSheet = sheet.getParent().getSheetByName(EXAM.basicSheetName);
  if (
    sheet.getRange("A20").getDisplayValue() !== "使用" ||
    sheet.getRange("J5").getDisplayValue() !== "対象の大問番号"
  ) {
    throw new Error(
      "設定シートが旧形式です。setupSheets() で新しい設定シートを作成してください。"
    );
  }
  if (!basicSheet || basicSheet.getRange("A3").getDisplayValue() !== "使用") {
    throw new Error("「基本パーツ」シートがありません。setupSheets() で作成してください。");
  }

  const parts = [];

  const basicRows = basicSheet.getRange(
    EXAM.basicStartRow,
    1,
    EXAM.basicRowCount,
    7
  ).getValues();

  basicRows.forEach(function (values, index) {
    if (values[0] !== true) return;

    const sheetRow = EXAM.basicStartRow + index;
    const type = String(values[2] || "").trim();

    if (!["考査見出し", "年組氏名欄", "得点枠"].includes(type)) {
      throw new Error(
        `基本パーツ ${sheetRow} 行目：パーツ種類を選択してください。`
      );
    }

    const defaultHeight =
      type === "考査見出し" ? 2 : 4;

    parts.push({
      kind: "basic",
      sheetRow: sheetRow,
      order: requiredInteger_(
        values[1], `基本パーツ ${sheetRow} 行目の順番`, 1, 999
      ),
      type: type,
      text: String(values[3] || ""),
      width: optionalInteger_(
        values[4], 33, `基本パーツ ${sheetRow} 行目の幅`,
        type === "年組氏名欄" ? 8 : 3, 33
      ),
      height: optionalInteger_(
        values[5],
        defaultHeight,
        `基本パーツ ${sheetRow} 行目の高さ`,
        2,
        20
      ),
      gap: optionalInteger_(
        values[6], 1, `基本パーツ ${sheetRow} 行目の後の空き`, 0, 10
      ),
    });
  });

  const questionRows = sheet.getRange(
    EXAM.questionStartRow,
    1,
    EXAM.questionRowCount,
    7
  ).getValues();

  const questionNumbers = new Set();

  questionRows.forEach(function (values, index) {
    if (values[0] !== true) return;

    const sheetRow = EXAM.questionStartRow + index;
    const number = String(values[2] || "").trim();

    if (!number) {
      throw new Error(
        `大問 ${sheetRow} 行目：大問番号を入力してください。`
      );
    }

    if (questionNumbers.has(number)) {
      throw new Error(
        `大問番号「${number}」が重複しています。大問 ${sheetRow} 行目を確認してください。`
      );
    }

    questionNumbers.add(number);

    parts.push({
      kind: "question",
      sheetRow: sheetRow,
      order: requiredInteger_(
        values[1], `大問 ${sheetRow} 行目の順番`, 1, 999
      ),
      number: number,
      rubric: String(values[3] || "").trim(),
      title: String(sheet.getRange(sheetRow, 9).getDisplayValue() || "").trim(),
      points: String(values[4] === "" ? "" : values[4]).trim(),
      width: optionalInteger_(
        values[5], 33, `大問 ${sheetRow} 行目の幅`, 6, 33
      ),
      gap: optionalInteger_(
        values[6], 1, `大問 ${sheetRow} 行目の後の空き`, 0, 10
      ),
      rows: [],
    });
  });

  const questionsByNumber = new Map();

  parts.forEach(function (part) {
    if (part.kind === "question") {
      questionsByNumber.set(part.number, part);
    }
  });

  // 表示文字をそのまま使い、(1) などを数値に変換しない。
  const detailRows = sheet.getRange(
    EXAM.detailStartRow,
    10,
    EXAM.detailRowCount,
    EXAM.detailColumnCount
  ).getDisplayValues();

  detailRows.forEach(function (values, index) {
    const number = String(values[0] || "").trim();

    // 対象大問が空欄、または使用しない大問なら出力対象にしない。
    if (!number || !questionsByNumber.has(number)) return;

    const sheetRow = EXAM.detailStartRow + index;
    const labelsText = String(values[3] || "").trim();
    const weightsText = String(values[4] || "").trim();

    const labels = labelsText
      ? labelsText.split(",").map(function (value) {
          return value.trim();
        })
      : [""];

    const emptyPositions = labels
      .map(function (label, index) {
        return label === "" ? index + 1 : null;
      })
      .filter(function (position) {
        return position !== null;
      });

    if (labelsText !== "" && emptyPositions.length > 0) {
      console.error("小問ラベルの入力エラー: " + JSON.stringify({
        settingsRow: sheetRow,
        questionNumber: number,
        rawInput: labelsText,
        parsedLabels: labels,
        emptyPositions: emptyPositions,
      }));

      throw new Error(
        `小問行 ${sheetRow} 行目（大問 ${number}）：` +
        `小問ラベルの ${emptyPositions.join("・")} 番目が空欄です。` +
        `入力値「${labelsText}」を確認してください。`
      );
    }

    const weights = weightsText
      ? weightsText.split(",").map(function (value) {
          return Number(value.trim());
        })
      : labels.map(function () { return 1; });

    if (
      weights.length !== labels.length ||
      weights.some(function (weight) {
        return !Number.isFinite(weight) || weight <= 0;
      })
    ) {
      throw new Error(
        `小問行 ${sheetRow} 行目：幅比は小問ラベルと同じ個数の正数にしてください。`
      );
    }

    questionsByNumber.get(number).rows.push({
      sheetRow: sheetRow,
      index: requiredInteger_(
        values[1], `小問行 ${sheetRow} 行目の行番号`, 1, 999
      ),
      rowLabel: String(values[1]),
      labels: labels,
      weights: weights,
      height: optionalInteger_(
        values[5], 2, `小問行 ${sheetRow} 行目の高さ`, 1, 10
      ),
      pattern: answerPattern_(values[7], labelsText, sheetRow),
      path: String(values[8] || `${number}-${values[1]}`),
      frameCount: effectiveFrameCount_(values, labels, answerPattern_(values[7], labelsText, sheetRow), sheetRow),
      itemsPerFrame: optionalInteger_(values[10], 1, `小問行 ${sheetRow} 行目の記入項目数`, 1, 10),
      evidence: String(values[11] || ""),
    });
  });

  parts.forEach(function (part) {
    if (part.kind !== "question") return;

    if (part.rows.length === 0) {
      throw new Error(
        `大問 ${part.sheetRow} 行目「${part.number}」：対応する小問行を③の表に1行以上入力してください。`
      );
    }

    part.rows.sort(function (a, b) {
      return a.index - b.index || a.sheetRow - b.sheetRow;
    });

    const usedIndexes = new Set();

    part.rows.forEach(function (row) {
      if (usedIndexes.has(row.index)) {
        throw new Error(
          `大問「${part.number}」：小問行の行番号 ${row.index} が重複しています。`
        );
      }

      usedIndexes.add(row.index);
      if (row.pattern === "記述欄" && row.frameCount > part.width - 2) {
        throw new Error(`小問行 ${row.sheetRow} 行目：枠数に対して大問の幅が足りません。`);
      }
      if (row.pattern === "小問複合枠") {
        if (!labelsForQuestionRow_(row)) throw new Error(`小問行 ${row.sheetRow} 行目：小問複合枠にはラベルを入力してください。`);
        validateAnswerWidths_(part, row);
      }
    });

    part.height =
      1 +
      part.rows.reduce(function (total, row) {
        return total + row.height;
      }, 0);
  });

  if (parts.length === 0) {
    throw new Error(
      "使用する基本パーツまたは大問にチェックを付けてください。"
    );
  }

  parts.sort(function (a, b) {
    return a.order - b.order || a.sheetRow - b.sheetRow;
  });

  return parts;
}

/**
 * 左上から配置し、収まらないパーツから右上へ送る。
 * 最終位置まで計算してから、出力シートを書き換える。
 */
function calculatePlacements_(parts) {
  const placements = [];
  let side = "左";
  let nextRow = EXAM.firstRow;

  parts.forEach(function (part) {
    const height = part.height;

    if (height > EXAM.outputRows - EXAM.firstRow + 1) {
      throw new Error(
        `${partName_(part)}は1つの領域の高さを超えます。小問行の高さを減らしてください。`
      );
    }

    if (nextRow + height - 1 > EXAM.outputRows) {
      if (side === "右") {
        throw new Error(
          `${partName_(part)}がB4横の右側にも収まりません。行の高さや空きを減らしてください。`
        );
      }

      side = "右";
      nextRow = EXAM.firstRow;
    }

    placements.push({
      part: part,
      side: side,
      row: nextRow,
      column:
        side === "左" ? EXAM.leftColumn : EXAM.rightColumn,
    });

    // 後の空きは次のパーツとの間だけに使う。
    nextRow += height + part.gap;
  });

  return placements;
}

/**
 * 小問枠の幅を描画前に確認する。
 */
function labelsForQuestionRow_(row) {
  return row.labels.length > 0 && row.labels.some(function (label) { return label !== ""; });
}

function validateAnswerWidths_(question, row) {
  const usableWidth = question.width - 2;
  const totalWeight = row.weights.reduce(function (sum, weight) {
    return sum + weight;
  }, 0);

  let cumulativeWeight = 0;
  let previousEnd = 0;

  row.weights.forEach(function (weight, index) {
    cumulativeWeight += weight;

    const end = index === row.weights.length - 1
      ? usableWidth
      : Math.round(
          usableWidth * cumulativeWeight / totalWeight
        );

    const segmentWidth = end - previousEnd;
    const minimumWidth = row.labels[index] ? 2 : 1;

    if (segmentWidth < minimumWidth) {
      throw new Error(
        `小問行 ${row.sheetRow} 行目：ラベル数に対して解答枠が狭すぎます。大問の幅または幅比を調整してください。`
      );
    }

    previousEnd = end;
  });
}

/**
 * 基本パーツを描画する。
 */
function drawBasic_(sheet, placement) {
  const part = placement.part;
  const row = placement.row;
  const column = placement.column;

  if (part.type === "考査見出し") {
    drawBox_(
      sheet,
      row,
      column,
      part.height,
      part.width,
      part.text || "考査",
      true
    );
    return;
  }

  if (part.type === "年組氏名欄") {
    // 「年｜記入欄｜組｜記入欄｜番｜記入欄｜氏名｜広い記入欄」。
    // 幅20列では 1:2:1:2:1:2:2:9 列。氏名の記入欄を最も広くする。
    const items = ["年", "", "組", "", "番", "", "氏名", ""];
    const weights = [1, 3, 1, 3, 1, 3, 2, 11];
    let usedColumns = 0;

    items.forEach(function (value, index) {
      const remainingItems = items.length - index - 1;
      const segmentWidth = index === items.length - 1
        ? part.width - usedColumns
        : Math.min(
            part.width - usedColumns - remainingItems,
            Math.max(1, Math.round(part.width * weights[index] / 25))
          );
      drawBox_(sheet, row, column + usedColumns, part.height,
        segmentWidth, value, false);
      usedColumns += segmentWidth;
    });

    return;
  }

  if (part.type === "得点枠") {
    const labels = ["知・技", "思・判・表", "合計"];
    const baseWidth = Math.floor(part.width / labels.length);
    let currentColumn = column;

    labels.forEach(function (label, index) {
      const width = index === labels.length - 1
        ? column + part.width - currentColumn
        : baseWidth;

      drawBox_(sheet, row, currentColumn, 1, width, label, false);
      drawBox_(
        sheet,
        row + 1,
        currentColumn,
        part.height - 1,
        width,
        "",
        false
      );

      currentColumn += width;
    });
  }
}

/**
 * 大問と、その小問行を描画する。
 */
function drawQuestion_(sheet, placement) {
  const question = placement.part;
  const row = placement.row;
  const column = placement.column;

  let title = "大問 " + question.number;

  if (question.rubric) {
    title += "　［" + question.rubric + "］";
  }

  if (question.points) {
    title += "　" + question.points + (/^\d+(?:\.\d+)?$/.test(question.points) ? "点" : "");
  }

  sheet.getRange(row, column, 1, question.width)
    .merge()
    .setNumberFormat("@")
    .setValue(title)
    .setFontFamily("Noto Sans JP")
    .setFontSize(10)
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  let currentRow = row + 1;

  question.rows.forEach(function (detail) {
    drawAnswerRow_(
      sheet,
      currentRow,
      column,
      question.width,
      detail
    );

    currentRow += detail.height;
  });
}

/**
 * 小問行の左端番号と、幅比に応じた解答欄を描画する。
 */
function drawAnswerRow_(sheet, row, column, width, detail) {
  const numberWidth = 2;

  drawBox_(
    sheet,
    row,
    column,
    detail.height,
    numberWidth,
    detail.rowLabel,
    false
  );

  const answerColumn = column + numberWidth;
  const usableWidth = width - numberWidth;

  if (detail.pattern === "記述欄") {
    const count = detail.frameCount || 1;
    for (let i = 0; i < count; i++) {
      const from = Math.round(usableWidth * i / count);
      const to = i === count - 1 ? usableWidth : Math.round(usableWidth * (i + 1) / count);
      drawBox_(sheet, row, answerColumn + from, detail.height, to - from, "", false);
    }
    return;
  }

  const totalWeight = detail.weights.reduce(function (sum, weight) {
    return sum + weight;
  }, 0);

  let cumulativeWeight = 0;
  let previousEnd = 0;

  detail.labels.forEach(function (label, index) {
    cumulativeWeight += detail.weights[index];

    const end = index === detail.labels.length - 1
      ? usableWidth
      : Math.round(
          usableWidth * cumulativeWeight / totalWeight
        );

    const segmentWidth = end - previousEnd;
    const segmentColumn = answerColumn + previousEnd;

    if (label) {
      drawBox_(
        sheet,
        row,
        segmentColumn,
        detail.height,
        1,
        label,
        false
      );

      drawBox_(
        sheet,
        row,
        segmentColumn + 1,
        detail.height,
        segmentWidth - 1,
        "",
        false
      );
    } else {
      drawBox_(
        sheet,
        row,
        segmentColumn,
        detail.height,
        segmentWidth,
        "",
        false
      );
    }

    previousEnd = end;
  });
}

function drawBox_(sheet, row, column, height, width, value, bold) {
  const range = sheet.getRange(
    row,
    column,
    height,
    width
  );

  if (height > 1 || width > 1) {
    range.merge();
  }

  range
    .setNumberFormat("@")
    .setValue(String(value === undefined || value === null ? "" : value))
    .setFontFamily("Noto Sans JP")
    .setFontSize(9)
    .setFontWeight(bold ? "bold" : "normal")
    .setHorizontalAlignment(bold ? "left" : "center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false);
}

/**
 * 出力シートをB4横想定のグリッドにする。
 */
function configureOutputSheet_(sheet) {
  ensureSize_(
    sheet,
    EXAM.outputRows,
    EXAM.outputColumns
  );

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(
    1,
    EXAM.outputColumns,
    EXAM.cellSizePx
  );
  sheet.setRowHeights(
    1,
    EXAM.outputRows,
    EXAM.cellSizePx
  );
  sheet.getRange(
    1,
    1,
    EXAM.outputRows,
    EXAM.outputColumns
  ).setBackground("#ffffff");
}

function requiredInteger_(value, name, minimum, maximum) {
  const number = Number(value);

  if (
    value === "" ||
    value === null ||
    !Number.isInteger(number) ||
    number < minimum ||
    number > maximum
  ) {
    throw new Error(
      `${name}は${minimum}～${maximum}の整数で入力してください。`
    );
  }

  return number;
}

function optionalInteger_(
  value,
  defaultValue,
  name,
  minimum,
  maximum
) {
  if (value === "" || value === null) {
    return defaultValue;
  }

  return requiredInteger_(
    value,
    name,
    minimum,
    maximum
  );
}

function partName_(part) {
  return part.kind === "basic"
    ? `基本パーツ ${part.sheetRow} 行目「${part.type}」`
    : `大問 ${part.sheetRow} 行目「${part.number}」`;
}

function ensureSize_(sheet, rows, columns) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      rows - sheet.getMaxRows()
    );
  }

  if (sheet.getMaxColumns() < columns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      columns - sheet.getMaxColumns()
    );
  }
}

function resetSheet_(sheet) {
  const all = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  all.breakApart();
  all.clearDataValidations();
  sheet.clear();
}

/**
 * パーツ編集サイドバーを表示
 */
function showPartEditor() {
  const html = HtmlService.createHtmlOutputFromFile("PartEditor")
    .setTitle("パーツ編集");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 編集対象のパーツ一覧を取得
 */
function getEditableParts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);
  if (sheet) ensureDetailPatternColumn_(sheet);
  const basicSheet = getBasicSheet_();

  if (!sheet) {
    throw new Error("設定シートがありません。先に setupSheets() を実行してください。");
  }

  const result = [];

  // 基本パーツ: A=使用、B=順番、C=種類、D=文字、E=幅、F=高さ、G=空き
  const basics = basicSheet.getRange(
    EXAM.basicStartRow, 1, EXAM.basicRowCount, 7
  ).getValues();

  basics.forEach((r, i) => {
    if (r[0] !== true) return;

    const row = EXAM.basicStartRow + i;
    result.push({
      id: `basic:${row}`,
      kind: "basic",
      row,
      label: `基本パーツ ${row}行目：${r[2] || "種類未設定"}`,
      order: r[1],
      type: r[2],
      text: r[3],
      width: r[4],
      height: r[5],
      gap: r[6],
    });
  });

  // 大問: A=使用、B=順番、C=番号、D=観点、E=配点、F=幅、G=空き
  const questions = sheet.getRange(
    EXAM.questionStartRow, 1, EXAM.questionRowCount, 7
  ).getValues();

  const details = sheet.getRange(
    EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount
  ).getDisplayValues();

  questions.forEach((r, i) => {
    if (r[0] !== true) return;

    const row = EXAM.questionStartRow + i;
    const number = String(r[2] || "").trim();
    const matchingDetails = details
      .map((detail, detailIndex) => ({ detail, detailIndex }))
      .filter(x => String(x.detail[0] || "").trim() === number);

    const totalHeight = matchingDetails.reduce((sum, x) => {
      return sum + (Number(x.detail[5]) || 2);
    }, 1); // 大問見出し分の1行

    result.push({
      id: `question:${row}`,
      kind: "question",
      row,
      label: `大問 ${number || "(番号未入力)"}`,
      order: r[1],
      number,
      rubric: r[3],
      title: sheet.getRange(row,9).getDisplayValue(),
      points: r[4],
      width: r[5],
      height: totalHeight,
      gap: r[6],
      detailCount: matchingDetails.length,
      details: matchingDetails.sort((a,b) => Number(a.detail[1]) - Number(b.detail[1]))
        .map(x => ({ sheetRow: EXAM.detailStartRow + x.detailIndex, index: x.detail[1],
          rowLabel: x.detail[2], labels: x.detail[3], weights: x.detail[4],
          height: x.detail[5] || 2, pattern: answerPattern_(x.detail[7], x.detail[3], EXAM.detailStartRow + x.detailIndex),
          path:x.detail[8], frameCount:x.detail[9], itemsPerFrame:x.detail[10], evidence:x.detail[11] })),
    });
  });

  return result.sort((a, b) =>
    (Number(a.order) || 0) - (Number(b.order) || 0) ||
    a.row - b.row
  );
}

/**
 * サイドバーから編集内容を保存して出力を再生成
 */
function updateEditablePart(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);

  const part = findEditablePart_(data.id);
  if (!part) throw new Error("編集対象が見つかりません。");

  const width = Number(data.width);
  const order = Number(data.order);
  const gap = Number(data.gap);

  if (!Number.isInteger(order) || order < 1) {
    throw new Error("順番は1以上の整数で入力してください。");
  }
  if (!Number.isInteger(width) || width < 3 || width > 33) {
    throw new Error("幅は3～33の整数で入力してください。");
  }
  if (!Number.isInteger(gap) || gap < 0 || gap > 10) {
    throw new Error("後の空きは0～10の整数で入力してください。");
  }

  if (part.kind === "basic") {
    const basicSheet = getBasicSheet_();
    const type = String(basicSheet.getRange(part.row, 3).getDisplayValue()).trim();
    if (type === "年組氏名欄" && width < 8) {
      throw new Error("年組氏名欄の幅は8～33の整数で入力してください。");
    }
    const height = Number(data.height);
    if (!Number.isInteger(height) || height < 2 || height > 20) {
      throw new Error("高さは2～20の整数で入力してください。");
    }

    basicSheet.getRange(part.row, 2).setValue(order);
    basicSheet.getRange(part.row, 4).setValue(data.text || "");
    basicSheet.getRange(part.row, 5, 1, 3)
      .setValues([[width, height, gap]]);
  } else {
    if (width < 6) {
      throw new Error("大問の幅は6～33の整数で入力してください。");
    }

    const height = Number(data.height);
    if (!Number.isInteger(height) || height < 2) {
      throw new Error("大問の高さは2以上の整数で入力してください。");
    }

    sheet.getRange(part.row, 2).setValue(order);
    sheet.getRange(part.row, 9).setValue(data.title || "");
    sheet.getRange(part.row, 4, 1, 4)
      .setValues([[
        data.rubric || "",
        data.points || "",
        width,
        gap,
      ]]);

    // 旧パネルの互換性: 高さの再配分は行わず、小問ごとの設定を維持する。
  }

  generateAnswerSheet();
  return "設定を更新して出力を再生成しました。";
}

/**
 * 選択パーツを複製
 */
function duplicateEditablePart(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);
  const part = findEditablePart_(id);

  if (!part) throw new Error("複製対象が見つかりません。");

  if (part.kind === "basic") {
    const basicSheet = getBasicSheet_();
    const targetRow = findEmptyRow_(
      basicSheet, EXAM.basicStartRow, EXAM.basicRowCount, 1
    );
    if (!targetRow) throw new Error("基本パーツの空き行がありません。");

    const values = basicSheet.getRange(part.row, 1, 1, 7).getValues()[0];
    values[1] = Number(values[1]) + 1;
    basicSheet.getRange(targetRow, 1, 1, 7).setValues([values]);
    basicSheet.getRange(targetRow, 1).insertCheckboxes().setValue(true);
  } else {
    const targetRow = findEmptyRow_(
      sheet, EXAM.questionStartRow, EXAM.questionRowCount, 1
    );
    if (!targetRow) throw new Error("大問の空き行がありません。");

    const values = sheet.getRange(part.row, 1, 1, 7).getValues()[0];
    const oldNumber = String(values[2]);
    const newNumber = makeCopyQuestionNumber_(sheet, oldNumber);

    values[1] = Number(values[1]) + 1;
    values[2] = newNumber;
    sheet.getRange(targetRow, 1, 1, 7).setValues([values]);
    sheet.getRange(targetRow, 9).setValue(sheet.getRange(part.row, 9).getValue());
    sheet.getRange(targetRow, 1).insertCheckboxes().setValue(true);

    copyQuestionDetails_(sheet, oldNumber, newNumber);
  }

  generateAnswerSheet();
  return "パーツを複製して出力を更新しました。";
}

/**
 * 選択パーツを削除
 */
function deleteEditablePart(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);
  const part = findEditablePart_(id);

  if (!part) throw new Error("削除対象が見つかりません。");

  (part.kind === "basic" ? getBasicSheet_() : sheet)
    .getRange(part.row, 1).setValue(false);

  if (part.kind === "question") {
    const detailValues = sheet.getRange(
      EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount
    ).getDisplayValues();

    detailValues.forEach((r, i) => {
      if (String(r[0] || "").trim() === part.number) {
        sheet.getRange(EXAM.detailStartRow + i, 10, 1, EXAM.detailColumnCount).clearContent();
      }
    });
  }

  generateAnswerSheet();
  return "パーツを削除して出力を更新しました。";
}

/**
 * 選択パーツの順番を上下に変更
 */
function moveEditablePart(id, direction) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);
  const part = findEditablePart_(id);

  if (!part) throw new Error("対象パーツが見つかりません。");

  const orderCell = (part.kind === "basic" ? getBasicSheet_() : sheet)
    .getRange(part.row, 2);
  const currentOrder = Number(orderCell.getValue());
  const nextOrder = currentOrder + (direction === "up" ? -1 : 1);

  if (nextOrder < 1) {
    throw new Error("順番は1未満にできません。");
  }

  orderCell.setValue(nextOrder);
  generateAnswerSheet();
  return "順番を変更して出力を更新しました。";
}

/**
 * IDから設定行を取得
 */
function findEditablePart_(id) {
  const [kind, rowText] = String(id || "").split(":");
  const row = Number(rowText);

  if (!Number.isInteger(row)) return null;

  if (
    kind === "basic" &&
    row >= EXAM.basicStartRow &&
    row < EXAM.basicStartRow + EXAM.basicRowCount
  ) {
    return { kind, row };
  }

  if (
    kind === "question" &&
    row >= EXAM.questionStartRow &&
    row < EXAM.questionStartRow + EXAM.questionRowCount
  ) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(EXAM.settingsSheetName);
    const number = String(sheet.getRange(row, 3).getDisplayValue()).trim();
    return { kind, row, number };
  }

  return null;
}

function getBasicSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.basicSheetName);
  if (!sheet) throw new Error("「基本パーツ」シートがありません。setupSheets() を実行してください。");
  return sheet;
}

/**
 * 大問の指定高さを小問行へ割り当てる
 */
function setQuestionHeight_(questionNumber, totalHeight) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.settingsSheetName);
  const details = sheet.getRange(
    EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount
  ).getDisplayValues();

  const matchingRows = [];

  details.forEach((r, i) => {
    if (String(r[0] || "").trim() === questionNumber) {
      matchingRows.push(EXAM.detailStartRow + i);
    }
  });

  if (!matchingRows.length) {
    throw new Error("大問に対応する小問行がありません。");
  }

  const available = Math.max(totalHeight - 1, matchingRows.length);
  const base = Math.floor(available / matchingRows.length);
  let remainder = available % matchingRows.length;

  matchingRows.forEach(row => {
    const rowHeight = base + (remainder > 0 ? 1 : 0);
    remainder -= remainder > 0 ? 1 : 0;
    sheet.getRange(row, 15).setValue(rowHeight);
  });
}

/**
 * 大問番号の重複しない複製番号を作る
 */
function makeCopyQuestionNumber_(sheet, originalNumber) {
  const values = sheet.getRange(
    EXAM.questionStartRow, 3, EXAM.questionRowCount, 1
  ).getDisplayValues().flat();

  let suffix = 2;
  let candidate = `${originalNumber}_コピー`;

  while (values.includes(candidate)) {
    candidate = `${originalNumber}_コピー${suffix}`;
    suffix++;
  }

  return candidate;
}

/**
 * 小問行を複製先の大問番号に付け替える
 */
function copyQuestionDetails_(sheet, oldNumber, newNumber) {
  const range = sheet.getRange(
    EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount
  );
  const values = range.getDisplayValues();
  const sourceRows = values.filter(r =>
    String(r[0] || "").trim() === oldNumber
  );

  const emptyRows = [];
  values.forEach((r, i) => {
    if (!r.some(value => value !== "")) {
      emptyRows.push(EXAM.detailStartRow + i);
    }
  });

  if (sourceRows.length > emptyRows.length) {
    throw new Error("小問行の空きが足りないため、大問を複製できません。");
  }

  sourceRows.forEach((r, i) => {
    const targetRow = emptyRows[i];
    r[0] = newNumber;
    sheet.getRange(targetRow, 10, 1, EXAM.detailColumnCount).setNumberFormat("@");
    sheet.getRange(targetRow, 10, 1, EXAM.detailColumnCount).setValues([r]);
  });
}

/**
 * 表の空き行を探す
 */
function findEmptyRow_(sheet, startRow, rowCount, column) {
  const values = sheet.getRange(startRow, column, rowCount, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== true && values[i][0] !== "") {
      continue;
    }

    const rowValues = sheet.getRange(startRow + i, 1, 1, 7).getValues()[0];
    if (rowValues.every(value => value === "" || value === false)) {
      return startRow + i;
    }
  }

  return null;
}

/**
 * 「問題文から追加」タブ用の解析・登録処理。
 * 同名の処理を別の ImportQuestion.gs にも残さないでください。
 */

function parseQuestionText(text) {
  const source = String(text || '').replace(/\r\n?/g, '\n');
  // 「4-1.」のような節番号がある場合、これを小問グループの境界にする。
  // 本文中の (1)、(a)、① はそのグループの解答欄ラベルとして扱う。
  const sectionMarker = /^[ \t　]*([0-9０-９]+)[-－–]([0-9０-９]+)[.．、][ \t　]*/gm;
  const sections = [];
  let match;
  while ((match = sectionMarker.exec(source)) !== null) {
    sections.push({at: match.index, bodyAt: sectionMarker.lastIndex,
      questionNumber: match[1].normalize('NFKC'),
      rowLabel: match[2].normalize('NFKC')});
  }
  if (sections.length) {
    const parents = new Set(sections.map(section => section.questionNumber));
    if (parents.size > 1) {
      throw new Error('異なる大問番号が混在しています。大問1件分ずつ貼り付けてください。');
    }
    return sections.map((section, index) => {
      const body = source.slice(section.bodyAt,
        sections[index + 1] ? sections[index + 1].at : source.length);
      const labels = collectAnswerLabels_(body, '', true);
      return {
        questionNumber: section.questionNumber,
        rowLabel: section.rowLabel,
        labels: labels,
        weights: labels.map(() => 1),
        height: 2,
        sourceText: body.trim(),
        suggestedPattern: labels.length ? 'sub_parens' : 'essay'
      };
    });
  }
  // 従来の行頭括弧形式も引き続き解析する。
  const marker = /^[ \t　]*[（(]([0-9０-９a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+)[）)]/gm;
  const points = [];
  while ((match = marker.exec(source)) !== null) {
    if (isKanaMarker_(match[1])) continue;
    points.push({at: match.index + match[0].search(/[（(]/), label: match[1]});
  }
  return points.map((point, index) => {
    const chunk = source.slice(point.at, points[index + 1] ? points[index + 1].at : source.length);
    const own = '(' + point.label.normalize('NFKC') + ')';
    const labels = collectAnswerLabels_(chunk, own, false);
    return {
      rowLabel: point.label.normalize('NFKC'), labels: labels,
      weights: labels.map(() => 1), height: 2,
      sourceText: chunk.trim(),
      suggestedPattern: labels.length ? 'sub_parens' : 'essay'
    };
  });
}

function collectAnswerLabels_(chunk, own, preferLineStarts) {
  const symbols = /([（(][0-9０-９a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+[）)]|\[[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+\]|〔[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+〕|【[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+】|[①-⑳㉑-㉟])/g;
  const candidateLines = chunk.split('\n').map(line => {
    const m = /^[ \t　]*([（(][0-9０-９a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+[）)]|\[[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+\]|〔[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+〕|【[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+】|[①-⑳㉑-㉟])/.exec(line);
    return m ? m[1] : null;
  }).filter(Boolean);
  const candidates = preferLineStarts && candidateLines.length ? candidateLines : (chunk.match(symbols) || []);
  const seen = new Set(), labels = [];
  candidates.forEach(symbol => {
    if (normalizeMarker_(symbol) === own || isKanaSymbol_(symbol) || seen.has(symbol)) return;
    seen.add(symbol);
    labels.push(symbol);
  });
  return labels;
}

function normalizeMarker_(value) {
  const m = value.match(/^[（(]([^）)]+)[）)]$/);
  return m ? '(' + m[1].normalize('NFKC') + ')' : value;
}

function isKanaMarker_(value) {
  return /^[ァ-ヶー]+$/.test(value.normalize('NFKC'));
}

function isKanaSymbol_(value) {
  const m = value.match(/^[（(]([^）)]+)[）)]$/);
  return !!m && isKanaMarker_(m[1]);
}

/** 解析結果を確認した後、大問表と小問行表へ同時に登録する。 */
function importQuestionFromText(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新中です。少し待ってから再実行してください。');
  try {
    return importQuestionFromTextLocked_(payload);
  } finally {
    lock.releaseLock();
  }
}

function importQuestionFromTextLocked_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXAM.settingsSheetName);
  if (!sheet) throw new Error('先に setupSheets() で設定シートを作ってください。');
  ensureDetailPatternColumn_(sheet);
  const number = String(payload.number || '').trim();
  if (!number) throw new Error('大問番号を入力してください。');
  if (/^[=+\-@]/.test(number)) throw new Error('大問番号の先頭に =、+、-、@ は使えません。');
  const order = importInt_(payload.order, '順番', 1, 999);
  const width = importInt_(payload.width, '幅', 6, 33);
  const gap = importInt_(payload.gap, '後の空き', 0, 10);
  const rows = payload.rows;
  if (!Array.isArray(rows) || !rows.length) throw new Error('小問が検出されていません。問題文の行頭に (1) などを入力してください。');

  const questionValues = sheet.getRange(EXAM.questionStartRow, 1, EXAM.questionRowCount, 7).getValues();
  const matchingIndexes = [];
  questionValues.forEach(function (r, i) {
    if (String(r[2] || '').trim() === number) matchingIndexes.push(i);
  });
  if (matchingIndexes.length > 1) {
    throw new Error('大問番号「' + number + '」が複数行にあります。設定表を修正してください。');
  }
  const existingIndex = matchingIndexes.length ? matchingIndexes[0] : -1;
  const questionIndex = existingIndex >= 0 ? existingIndex :
    questionValues.findIndex(function (r) { return r.every(function (v) { return v === '' || v === false; }); });
  if (questionIndex < 0) throw new Error('大問表の空き行がありません。');

  const details = sheet.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount).getValues();
  const emptyIndexes = [];
  const oldDetailIndexes = [];
  details.forEach(function (r, i) {
    if (r.every(function (v) { return v === '' || v === false; })) emptyIndexes.push(i);
    if (existingIndex >= 0 && String(r[0] || '').trim() === number) oldDetailIndexes.push(i);
  });

  // 1回目の呼び出しは競合情報だけを返し、シートには一切書き込まない。
  const fingerprint = existingIndex >= 0 ? importQuestionFingerprint_(questionValues[existingIndex], oldDetailIndexes.map(function (i) {
    return { index: i, values: details[i] };
  })) : '';
  if (existingIndex >= 0 && payload.replaceExisting !== true) {
    return {
      conflict: true,
      number: number,
      questionRow: EXAM.questionStartRow + existingIndex,
      detailCount: oldDetailIndexes.length,
      fingerprint: fingerprint
    };
  }
  if (existingIndex >= 0 && String(payload.existingFingerprint || '') !== fingerprint) {
    throw new Error('確認後に大問「' + number + '」の内容が変わりました。もう一度確認してください。');
  }
  if (existingIndex < 0 && payload.replaceExisting === true) {
    throw new Error('確認した大問が見つかりません。もう一度操作してください。');
  }
  const availableIndexes = oldDetailIndexes.concat(emptyIndexes);
  if (availableIndexes.length < rows.length) throw new Error('小問行表の空き行が足りません。');

  const prepared = rows.map(function (row, i) {
    const labels = Array.isArray(row.labels) ? row.labels.map(function (x) { return String(x).trim(); }) : [];
    const weights = Array.isArray(row.weights) ? row.weights.map(Number) : [];
    if (labels.length === 0) {
      if (weights.length !== 0 && (weights.length !== 1 || weights[0] <= 0)) throw new Error((i + 1) + '行目の幅比を確認してください。');
    } else if (labels.some(function (x) { return !x || x.includes(','); }) ||
        weights.length !== labels.length || weights.some(function (x) { return !Number.isFinite(x) || x <= 0; })) {
      throw new Error((i + 1) + '行目はラベルと幅比を同じ個数の正数で入力してください。');
    }
    const selectedPattern = row.pattern || (row.suggestedPattern === 'sub_parens' ? '小問複合枠' :
      row.suggestedPattern === 'essay' ? '記述欄' : '');
    const pattern = answerPattern_(selectedPattern, labels.join(','), EXAM.detailStartRow + i);
    if (pattern === '小問複合枠' && !labels.length) {
      throw new Error((i + 1) + '行目：小問複合枠にはラベルを入力してください。');
    }
    return [number, i + 1, String(row.rowLabel || ''), labels.join(','), weights.length ? weights.join(',') : '1',
      importInt_(row.height, (i + 1) + '行目の高さ', 1, 10), '問題文から追加', pattern,
      String(row.path || number + '-' + (i + 1)), labels.length || 1,
      Number(row.itemsPerFrame) || 1, String(row.evidence || '')];
  });

  const qRow = EXAM.questionStartRow + questionIndex;
  const oldQuestion = [questionValues[questionIndex].slice()];
  const targetIndexes = availableIndexes.slice(0, rows.length);
  const changedIndexes = Array.from(new Set(oldDetailIndexes.concat(targetIndexes)));
  const oldDetails = changedIndexes.map(function (i) {
    return { index: i, values: details[i].slice() };
  });
  try {
    sheet.getRange(qRow, 3).setNumberFormat('@');
    sheet.getRange(qRow, 1, 1, 7).setValues([[true, order, number, String(payload.rubric || ''), String(payload.points || ''), width, gap]]);
    oldDetailIndexes.forEach(function (i) {
      sheet.getRange(EXAM.detailStartRow + i, 10, 1, EXAM.detailColumnCount).clearContent();
    });
    prepared.forEach(function (values, i) {
      const target = sheet.getRange(EXAM.detailStartRow + targetIndexes[i], 10, 1, EXAM.detailColumnCount);
      target.setNumberFormat('@');
      target.setValues([values]);
    });
    // 出力を書き換える前に、既存の入力を含む全体を検査する。
    calculatePlacements_(readParts_(sheet));
  } catch (e) {
    sheet.getRange(qRow, 1, 1, 7).setValues(oldQuestion);
    oldDetails.forEach(function (item) {
      sheet.getRange(EXAM.detailStartRow + item.index, 10, 1, EXAM.detailColumnCount).setValues([item.values]);
    });
    throw e;
  }
  generateAnswerSheet();
  return (existingIndex >= 0 ? '大問 ' + number + ' の内容を更新しました。' :
    '大問 ' + number + ' と小問 ' + rows.length + ' 行を設定へ追加しました。');
}

function importQuestionFingerprint_(question, details) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify({ question: question, details: details })
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function importInt_(value, label, min, max) {
  const n = Number(value);
  if (value === '' || value === null || value === undefined || !Number.isInteger(n) || n < min || n > max) {
    throw new Error(label + 'は' + min + '～' + max + 'の整数にしてください。');
  }
  return n;
}


/** 大問と小問を一括保存。入力を検証し、出力に収まらなければ元に戻す。 */
function updateQuestionWithDetails(data) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新中です。少し待ってください。');
  try {
    return updateQuestionWithDetailsLocked_(data);
  } finally {
    lock.releaseLock();
  }
}

function updateQuestionWithDetailsLocked_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXAM.settingsSheetName);
  if (!sheet) throw new Error('設定シートがありません。');
  ensureDetailPatternColumn_(sheet);
  const part = findEditablePart_(data.id);
  if (!part || part.kind !== 'question' || sheet.getRange(part.row, 1).getValue() !== true) {
    throw new Error('編集する大問が見つかりません。再読み込みしてください。');
  }
  const original = sheet.getRange(part.row, 1, 1, 7).getValues()[0];
  const originalTitle = sheet.getRange(part.row, 9).getValue();
  const oldNumber = String(original[2] || '').trim();
  if (String(data.originalNumber || '') !== oldNumber) {
    throw new Error('編集開始後に大問番号が変わりました。再読み込みしてください。');
  }
  const number = String(data.number || '').trim();
  if (!number || /^[=+@-]/.test(number)) throw new Error('大問番号を入力してください。先頭に =、+、-、@ は使えません。');
  const others = sheet.getRange(EXAM.questionStartRow, 3, EXAM.questionRowCount, 1).getDisplayValues();
  if (others.some((r, i) => i !== part.row - EXAM.questionStartRow && String(r[0]).trim() === number)) {
    throw new Error('大問番号「' + number + '」は既に使われています。');
  }
  const order = requiredInteger_(data.order, '順番', 1, 999);
  const width = requiredInteger_(data.width, '大問の幅', 6, 33);
  const gap = requiredInteger_(data.gap, '後の空き', 0, 10);
  const rows = data.details;
  if (!Array.isArray(rows) || !rows.length) throw new Error('小問を1件以上入力してください。');
  const all = sheet.getRange(EXAM.detailStartRow, 10, EXAM.detailRowCount, EXAM.detailColumnCount).getValues();
  const existing = [], empty = [];
  all.forEach((r, i) => {
    if (String(r[0] || '').trim() === oldNumber) existing.push(i);
    else if (r.every(v => v === '' || v === false)) empty.push(i);
  });
  if (rows.length > existing.length + empty.length) throw new Error('小問行の空きが足りません。');
  const prepared = rows.map((item, i) => {
    const rowLabel = String(item.rowLabel == null ? '' : item.rowLabel).trim();
    if (/^[=+@-]/.test(rowLabel)) throw new Error('小問' + (i + 1) + 'の番号の先頭に =、+、-、@ は使えません。');
    const labels = Array.isArray(item.labels) ? item.labels.map(v => String(v).trim()) : [];
    const weights = Array.isArray(item.weights) ? item.weights.map(Number) : [];
    if (labels.some(v => !v || v.includes(',')) ||
        (labels.length && labels.length !== weights.length) ||
        weights.some(v => !Number.isFinite(v) || v <= 0)) {
      throw new Error('小問' + (i + 1) + 'のラベルと幅比を確認してください。');
    }
    const pattern = answerPattern_(item.pattern, labels.join(','), EXAM.detailStartRow + i);
    if (pattern === '小問複合枠' && !labels.length) {
      throw new Error('小問' + (i + 1) + '：小問複合枠にはラベルを1つ以上設定してください。');
    }
    if (pattern === '小問複合枠') {
      validateAnswerWidths_({width}, {labels, weights, sheetRow: EXAM.detailStartRow + i});
    }
    const height = requiredInteger_(item.height, '小問' + (i + 1) + 'の高さ', 1, 10);
    const count = optionalInteger_(item.frameCount == null ? '' : item.frameCount, pattern === '小問複合枠' ? labels.length : 1,
      '小問' + (i + 1) + 'の枠数', 1, 30);
    if (pattern === '小問複合枠' && count !== labels.length) throw new Error('小問' + (i + 1) + 'の枠数とラベル数が違います。');
    if (pattern === '記述欄' && count > width - 2) throw new Error('小問' + (i + 1) + 'の枠数に対して幅が足りません。');
    return [number, i + 1, rowLabel, labels.join(','),
      labels.length ? weights.join(',') : '1', height, 'パーツ編集', pattern,
      String(item.path || number + '-' + (i + 1)), count,
      optionalInteger_(item.itemsPerFrame == null ? '' : item.itemsPerFrame, 1, '小問' + (i + 1) + 'の記入項目数', 1, 10),
      String(item.evidence || '')];
  });
  const slots = existing.concat(empty).slice(0, prepared.length);
  const changed = Array.from(new Set(existing.concat(slots))).map(i => ({index: i, values: all[i].slice()}));
  try {
    sheet.getRange(part.row, 9).setValue(String(data.title || ""));
    sheet.getRange(part.row, 3).setNumberFormat('@');
    sheet.getRange(part.row, 1, 1, 7).setValues([[true, order, number,
      String(data.rubric || ''), String(data.points || ''), width, gap]]);
    existing.forEach(i => sheet.getRange(EXAM.detailStartRow + i, 10, 1, EXAM.detailColumnCount).clearContent());
    prepared.forEach((values, i) => sheet.getRange(EXAM.detailStartRow + slots[i], 10, 1, EXAM.detailColumnCount)
      .setNumberFormat('@').setValues([values]));
    SpreadsheetApp.flush();
    calculatePlacements_(readParts_(sheet));
  } catch (error) {
    sheet.getRange(part.row, 1, 1, 7).setValues([original]);
    sheet.getRange(part.row, 9).setValue(originalTitle);
    changed.forEach(x => sheet.getRange(EXAM.detailStartRow + x.index, 10, 1, EXAM.detailColumnCount)
      .setValues([x.values]));
    throw new Error('保存できませんでした。設定を戻しました。詳細: ' + (error.message || error));
  }
  generateAnswerSheet();
  return '大問 ' + number + ' と小問 ' + rows.length + ' 件を更新しました。';
}
