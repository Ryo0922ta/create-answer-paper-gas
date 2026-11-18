/**
 * 解答用紙ジェネレーター
 *
 * 作業データの内容から、B4横・左右2領域の「出力」シートを生成します。
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
 * 不足している設定・基本パーツ・出力シートを作る。既存シートは変更しない。
 */
function setupSheets() {
  return withOperationLog_('setup_sheets', function () {
    return setupSheetsLogged_();
  });
}

function setupSheetsLogged_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  if (!ss.getSheetByName(EXAM.basicSheetName)) {
    const basic = ss.insertSheet(EXAM.basicSheetName);
    ensureSize_(basic, 13, 8);
    configureBasicSheet_(basic, []);
    created.push('基本パーツ');
  }
  if (!ss.getSheetByName(EXAM.outputSheetName)) {
    withDraftLock_(function () {
      // 出力だけが削除された場合、残っている管理情報を勝手に上書きしない。
      const previous = readDraft_();
      if (previous.parts.length) {
        throw new Error('「出力」シートはありませんが、以前の作業データが残っています。作業データを確認してからセットアップしてください。');
      }
      const output = ss.insertSheet(EXAM.outputSheetName);
      try {
        configureOutputSheet_(output);
        const next = { parts: ['考査見出し', '得点枠', '年組氏名欄'].map(function (type, index) {
          return { id: Utilities.getUuid(), kind: 'basic', sheetRow: index + 1,
            sourceRow: '', order: index + 1, type: type, text: '', width: 33,
            height: type === '考査見出し' ? 2 : 4, gap: 1 };
        }) };
        commitDraft_(next);
      } catch (error) {
        // この実行で新規作成したシートだけを取り除き、再実行可能にする。
        ss.deleteSheet(output);
        throw error;
      }
    });
    created.push('出力');
  }
  logEvent_('info', 'setup_sheets', 'sheets_created', { createdCount: created.length });
  SpreadsheetApp.getUi().alert(created.length
    ? '「' + created.join('」「') + '」シートを作成しました。'
    : '「基本パーツ」「出力」シートはすでに存在します。');
}

/** モーダルへ渡す考査見出し・氏名欄・得点枠の元データ。 */
function configureBasicSheet_(sheet, previousRows) {
  sheet.getRange("A1:G2").merge()
    .setValue("基本パーツの元データ\n出力への追加は「パーツ編集」→「基本パーツ」から行います。")
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
  throw new Error("出力の一括再生成は利用できません。パーツ編集から操作してください。");
  SpreadsheetApp.getUi().alert(
    "編集中の内容から出力シートを更新しました。\n" +
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

  const savedBasic = sheet.getRange(DRAFT.basicHeaderRow, 1).getDisplayValue() === '使用';
  const basicSource = savedBasic ? sheet : basicSheet;
  const basicStartRow = savedBasic ? DRAFT.basicFirstRow : EXAM.basicStartRow;
  const basicRows = basicSource.getRange(
    basicStartRow,
    1,
    EXAM.basicRowCount,
    7
  ).getValues();

  basicRows.forEach(function (values, index) {
    if (values[0] !== true) return;

    const sheetRow = basicStartRow + index;
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
      logEvent_('warn', 'read_settings', 'empty_labels', {
        sheetRow: sheetRow, emptyCount: emptyPositions.length,
      });

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
      rowLabel: String(values[2] || values[1]),
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
 * 最終位置を計算し、描画前に収まることを確認する。
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
  answerRowLayout_(question.width, {
    rowLabel: row.rowLabel, pattern: '小問複合枠',
    labels: row.labels, weights: row.weights, frameCount: row.labels.length,
    sheetRow: row.sheetRow
  });
}

/** 出力シートと登録前プレビューで共用する、列単位の小問枠配置。 */
function answerRowLayout_(width, detail) {
  const usableWidth = width - 2;
  const boxes = [{ width: 2, value: detail.rowLabel || '', kind: 'number' }];
  if (!Number.isInteger(width) || width < 3) throw new Error('大問の幅を確認してください。');

  if (detail.pattern === '記述欄') {
    const count = Number(detail.frameCount || 1);
    if (!Number.isInteger(count) || count < 1 || count > usableWidth) {
      throw new Error('記述欄の枠数が大問の幅に収まりません。');
    }
    for (let i = 0; i < count; i++) {
      const from = Math.round(usableWidth * i / count);
      const to = i === count - 1 ? usableWidth : Math.round(usableWidth * (i + 1) / count);
      boxes.push({ width: to - from, value: '', kind: 'answer' });
    }
    return boxes;
  }

  if (detail.pattern !== '小問複合枠' || !Array.isArray(detail.labels) ||
      !Array.isArray(detail.weights) || !detail.labels.length ||
      detail.labels.length !== detail.weights.length ||
      detail.weights.some(function (weight) { return !Number.isFinite(weight) || weight <= 0; })) {
    throw new Error('小問のラベル、幅比、解答欄パターンを確認してください。');
  }
  const totalWeight = detail.weights.reduce(function (sum, weight) { return sum + weight; }, 0);
  if (!Number.isFinite(totalWeight) || totalWeight <= 0) {
    throw new Error('小問の幅比を確認してください。');
  }
  let cumulativeWeight = 0;
  let previousEnd = 0;
  detail.labels.forEach(function (label, index) {
    cumulativeWeight += detail.weights[index];
    const end = index === detail.labels.length - 1
      ? usableWidth : Math.round(usableWidth * cumulativeWeight / totalWeight);
    const segmentWidth = end - previousEnd;
    if (segmentWidth < (label ? 2 : 1)) {
      const location = detail.sheetRow ? `小問行 ${detail.sheetRow} 行目：` : '';
      throw new Error(location + 'ラベル数に対して解答枠が狭すぎます。大問の幅または幅比を調整してください。');
    }
    if (label) boxes.push({ width: 1, value: label, kind: 'label' });
    boxes.push({ width: segmentWidth - (label ? 1 : 0), value: '', kind: 'answer' });
    previousEnd = end;
  });
  return boxes;
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
  let currentColumn = column;
  answerRowLayout_(width, detail).forEach(function (box) {
    drawBox_(sheet, row, currentColumn, detail.height, box.width, box.value, false);
    currentColumn += box.width;
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
    EXAM.outputRows + 1,
    EXAM.outputColumns + 1
  );

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(
    1,
    sheet.getMaxColumns(),
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
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  ).setBackground("#e6e6e6");
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
 * パーツ編集ダイアログを表示
 */
function showPartEditor() {
  return withOperationLog_('open_editor', function () {
    return showPartEditorLogged_();
  });
}

function showPartEditorLogged_() {
  const html = HtmlService.createHtmlOutputFromFile("PartEditor")
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, "パーツ編集");
}

/**
 * 編集対象のパーツ一覧を取得
 */
function getEditableParts() {
  return getDraftEditableParts_();
}

function showDeletePartDialog() {
  const candidate = withOperationLog_('find_delete_candidate', function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== EXAM.outputSheetName) throw new Error('「出力」シートで削除するパーツ内のセルを選んでください。');
    const list = sheet.getActiveRangeList();
    const ranges = list && list.getRanges();
    if (!ranges || ranges.length !== 1) throw new Error('削除するパーツ内のセルを１つ選んでください。');
    const range = ranges[0];
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
      const merges = range.getMergedRanges();
      if (merges.length !== 1 || merges[0].getA1Notation() !== range.getA1Notation()) {
        throw new Error('削除するパーツ内のセルを１つ選んでください。');
      }
    }
    const selected = {row:range.getRow(), column:range.getColumn()};
    const signatures = outputDeleteSignatures_(readDraft_().parts);
    if (!signatures.length) throw new Error('照合できる登録済みパーツがありません。');
    const height = Math.max(...signatures.map(function (item) { return item.rect.height; }));
    const width = Math.max(...signatures.map(function (item) { return item.rect.width; }));
    const row = Math.max(1, selected.row - height + 1), column = Math.max(1, selected.column - width + 1);
    const search = {row:row, column:column,
      height:Math.min(sheet.getMaxRows(), selected.row + height - 1) - row + 1,
      width:Math.min(sheet.getMaxColumns(), selected.column + width - 1) - column + 1};
    const grid = outputReadRanges_(sheet, [search]);
    const found = outputFindDeleteCandidate_(grid, signatures, selected, search);
    const origin = found.signature.rect;
    if (!outputSamePosition_(origin, found.rect)) {
      const originalGrid = outputContains_(search, origin) ? grid : outputReadRanges_(sheet, [origin]);
      outputRejectDuplicatedDelete_(originalGrid, found.signature, found.rect);
    }
    return {part:found.signature.part, rect:found.rect};
  });
  const {part, rect} = candidate;
  const operationId = Utilities.getUuid();
  putDeleteOperation_(operationId, {state:'pending', partId:part.id, rect:rect, revision:deletePartRevision_(part)}, 120);
  const template = HtmlService.createTemplateFromFile('DeletePartDialog');
  template.operationId = operationId;
  template.label = part.kind === 'question' ? '大問 ' + part.number : part.type + (part.text ? '：' + part.text : '');
  template.range = outputA1_(rect);
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(420).setHeight(240), 'パーツの削除');
  return operationId;
}

function deletePartRevision_(part) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(part)));
}

function putDeleteOperation_(id, operation, ttl) {
  CacheService.getUserCache().put('delete_part_' + id, JSON.stringify(operation), ttl);
}

function readDeleteOperation_(id) {
  if (!/^[0-9a-f-]{36}$/i.test(id)) throw new Error('削除操作が無効です。');
  const saved = CacheService.getUserCache().get('delete_part_' + id);
  return saved ? JSON.parse(saved) : {state:'expired'};
}

function getDeletePartStatus(id) {
  const operation = readDeleteOperation_(id);
  return {state:operation.state, message:operation.message || ''};
}

function transitionPendingDelete_(id, state) {
  const lock = LockService.getUserLock();
  lock.waitLock(3000);
  try {
    const operation = readDeleteOperation_(id);
    if (operation.state !== 'pending') return operation;
    putDeleteOperation_(id, {...operation, state:state}, 600);
    return operation;
  } finally {
    lock.releaseLock();
  }
}

function cancelDeletePart(id) {
  transitionPendingDelete_(id, 'cancelled');
  return getDeletePartStatus(id);
}

function confirmDeletePart(id) {
  const operation = transitionPendingDelete_(id, 'running');
  if (operation.state !== 'pending') throw new Error('削除操作は終了済みか実行中です。');
  let result;
  try {
    result = withOperationLog_('delete_part', function () {
      return deleteDraftPart_(operation.partId, operation.revision, operation.rect);
    });
  } catch (error) {
    putDeleteOperation_(id, {state:'failed', message:error.message}, 600);
    throw error;
  }
  putDeleteOperation_(id, {state:'complete', message:result}, 600);
  return result;
}

function getSelectedOutputCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ranges = sheet.getActiveRangeList();
  const selected = ranges && ranges.getRanges();
  if (sheet.getName() !== EXAM.outputSheetName || !selected || selected.length !== 1 ||
      selected[0].getNumRows() !== 1 || selected[0].getNumColumns() !== 1) {
    throw new Error('出力シートで左上にしたいセルを１つ選択してください。');
  }
  return selected[0].getA1Notation();
}

function unregisterChangedPart(id) {
  return withDraftLock_(function () {
    const draft = readDraft_();
    const index = draft.parts.findIndex(function (p) { return p.id === id; });
    if (index < 0) throw new Error('対象パーツが見つかりません。再読み込みしてください。');
    const part = draft.parts[index], rect = outputRect_(part);
    if (!rect || !part.outputState.fingerprint) throw new Error('対象パーツの状態を確認できません。');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAM.outputSheetName);
    if (!sheet) throw new Error('出力シートが見つかりません。');
    const grid = outputReadRanges_(sheet, [rect]);
    if (outputFingerprint_(grid, rect) === part.outputState.fingerprint) throw new Error('出力と作業データが一致しているため登録解除できません。削除ボタンを使ってください。');
    draft.parts.splice(index, 1);
    writeDraft_(draft);
    return '作業データの登録を解除しました。出力シートは変更していません。';
  });
}

/**
 * パーツ編集ダイアログの基本パーツ変更を作業データと出力へ反映する。
 */
function updateEditablePart(data) {
  return withOperationLog_('edit_basic', function () {
    return updateEditablePartLogged_(data);
  });
}

function updateEditablePartLogged_(data) {
  return updateDraftBasicPart_(data);
}

/**
 * 選択パーツを複製
 */
function duplicateEditablePart(id) {
  return withOperationLog_('duplicate_part', function () {
    return duplicateEditablePartLogged_(id);
  });
}

function duplicateEditablePartLogged_(id) {
  return duplicateDraftPart_(id);
}

/**
 * 選択パーツを削除
 */
function deleteEditablePart(id) {
  return withOperationLog_('delete_part', function () {
    return deleteEditablePartLogged_(id);
  });
}

function deleteEditablePartLogged_(id) {
  return deleteDraftPart_(id);
}

/**
 * 選択パーツの順番を上下に変更
 */
function moveEditablePart(id, direction) {
  return withOperationLog_('move_part', function () {
    return moveEditablePartLogged_(id, direction);
  });
}

function moveEditablePartLogged_(id, direction) {
  return moveDraftPart_(id, direction);
}

function getBasicSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EXAM.basicSheetName);
  if (!sheet) throw new Error("「基本パーツ」シートがありません。setupSheets() を実行してください。");
  return sheet;
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
  const symbols = /([（(][ \t　]*[0-9０-９a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+[ \t　]*[）)]|\[[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+\]|〔[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+〕|【[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+】|[①-⑳㉑-㉟])/g;
  const candidateLines = chunk.split('\n').map(line => {
    const m = /^[ \t　]*([（(][ \t　]*[0-9０-９a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+[ \t　]*[）)]|\[[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+\]|〔[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+〕|【[0-9０-９a-zA-ZＡ-Ｚａ-ｚ]+】|[①-⑳㉑-㉟])/.exec(line);
    return m ? m[1] : null;
  }).filter(Boolean);
  const candidates = preferLineStarts && candidateLines.length ? candidateLines : (chunk.match(symbols) || []);
  const seen = new Set(), labels = [];
  candidates.forEach(symbol => {
    const normalized = normalizeMarker_(symbol);
    if (normalized === own || isKanaSymbol_(symbol) || seen.has(normalized)) return;
    seen.add(normalized);
    labels.push(normalized);
  });
  return labels;
}

function normalizeMarker_(value) {
  const m = value.match(/^[（(]([^）)]+)[）)]$/);
  return m ? '(' + m[1].trim().normalize('NFKC') + ')' : value;
}

function isKanaMarker_(value) {
  return /^[ァ-ヶー]+$/.test(value.normalize('NFKC'));
}

function isKanaSymbol_(value) {
  const m = value.match(/^[（(]([^）)]+)[）)]$/);
  return !!m && m[1] === m[1].trim() && isKanaMarker_(m[1]);
}

/** モーダルで編集した大問は作業データと出力だけに反映する。 */
function updateQuestionWithDetails(data) {
  return withOperationLog_('edit_question', function () {
    return updateQuestionWithDetailsLogged_(data);
  });
}

function updateQuestionWithDetailsLogged_(data) {
  return updateDraftQuestion_(data);
}
