/** シートの内容を保持しながらパーツ単位で出力を更新する。 */
function outputRect_(part) {
  const p = part.outputState && part.outputState.rect;
  return p && { row: p.row, column: p.column, height: p.height, width: p.width };
}

function outputOverlap_(a, b) {
  return a.row < b.row + b.height && b.row < a.row + a.height &&
    a.column < b.column + b.width && b.column < a.column + a.width;
}

function outputA1_(rect) {
  function column(n) {
    let value = '';
    while (n) { n--; value = String.fromCharCode(65 + n % 26) + value; n = Math.floor(n / 26); }
    return value;
  }
  return column(rect.column) + rect.row + ':' +
    column(rect.column + rect.width - 1) + (rect.row + rect.height - 1);
}

function outputReadGrid_(sheet) {
  return withOperationLog_('read_output', function () {
    return outputReadGridLogged_(sheet);
  });
}

function outputReadGridLogged_(sheet) {
  const ss = sheet.getParent();
  const result = Sheets.Spreadsheets.get(ss.getId(), {
    ranges: ["'" + sheet.getName().replace(/'/g, "''") + "'!A1:BT51"],
    fields: 'sheets(properties(sheetId,gridProperties(rowCount,columnCount)),data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,userEnteredFormat,note,dataValidation,textFormatRuns))),merges)',
  });
  if (!result.sheets || result.sheets.length !== 1) throw new Error('出力シートを読み取れませんでした。');
  const data = result.sheets[0];
  const props = data.properties, block = data.data && data.data[0];
  if (!props || (props.sheetId == null ? 0 : props.sheetId) !== sheet.getSheetId() || !props.gridProperties ||
      props.gridProperties.rowCount < EXAM.outputRows || props.gridProperties.columnCount < EXAM.outputColumns ||
      !block || (block.startRow || 0) !== 0 || (block.startColumn || 0) !== 0 ||
      (block.rowData != null && !Array.isArray(block.rowData)) ||
      (data.merges != null && !Array.isArray(data.merges))) {
    throw new Error('出力シートの範囲を完全に読み取れませんでした。配置を中止しました。');
  }
  if (sheet.getImages().length || sheet.getDrawings().length) {
    throw new Error('出力シートに画像または図形があります。重なりを確認できないため配置を中止しました。');
  }
  const rows = block.rowData || [];
  if (rows.length > EXAM.outputRows || rows.some(function (row) {
    return !row || (row.values != null && (!Array.isArray(row.values) || row.values.length > EXAM.outputColumns));
  })) throw new Error('出力シートのセル情報を読み取れませんでした。');
  const merges = data.merges || [];
  merges.forEach(outputMergeRect_);
  return { rows: rows, merges: merges };
}

/** 配置先のみを読み取り、通常の出力グリッドと同じ座標へ戻す。 */
function outputReadRanges_(sheet, rects) {
  return withOperationLog_('read_output_ranges', function () {
    const prefix = "'" + sheet.getName().replace(/'/g, "''") + "'!";
    const result = Sheets.Spreadsheets.get(sheet.getParent().getId(), {
      ranges: rects.map(function (r) { return prefix + outputA1_(r); }),
      fields: 'sheets(properties(sheetId,gridProperties(rowCount,columnCount)),data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,userEnteredFormat,note,dataValidation,textFormatRuns))),merges)',
    });
    if (sheet.getImages().length || sheet.getDrawings().length) {
      throw new Error('出力シートに画像または図形があります。重なりを確認できないため配置を中止しました。');
    }
    return outputGridFromRanges_(sheet, result, rects);
  });
}

function outputGridFromRanges_(sheet, result, rects) {
    const data = result && result.sheets && result.sheets.length === 1 && result.sheets[0];
    const props = data && data.properties;
    if (!props || (props.sheetId == null ? 0 : props.sheetId) !== sheet.getSheetId() || !props.gridProperties ||
        props.gridProperties.rowCount < EXAM.outputRows || props.gridProperties.columnCount < EXAM.outputColumns ||
        !Array.isArray(data.data) || data.data.length !== rects.length ||
        (data.merges != null && !Array.isArray(data.merges))) {
      throw new Error('配置先を完全に読み取れませんでした。配置を中止しました。');
    }
    const rows = [];
    data.data.forEach(function (block, index) {
      const rect = rects[index];
      if (!block || (block.startRow || 0) !== rect.row - 1 ||
          (block.startColumn || 0) !== rect.column - 1 ||
          (block.rowData != null && !Array.isArray(block.rowData)) ||
          (block.rowData || []).length > rect.height) {
        throw new Error('配置先のセル情報を読み取れませんでした。');
      }
      (block.rowData || []).forEach(function (row, rowOffset) {
        if (!row || (row.values != null && (!Array.isArray(row.values) || row.values.length > rect.width))) {
          throw new Error('配置先のセル情報を読み取れませんでした。');
        }
        const rowIndex = rect.row - 1 + rowOffset;
        if (!rows[rowIndex]) rows[rowIndex] = {values: []};
        (row.values || []).forEach(function (cell, colOffset) {
          rows[rowIndex].values[rect.column - 1 + colOffset] = cell;
        });
      });
    });
    const merges = data.merges || [];
    merges.forEach(outputMergeRect_);
    return {rows: rows, merges: merges};
}

function outputMergeRect_(m) {
  const row = m.startRowIndex == null ? 0 : m.startRowIndex;
  const col = m.startColumnIndex == null ? 0 : m.startColumnIndex;
  if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(m.endRowIndex) ||
      !Number.isInteger(m.endColumnIndex) || m.endRowIndex <= row || m.endColumnIndex <= col) {
    throw new Error('結合セルの範囲を読み取れませんでした。配置を中止しました。');
  }
  return { row: row + 1, column: col + 1, height: m.endRowIndex - row, width: m.endColumnIndex - col };
}

function outputCanonical_(value) {
  if (Array.isArray(value)) return value.map(outputCanonical_);
  if (value && typeof value === 'object') {
    const result = {};
    Object.keys(value).sort().forEach(function (key) {
      if (value[key] !== undefined && value[key] !== null) result[key] = outputCanonical_(value[key]);
    });
    return result;
  }
  return value;
}

function outputFingerprintCell_(cell) {
  const copy = Object.assign({}, cell || {});
  // 計算結果は空き判定に使うが、パーツ指紋には含めない。
  delete copy.effectiveValue;
  return copy;
}

function outputFingerprint_(grid, rect) {
  const cells = [];
  for (let row = rect.row - 1; row < rect.row - 1 + rect.height; row++) {
    const values = (grid.rows[row] && grid.rows[row].values) || [];
    const line = [];
    for (let col = rect.column - 1; col < rect.column - 1 + rect.width; col++) {
      line.push(outputFingerprintCell_(values[col]));
    }
    cells.push(line);
  }
  const merges = grid.merges.filter(function (m) {
    return outputOverlap_(rect, outputMergeRect_(m));
  }).map(function (m) {
    return [m.startRowIndex, m.endRowIndex, m.startColumnIndex, m.endColumnIndex];
  }).sort(function (a, b) { return JSON.stringify(a).localeCompare(JSON.stringify(b)); });
  return JSON.stringify(outputCanonical_({ cells: cells, merges: merges }));
}

function outputRelativeMerges_(merges, rect) {
  return merges.map(function (merge) {
    return merge.map(function (index, i) { return (index || 0) - (i < 2 ? rect.row : rect.column) + 1; });
  }).sort(function (a, b) { return JSON.stringify(a).localeCompare(JSON.stringify(b)); });
}

function outputRelativeFingerprint_(fingerprint, rect) {
  const data = JSON.parse(fingerprint);
  return JSON.stringify({cells:data.cells, merges:outputRelativeMerges_(data.merges, rect)});
}

function outputDeleteSignatures_(parts) {
  return parts.flatMap(function (part) {
    const rect = outputRect_(part);
    if (!rect || !part.outputState.fingerprint ||
        ![rect.row, rect.column, rect.height, rect.width].every(function (n) { return Number.isInteger(n) && n > 0; }) ||
        rect.height > EXAM.outputRows || rect.width > EXAM.outputColumns) return [];
    const saved = JSON.parse(part.outputState.fingerprint);
    // 空欄だけの登録や古い形式を、任意の空きセルと照合しない。
    if (!Array.isArray(saved.cells) || !Array.isArray(saved.merges) ||
        saved.cells.length !== rect.height || saved.cells.some(function (row) {
          return !Array.isArray(row) || row.length !== rect.width;
        }) || !saved.cells.some(function (row) {
          return row.some(function (cell) { return !outputCellBlank_(cell); });
        })) return [];
    return [{part:part, rect:rect, first:JSON.stringify(saved.cells[0][0]),
      fingerprint:outputRelativeFingerprint_(part.outputState.fingerprint, rect)}];
  });
}

function outputContains_(outer, inner) {
  return inner.row >= outer.row && inner.column >= outer.column &&
    inner.row + inner.height <= outer.row + outer.height &&
    inner.column + inner.width <= outer.column + outer.width;
}

function outputSamePosition_(a, b) {
  return a.row === b.row && a.column === b.column;
}

function outputRejectDuplicatedDelete_(grid, signature, rect) {
  const origin = signature.rect;
  if (!outputSamePosition_(origin, rect) &&
      outputRelativeFingerprint_(outputFingerprint_(grid, origin), origin) === signature.fingerprint) {
    throw new Error('登録位置にも同じパーツが残っています。削除するパーツを安全に特定できません。');
  }
}

function outputFindDeleteCandidate_(grid, signatures, selected, search) {
  let match = null;
  for (const signature of signatures) {
    const {height, width} = signature.rect;
    for (let row = Math.max(search.row, selected.row - height + 1); row <= selected.row; row++) {
      const values = (grid.rows[row - 1] || {}).values || [];
      for (let column = Math.max(search.column, selected.column - width + 1); column <= selected.column; column++) {
        const rect = {row:row, column:column, height:height, width:width};
        if (!outputContains_(search, rect) ||
            JSON.stringify(outputCanonical_(outputFingerprintCell_(values[column - 1]))) !== signature.first) continue;
        if (outputRelativeFingerprint_(outputFingerprint_(grid, rect), rect) !== signature.fingerprint) continue;
        if (match) throw new Error('選択セルの削除候補が複数あります。対象を安全に特定できません。');
        match = {signature:signature, rect:rect};
      }
    }
  }
  if (!match) throw new Error('選択セルから登録済みパーツを特定できません。パーツ内のセルを選び直してください。直接編集されたパーツは削除できません。');
  return match;
}

function outputSelectedRemovalPlan_(previous, next, selected, before) {
  const part = previous.parts.find(function (item) { return item.id === selected.partId; });
  const [signature] = outputDeleteSignatures_(part ? [part] : []);
  const rect = selected.rect;
  if (!signature || next.parts.length !== previous.parts.length - 1 ||
      next.parts.some(function (item) { return item.id === selected.partId; }) ||
      rect.height !== signature.rect.height || rect.width !== signature.rect.width ||
      outputRelativeFingerprint_(outputFingerprint_(before, rect), rect) !== signature.fingerprint) {
    throw new Error('確認後に削除対象が変わりました。セルを選び直してください。');
  }
  outputRejectDuplicatedDelete_(before, signature, rect);
  return {removals:[{part:part, rect:rect}], additions:[]};
}

function outputCellBlank_(cell) {
  if (!cell) return true;
  if (cell.userEnteredValue || cell.effectiveValue || cell.note || cell.dataValidation || cell.textFormatRuns) return false;
  const format = cell.userEnteredFormat || {};
  const keys = Object.keys(format).filter(function (key) {
    if (key === 'backgroundColor') {
      const color = format[key] || {};
      return !(color.red === 1 && color.green === 1 && color.blue === 1 && (color.alpha === undefined || color.alpha === 1));
    }
    if (key === 'backgroundColorStyle') {
      const rgb = (format[key] || {}).rgbColor || {};
      return !(rgb.red === 1 && rgb.green === 1 && rgb.blue === 1 && (rgb.alpha === undefined || rgb.alpha === 1));
    }
    return true;
  });
  return keys.length === 0;
}

function outputVacant_(grid, rect, occupied) {
  if (rect.row < 1 || rect.column < 1 || rect.row + rect.height - 1 > EXAM.outputRows ||
      rect.column + rect.width - 1 > EXAM.outputColumns) return false;
  if (occupied.some(function (r) { return outputOverlap_(rect, r); })) return false;
  if (grid.merges.some(function (m) {
    return outputOverlap_(rect, outputMergeRect_(m));
  })) return false;
  for (let row = rect.row - 1; row < rect.row - 1 + rect.height; row++) {
    const values = (grid.rows[row] && grid.rows[row].values) || [];
    for (let col = rect.column - 1; col < rect.column - 1 + rect.width; col++) {
      if (!outputCellBlank_(values[col])) return false;
    }
  }
  return true;
}

/** 入力された左上セルを配置座標に変換する。 */
function outputCellPosition_(cell) {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(String(cell || '').trim().toUpperCase());
  if (!match) throw new Error('配置位置は C12 のように左上のセルを入力してください。');
  let column = 0;
  for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (!Number.isSafeInteger(column) || !Number.isSafeInteger(row) ||
      column > EXAM.outputColumns || row > EXAM.outputRows) {
    throw new Error('配置位置は A1:BT51 の範囲内で指定してください。');
  }
  return { row: row, column: column };
}

function outputFindSpace_(grid, part, occupied, cachedCellMap) {
  const height = part.height, width = part.width;
  const cellMap = cachedCellMap || outputSnapshot_(grid).occupied;
  for (const col of [EXAM.leftColumn, EXAM.rightColumn]) {
    for (let row = EXAM.firstRow; row + height - 1 <= EXAM.outputRows; row++) {
      const rect = { row: row, column: col, width: width, height: height };
      if (col + width - 1 > EXAM.outputColumns ||
          occupied.some(function (existing) { return outputOverlap_(rect, existing); })) continue;
      let vacant = true;
      for (let r = row - 1; r < row - 1 + height; r++) {
        if (cellMap[r].slice(col - 1, col - 1 + width).includes('1')) {
          vacant = false;
          break;
        }
      }
      if (vacant) return rect;
    }
  }
  throw new Error('新しいパーツを配置できる空きがありません。出力シートを確認してください。');
}

/** 初回の得点枠と氏名欄は、それぞれ AZ44 と AZ48 を左上に置く。 */
function outputFindBasicBottomRight_(grid, part, occupied) {
  const rect = {
    row: part.type === '年組氏名欄' ? 48 : 44,
    column: 52,
    width: part.width,
    height: part.height,
  };
  if (rect.row < EXAM.firstRow || !outputVacant_(grid, rect, occupied)) {
    throw new Error('右下の配置場所 ' + outputA1_(rect) + ' に既存の内容があるか、用紙内に収まりません。');
  }
  return rect;
}

function outputSnapshot_(grid) {
  const occupied = Array.from({ length: EXAM.outputRows }, function (_, row) {
    const values = (grid.rows[row] && grid.rows[row].values) || [];
    return Array.from({ length: EXAM.outputColumns }, function (_, col) {
      return outputCellBlank_(values[col]) ? '0' : '1';
    });
  });
  grid.merges.forEach(function (merge) {
    const rect = outputMergeRect_(merge);
    for (let row = Math.max(1, rect.row); row <= Math.min(EXAM.outputRows, rect.row + rect.height - 1); row++) {
      for (let col = Math.max(1, rect.column); col <= Math.min(EXAM.outputColumns, rect.column + rect.width - 1); col++) {
        occupied[row - 1][col - 1] = '1';
      }
    }
  });
  return { scannedAt: new Date().toISOString(), occupied: occupied.map(function (row) { return row.join(''); }) };
}

/** 候補計算用の一時キャッシュ。確定時は必ず実セルを再確認する。 */
function outputCache_(sheet, grid) {
  try {
    const cache = CacheService.getDocumentCache();
    if (cache) cache.put('placement_occupancy_' + sheet.getSheetId(), JSON.stringify(outputSnapshot_(grid).occupied), 120);
  } catch (_) { /* キャッシュ不可時は通常の読み取りに戻す。 */ }
}

function outputCachedGrid_(sheet) {
  try {
    const cache = CacheService.getDocumentCache();
    const raw = cache && cache.get('placement_occupancy_' + sheet.getSheetId());
    if (!raw) return null;
    const occupied = JSON.parse(raw);
    if (!Array.isArray(occupied) || occupied.length !== EXAM.outputRows || occupied.some(function (row) {
      return typeof row !== 'string' || row.length !== EXAM.outputColumns || /[^01]/.test(row);
    })) return null;
    return {rows: occupied.map(function (row) { return {values: Array.from(row, function (cell) {
      return cell === '1' ? {userEnteredValue: {boolValue: true}} : {};
    })}; }), merges: []};
  } catch (_) { return null; }
}

function warmPlacementCache() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAM.outputSheetName);
  if (!sheet || outputCachedGrid_(sheet)) return;
  outputCache_(sheet, outputReadGrid_(sheet));
}

function refreshOutputSnapshot() {
  return withDraftLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EXAM.outputSheetName);
    if (!sheet) throw new Error('「出力」シートがありません。setupSheets を実行してください。');
    const grid = outputReadGrid_(sheet);
    const draft = readDraft_();
    draft.outputSnapshot = outputSnapshot_(grid);
    writeDraft_(draft);
    return { scannedAt: draft.outputSnapshot.scannedAt };
  });
}

function outputPartContent_(part) {
  const copy = JSON.parse(JSON.stringify(part));
  delete copy.outputState;
  delete copy.order;
  delete copy.gap;
  delete copy.sheetRow;
  // 大問の高さは小問から導出されるため、記録時期による未設定や古い値は差分にしない。
  if (copy.kind === 'question') delete copy.height;
  return JSON.stringify(outputCanonical_(copy));
}

function outputPlan_(previous, next, grid) {
  const oldById = new Map(previous.parts.map(function (part) { return [part.id, part]; }));
  const nextById = new Map(next.parts.map(function (part) { return [part.id, part]; }));
  const existingBottomTypes = new Set(previous.parts.filter(function (part) {
    return part.kind === 'basic' && ['年組氏名欄', '得点枠'].includes(part.type);
  }).map(function (part) { return part.type; }));
  const removals = [], additions = [], occupied = [];
  let cellMap;
  previous.parts.forEach(function (old) {
    const rect = outputRect_(old);
    const current = nextById.get(old.id);
    const changed = !current || outputPartContent_(old) !== outputPartContent_(current) ||
      old.order !== current.order || old.gap !== current.gap;
    if (!rect || !old.outputState.fingerprint) {
      if (changed) throw new Error('既存パーツは出力上の変更を確認できないため編集できません。新しく追加してください。');
      return;
    }
    if (changed) {
      if (outputFingerprint_(grid, rect) !== old.outputState.fingerprint) {
        throw new Error('パーツ ' + outputA1_(rect) + ' は直接編集されているため変更できません。');
      }
      removals.push({ part: old, rect: rect });
    } else {
      current.outputState = old.outputState;
      occupied.push(rect);
    }
  });
  // 既存パーツの移動・拡張先には、変更対象自身以外のセルも含めて上書きしない。
  next.parts.forEach(function (part) {
    const old = oldById.get(part.id);
    const change = !old || removals.some(function (entry) { return entry.part.id === part.id; });
    if (!change) return;
    const oldRect = old && outputRect_(old);
    const sameGeometry = oldRect && oldRect.width === part.width && oldRect.height === part.height &&
      old.order === part.order && old.gap === part.gap;
    let rect;
    if (part.placementCell) {
      const position = outputCellPosition_(part.placementCell);
      rect = { row: position.row, column: position.column, width: part.width, height: part.height };
      if (!outputVacant_(grid, rect, occupied)) {
        throw new Error('大問' + part.number + 'の指定範囲 ' + outputA1_(rect) + ' は用紙外か、既存の内容と重なっています。');
      }
    } else if (!old && part.kind === 'basic' && ['年組氏名欄', '得点枠'].includes(part.type) &&
        !existingBottomTypes.has(part.type)) {
      // AZ から用紙右端 BT までは 21 列。はみ出す幅は用紙内に収める。
      part.width = Math.min(part.width, EXAM.outputColumns - 52 + 1);
      rect = outputFindBasicBottomRight_(grid, part, occupied);
      existingBottomTypes.add(part.type);
    } else if (sameGeometry) {
      rect = oldRect;
    } else {
      if (!cellMap) cellMap = outputSnapshot_(grid).occupied;
      rect = outputFindSpace_(grid, part, occupied, cellMap);
    }
    occupied.push(rect);
    additions.push({ part: part, rect: rect });
  });
  return { removals: removals, additions: additions };
}

function outputPlanSignature_(plan) {
  return JSON.stringify(plan.additions.map(function (entry) {
    return [entry.part.kind, entry.part.kind === 'question' ? entry.part.number : entry.part.type, entry.rect];
  }));
}

/** 変更したセル範囲だけを、更新前に読み取った状態に戻す。 */
function restoreOutputRanges_(ss, sheet, grid, rects) {
  const requests = [];
  const sheetId = sheet.getSheetId();
  const merges = [];
  rects.forEach(function (r) {
    const range = { sheetId: sheetId, startRowIndex: r.row - 1,
      endRowIndex: r.row - 1 + r.height, startColumnIndex: r.column - 1,
      endColumnIndex: r.column - 1 + r.width };
    requests.push({ unmergeCells: { range: range } });
    const rows = [];
    for (let row = r.row - 1; row < r.row - 1 + r.height; row++) {
      const source = (grid.rows[row] && grid.rows[row].values) || [];
      const values = [];
      for (let col = r.column - 1; col < r.column - 1 + r.width; col++) {
        const cell = Object.assign({}, source[col] || {});
        delete cell.effectiveValue;
        values.push(cell);
      }
      rows.push({ values: values });
    }
    requests.push({ updateCells: { start: { sheetId: sheetId, rowIndex: r.row - 1,
      columnIndex: r.column - 1 }, rows: rows,
      fields: 'userEnteredValue,userEnteredFormat,note,dataValidation,textFormatRuns' } });
    grid.merges.forEach(function (merge) {
      const original = outputMergeRect_(merge);
      if (!outputOverlap_(r, original)) return;
      if (original.row < r.row || original.column < r.column ||
          original.row + original.height > r.row + r.height ||
          original.column + original.width > r.column + r.width) {
        throw new Error('変更範囲をまたぐ結合セルがあるため復旧できません。');
      }
      const key = outputA1_(original);
      if (!merges.some(function (item) { return item.key === key; })) {
        merges.push({ key: key, rect: original });
      }
    });
  });
  merges.forEach(function (item) {
    const r = item.rect;
    requests.push({ mergeCells: { range: { sheetId: sheetId, startRowIndex: r.row - 1,
      endRowIndex: r.row - 1 + r.height, startColumnIndex: r.column - 1,
      endColumnIndex: r.column - 1 + r.width }, mergeType: 'MERGE_ALL' } });
  });
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
}

/** 既存の描画関数を Sheets API の一括リクエストに変換する。 */
function outputDrawRequests_(sheetId, plan) {
  const requests = [];
  function rangeOf(row, column, height, width) {
    return { sheetId: sheetId, startRowIndex: row - 1, endRowIndex: row - 1 + height,
      startColumnIndex: column - 1, endColumnIndex: column - 1 + width };
  }
  plan.removals.forEach(function (entry) {
    const r = entry.rect, range = rangeOf(r.row, r.column, r.height, r.width);
    requests.push({ unmergeCells: { range: range } });
    requests.push({ updateCells: { range: range, fields: 'userEnteredValue,userEnteredFormat,note,dataValidation,textFormatRuns' } });
    requests.push({ repeatCell: { range: range, cell: { userEnteredFormat: { backgroundColor: {red:1,green:1,blue:1} } },
      fields: 'userEnteredFormat.backgroundColor' } });
  });
  const ranges = [];
  const sheet = { getRange: function (row, column, height, width) {
    const record = { range: rangeOf(row, column, height, width), format: {}, merge: false, border: false };
    ranges.push(record);
    const api = {};
    api.merge = function () { record.merge = true; return api; };
    api.setNumberFormat = function () { record.format.numberFormat = { type: 'TEXT' }; return api; };
    api.setValue = function (value) { record.value = String(value); return api; };
    api.setFontFamily = function (value) { record.format.textFormat = record.format.textFormat || {}; record.format.textFormat.fontFamily = value; return api; };
    api.setFontSize = function (value) { record.format.textFormat = record.format.textFormat || {}; record.format.textFormat.fontSize = value; return api; };
    api.setFontWeight = function (value) { record.format.textFormat = record.format.textFormat || {}; record.format.textFormat.bold = value === 'bold'; return api; };
    api.setHorizontalAlignment = function (value) { record.format.horizontalAlignment = value.toUpperCase(); return api; };
    api.setVerticalAlignment = function (value) { record.format.verticalAlignment = value === 'middle' ? 'MIDDLE' : value.toUpperCase(); return api; };
    api.setBorder = function () { record.border = true; return api; };
    return api;
  } };
  plan.additions.forEach(function (entry) {
    const r = entry.rect, placement = {part: entry.part, row: r.row, column: r.column};
    if (entry.part.kind === 'basic') drawBasic_(sheet, placement);
    else drawQuestion_(sheet, placement);
  });
  ranges.forEach(function (item) {
    if (item.merge) requests.push({mergeCells: {range:item.range, mergeType:'MERGE_ALL'}});
    const fields = Object.keys(item.format).map(function (key) { return 'userEnteredFormat.' + key; }).join(',');
    if (fields) requests.push({repeatCell: {range:item.range, cell:{userEnteredFormat:item.format}, fields:fields}});
    if (item.value !== undefined) requests.push({updateCells: {start: {sheetId:sheetId,
      rowIndex:item.range.startRowIndex, columnIndex:item.range.startColumnIndex},
      rows:[{values:[{userEnteredValue:{stringValue:item.value}}]}], fields:'userEnteredValue'}});
    if (item.border) {
      const line = {style:'SOLID'};
      requests.push({updateBorders: {range:item.range, top:line, bottom:line, left:line, right:line}});
    }
  });
  return requests;
}

/** 変更のない既存パーツと、新規大問のみの場合に確認済みの範囲を再利用する。 */
function outputConfirmedAdditions_(previous, next, signature) {
  if (!signature) return null;
  const oldById = new Map(previous.parts.map(function (part) { return [part.id, part]; }));
  const nextById = new Map(next.parts.map(function (part) { return [part.id, part]; }));
  if (previous.parts.some(function (old) {
    const part = nextById.get(old.id);
    return !part || outputPartContent_(old) !== outputPartContent_(part) ||
      old.order !== part.order || old.gap !== part.gap;
  })) return null;
  const added = next.parts.filter(function (part) { return !oldById.has(part.id); });
  if (!added.length || added.some(function (part) { return part.kind !== 'question'; })) return null;
  let entries;
  try { entries = JSON.parse(signature); } catch (_) { throw new Error('配置候補を再確認してください。'); }
  if (!Array.isArray(entries) || entries.length !== added.length) throw new Error('配置候補を再確認してください。');
  const occupied = previous.parts.map(outputRect_).filter(Boolean);
  const plan = {removals: [], additions: []};
  added.forEach(function (part, index) {
    const entry = entries[index], rect = entry && entry[2];
    if (!Array.isArray(entry) || entry[0] !== 'question' || entry[1] !== part.number || !rect ||
        !Number.isInteger(rect.row) || !Number.isInteger(rect.column) ||
        rect.width !== part.width || rect.height !== part.height ||
        rect.row < 1 || rect.column < 1 || rect.row + rect.height - 1 > EXAM.outputRows ||
        rect.column + rect.width - 1 > EXAM.outputColumns ||
        occupied.some(function (other) { return outputOverlap_(rect, other); })) {
      throw new Error('配置候補が変更されています。もう一度確認してください。');
    }
    if (part.placementCell) {
      const position = outputCellPosition_(part.placementCell);
      if (position.row !== rect.row || position.column !== rect.column) throw new Error('配置候補を再確認してください。');
    }
    occupied.push(rect);
    plan.additions.push({part: part, rect: rect});
  });
  if (outputPlanSignature_(plan) !== signature) throw new Error('配置候補を再確認してください。');
  return plan;
}

/** 同じ位置・寸法の大問だけを置き換える場合は、その範囲だけを読む。 */
function outputConfirmedReplacements_(previous, next, signature, preview) {
  if ((!signature && !preview) || previous.parts.length !== next.parts.length) return null;
  const oldById = new Map(previous.parts.map(function (part) { return [part.id, part]; }));
  const changed = [];
  for (const part of next.parts) {
    const old = oldById.get(part.id);
    if (!old) return null;
    if (outputPartContent_(old) === outputPartContent_(part) && old.order === part.order && old.gap === part.gap) continue;
    const rect = outputRect_(old);
    if (part.kind !== 'question' || !rect || !old.outputState.fingerprint || part.placementCell ||
        old.order !== part.order || old.gap !== part.gap || rect.width !== part.width || rect.height !== part.height ||
        rect.row < 1 || rect.column < 1 || rect.row + rect.height - 1 > EXAM.outputRows ||
        rect.column + rect.width - 1 > EXAM.outputColumns) return null;
    changed.push({part:part, rect:rect});
  }
  if (!changed.length) return null;
  const plan = {removals:[], additions:changed};
  if (signature && outputPlanSignature_(plan) !== signature) return null;
  if (changed.some(function (entry) {
    return previous.parts.some(function (old) {
      const rect = outputRect_(old);
      return old.id !== entry.part.id && rect && outputOverlap_(entry.rect, rect);
    });
  })) return null;
  return plan;
}

/** 他のパーツを変えない削除では、削除対象だけを照合する。 */
function outputOnlyRemovals_(previous, next) {
  const nextById = new Map(next.parts.map(function (part) { return [part.id, part]; }));
  const oldById = new Map(previous.parts.map(function (part) { return [part.id, part]; }));
  if (next.parts.some(function (part) {
    const old = oldById.get(part.id);
    return !old || outputPartContent_(old) !== outputPartContent_(part) || old.order !== part.order || old.gap !== part.gap;
  })) return null;
  const removed = previous.parts.filter(function (part) { return !nextById.has(part.id); });
  if (!removed.length) return null;
  const rects = removed.map(outputRect_);
  if (removed.some(function (part, i) {
    const r = rects[i];
    return !r || !part.outputState.fingerprint || r.row < 1 || r.column < 1 ||
      r.row + r.height - 1 > EXAM.outputRows || r.column + r.width - 1 > EXAM.outputColumns ||
      previous.parts.some(function (other) {
        const otherRect = outputRect_(other);
        return other.id !== part.id && otherRect && outputOverlap_(r, otherRect);
      });
  })) throw new Error('削除対象の範囲を安全に確認できません。');
  return rects;
}

function updateOutputParts_(previous, next, expectedSignature, selectedRemoval) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const output = withOperationLog_('placement_get_sheet', function () {
    return ss.getSheetByName(EXAM.outputSheetName);
  });
  if (!output) throw new Error('「出力」シートがありません。setupSheets を実行してください。');
  const additions = outputConfirmedAdditions_(previous, next, expectedSignature);
  const replacements = additions ? null : outputConfirmedReplacements_(previous, next, expectedSignature);
  const confirmed = additions || replacements;
  const removalRects = selectedRemoval ? [selectedRemoval.rect] : (confirmed ? null : outputOnlyRemovals_(previous, next));
  const originalRemoval = selectedRemoval && outputRect_(previous.parts.find(function (part) {
    return part.id === selectedRemoval.partId;
  }) || {});
  const rects = confirmed ? confirmed.additions.map(function (entry) { return entry.rect; }) :
    originalRemoval && !outputSamePosition_(originalRemoval, selectedRemoval.rect) ?
      removalRects.concat([originalRemoval]) : removalRects;
  const before = rects ? outputReadRanges_(output, rects) : outputReadGrid_(output);
  if (removalRects && before.merges.some(function (merge) {
    const merged = outputMergeRect_(merge);
    return removalRects.some(function (r) { return outputOverlap_(r, merged) && !outputContains_(r, merged); });
  })) throw new Error('削除範囲をまたぐ結合セルがあります。');
  const plan = selectedRemoval ? outputSelectedRemovalPlan_(previous, next, selectedRemoval, before) :
    (replacements ? outputPlan_(previous, next, before) : (confirmed || outputPlan_(previous, next, before)));
  if (additions && rects.some(function (rect) { return !outputVacant_(before, rect, []); })) {
    throw new Error('配置先が変更されました。配置候補を再確認してください。');
  }
  // 配置指定は今回の操作だけに使い、パーツの内容として保存しない。
  next.parts.forEach(function (part) { delete part.placementCell; });
  plan.additions.forEach(function (entry) {
    logEvent_('info', 'update_output', 'candidate', { range: outputA1_(entry.rect) });
  });
  if (expectedSignature && outputPlanSignature_(plan) !== expectedSignature) {
    throw new Error('出力の状態が変わりました。配置候補を再確認してください。');
  }
  if (!plan.removals.length && !plan.additions.length) {
    writeDraft_(next);
    return;
  }
  let outputTouched = false, draftTouched = false;
  try {
    // 範囲指定の更新では直前の読み取り結果を使用する。
    if (!rects && JSON.stringify(outputCanonical_(outputReadGrid_(output))) !==
        JSON.stringify(outputCanonical_(before))) {
      throw new Error('出力シートが更新中に変更されました。もう一度操作してください。');
    }
    outputTouched = true;
    const drawResult = withOperationLog_('placement_draw', function () {
      const request = {requests: outputDrawRequests_(output.getSheetId(), plan)};
      if (confirmed) {
        const prefix = "'" + output.getName().replace(/'/g, "''") + "'!";
        request.includeSpreadsheetInResponse = true;
        request.responseIncludeGridData = true;
        request.responseRanges = rects.map(function (r) { return prefix + outputA1_(r); });
        return Sheets.Spreadsheets.batchUpdate(request, ss.getId(), {fields:
          'updatedSpreadsheet(sheets(properties(sheetId,gridProperties(rowCount,columnCount)),data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,userEnteredFormat,note,dataValidation,textFormatRuns))),merges))'});
      }
      return Sheets.Spreadsheets.batchUpdate(request, ss.getId());
    });
    const after = removalRects ? null : (confirmed ? outputGridFromRanges_(output, drawResult && drawResult.updatedSpreadsheet, rects) : outputReadGrid_(output));
    plan.additions.forEach(function (entry) {
      entry.part.outputState = { rect: entry.rect, fingerprint: outputFingerprint_(after, entry.rect) };
    });
    if (confirmed || removalRects) delete next.outputSnapshot;
    else next.outputSnapshot = outputSnapshot_(after);
    draftTouched = true;
    withOperationLog_('placement_save', function () { writeDraft_(next); });
    SpreadsheetApp.flush();
  } catch (error) {
    logEvent_('warn', 'update_output', 'restore_started');
    try {
      if (outputTouched) {
        const rects = [];
        plan.removals.concat(plan.additions).forEach(function (entry) {
          if (!rects.some(function (r) { return outputA1_(r) === outputA1_(entry.rect); })) rects.push(entry.rect);
        });
        restoreOutputRanges_(ss, output, before, rects);
      }
      if (draftTouched) writeDraft_(previous);
      SpreadsheetApp.flush();
      logEvent_('info', 'update_output', 'restore_complete');
    } catch (restoreError) {
      logEvent_('error', 'update_output', 'restore_failed');
      throw new Error('出力更新と復旧に失敗しました。出力シートの変更履歴を確認してください。');
    }
    throw error;
  }
}
