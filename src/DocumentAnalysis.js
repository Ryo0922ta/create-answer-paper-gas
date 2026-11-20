/** 文書全体の解析。候補は出力へ配置する前にパネルで確認する。 */
const DOCUMENT_ANALYSIS = {
  circle: Array.from('①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟'),
  kana: Array.from('アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン'),
};

function analyzeDocumentText(text, fallbackNumber) {
  return withOperationLog_('analyze_document', function () {
    return analyzeDocumentTextLogged_(text, fallbackNumber);
  }, function (result) { return { questionCount: result.length }; });
}

function analyzeDocumentTextLogged_(text, fallbackNumber) {
  const source = String(text || '').replace(/\r\n?/g, '\n');
  if (!source.trim()) throw new Error('問題文を入力してください。');
  const headers = [];
  const heading = /^[ \t　]*(?:【[ \t　]*大問[ \t　]*([0-9０-９]{1,2})[ \t　]*】|大問[ \t　]*([0-9０-９]{1,2}))(.*)$/gm;
  let m;
  while ((m = heading.exec(source))) {
    const number = String(m[1] || m[2]).normalize('NFKC');
    const n = Number(number);
    if (n < 1 || n > 12) throw new Error('大問番号は1～12にしてください: ' + number);
    headers.push({at:m.index,bodyAt:heading.lastIndex,number,heading:String(m[3] || '').trim()});
  }
  let blocks;
  if (headers.length) {
    const used = new Set();
    blocks = headers.map((h,i) => {
      if (used.has(h.number)) throw new Error('大問' + h.number + 'が重複しています。');
      used.add(h.number);
      return {number:h.number, heading:h.heading, detectedBy:'行頭の大問見出し',
        body:source.slice(h.bodyAt, headers[i+1] ? headers[i+1].at : source.length)};
    });
  } else {
    const sections = parseQuestionText(source);
    const number = sections[0] && sections[0].questionNumber;
    const selected = number || String(fallbackNumber == null ? '' : fallbackNumber).trim().normalize('NFKC');
    if (!/^(?:[1-9]|1[0-2])$/.test(selected)) {
      throw new Error('大問の見出しがありません。「見出しがない場合の大問番号」に1～12を入力してください。');
    }
    blocks = [{number:selected,heading:'',body:source,
      detectedBy:number?'行頭の枝番（'+number+'-1. など）':'入力された大問番号'}];
  }
  return blocks.map(block => analyzeMajorBlock_(block));
}

/** 設定を変更せず、出力と同じ列分割で解析候補の枠を返す。 */
function previewAnalyzedDocument(questions) {
  if (!Array.isArray(questions)) throw new Error('解析候補を確認してください。');
  return questions.map(function (question, index) {
    const width = Number(question.width);
    const order = Number(question.order == null ? index + 1 : question.order);
    const result = { number: String(question.number || ''),
      order: Number.isFinite(order) ? order : 0,
      width: Number.isFinite(width) ? width : 0,
      index: index, groups: [] };
    if (!Number.isInteger(width) || width < 6 || width > 33 ||
        !Number.isInteger(order) || order < 1) {
      result.error = '大問の順番または幅を確認してください。';
      return result;
    }
    result.groups = (Array.isArray(question.groups) ? question.groups : []).map(function (group) {
      const row = { path: String(group.path || ''), pattern: String(group.pattern || ''),
        frameCount: Number(group.frameCount), height: Number(group.height) };
      try {
        if (!Number.isInteger(row.frameCount) || row.frameCount < 1 || row.frameCount > 30 ||
            !Number.isInteger(row.height) || row.height < 1 || row.height > 10) {
          throw new Error('小問の枠数または高さを確認してください。');
        }
        const labels = Array.isArray(group.labels) ? group.labels.map(String) : [];
        if (row.pattern === '小問複合枠' && row.frameCount !== labels.length) {
          throw new Error('ラベル数と枠数を一致させてください。');
        }
        row.boxes = answerRowLayout_(width, { rowLabel: String(group.rowLabel || ''),
          pattern: row.pattern, frameCount: row.frameCount, labels: labels,
          weights: Array.isArray(group.weights) ? group.weights.map(Number) : [] });
      } catch (error) {
        row.error = String(error.message || error);
      }
      return row;
    });
    result.totalFrames = result.groups.every(function (group) { return !group.error; })
      ? result.groups.reduce(function (sum, group) { return sum + group.frameCount; }, 0) : null;
    return result;
  }).sort(function (a, b) { return a.order - b.order || a.index - b.index; });
}

function analyzeMajorBlock_(block) {
  const body = block.body;
  const rubricMatch = /[［\[【]\s*(知[・･\s]*技|思[・･\s]*判[・･\s]*表)\s*[］\]】]/.exec(block.heading + '\n' + body);
  const pointsMatch = /(?:合計|計|配点)\s*([0-9０-９]+)\s*点/.exec(block.heading + '\n' + body);
  const titleMatch = /(?:テーマ|題名)\s*[:：]\s*([^\n]+)/.exec(block.heading + '\n' + body);
  const title = (titleMatch ? titleMatch[1] : block.heading)
    .replace(/[［\[【].+?[］\]】]/g,'')
    .replace(/(?:合計|計|配点)\s*[0-9０-９]+\s*点/g,'').trim();
  const starts = [];
  const section = /^[ \t　]*([0-9０-９]+)[-－–]([0-9０-９]+)[.．、][ \t　]*/gm;
  let m;
  while ((m = section.exec(body))) {
    if (m[1].normalize('NFKC') === block.number) {
      starts.push({at:m.index,bodyAt:section.lastIndex,label:m[2].normalize('NFKC')});
    }
  }
  if (!starts.length) {
    const item = /^[ \t　]*(?:[（(]([0-9０-９]+)[）)]|問[ \t　]*([0-9０-９]+))[ \t　]*/gm;
    while ((m = item.exec(body))) {
      starts.push({at:m.index,bodyAt:item.lastIndex,label:(m[1] || m[2]).normalize('NFKC'),
        prefix:m[1] ? 'paren' : 'ask'});
    }
  }
  if (!starts.length) {
    const item = /^[ \t　]*[（(]([a-zA-ZＡ-Ｚａ-ｚぁ-んァ-ヶｦ-ﾟ]+)[）)][ \t　]*/gm;
    while ((m = item.exec(body))) {
      if (isKanaMarker_(m[1])) continue;
      starts.push({at:m.index,bodyAt:item.lastIndex,label:m[1].normalize('NFKC'),prefix:'paren'});
    }
  }
  const groups = [];
  if (starts.length) {
    starts.forEach((start,i) => {
      const text = body.slice(start.bodyAt, starts[i+1] ? starts[i+1].at : body.length);
      const path = block.number + '-' + (start.prefix === 'paren' ? '(' + start.label + ')' :
        start.prefix === 'ask' ? '問' + start.label : start.label);
      groups.push(...analyzeSubgroup_(block.number,start.label,path,text,start.prefix === 'paren'));
    });
  } else {
    groups.push(...analyzeSubgroup_(block.number,'1',block.number+'-1',body,false));
    groups.forEach(g => { g.status='要確認'; g.evidence+='／小問見出しがないためグループを仮作成'; });
  }
  return {
    number:block.number, detectedBy:block.detectedBy, title:title, rubric:rubricMatch ? rubricMatch[1] : '',
    points:pointsMatch ? pointsMatch[1] : '', groups,
    totalFrames:groups.reduce((n,g)=>n+g.frameCount,0),
    expectedFrames:'', order:Number(block.number), width:33, gap:1
  };
}

function analyzeSubgroup_(number, label, path, text, hasParentMarker) {
  // ①の下に(A)(B)、②の下に(ｱ)～(ｷ)がある場合は、①と②を別の末端グループにする。
  const sub = [];
  const circleLine = /^[ \t　]*([①-⑳㉑-㉟])(?=[ \t　（(、．.])/gm;
  let m;
  while ((m = circleLine.exec(text))) sub.push({at:m.index,bodyAt:circleLine.lastIndex,label:m[1]});
  if (sub.length > 1) {
    const chunks = sub.map((s,i) => text.slice(s.bodyAt,sub[i+1] ? sub[i+1].at : text.length));
    if (chunks.some(chunk => /[（(][A-Za-zＡ-Ｚａ-ｚｦ-ﾟァ-ヶ][）)]|[ｦ-ﾟァ-ヶ]\s*[〜～~－-]\s*[ｦ-ﾟァ-ヶ]/.test(chunk))) {
      return sub.map((s,i) => analyzeLeaf_(number,label,path+s.label,chunks[i],s.label,false));
    }
  }
  return [analyzeLeaf_(number,label,path,text,'',hasParentMarker)];
}

function analyzeLeaf_(number, rowLabel, path, text, ownCircle, hasParentMarker) {
  // 選択肢欄の記号は解答欄の候補に含めない。
  const answerText = text.split(/(?:^|\n)[ \t　]*(?:【選択肢】|\[選択肢\])/)[0];
  const actualLines = answerText.split('\n').map(line => {
    const m = /^[ \t　]*([①-⑳㉑-㉟]|[（(][0-9０-９a-zA-ZＡ-Ｚａ-ｚｦ-ﾟァ-ヶ]+[）)]|\[[0-9A-Za-z]+\])/.exec(line);
    return m ? m[1] : '';
  }).filter(Boolean).filter(x=>x!==ownCircle);
  const spans = expandAnswerRanges_(answerText);
  let labels, evidence, status='自動';
  const candidates=actualLines.filter(x=>!hasParentMarker ||
    x!== '('+rowLabel+')' && x!== '（'+rowLabel+'）');
  if (candidates.length>=2 && !ownCircle) {
    labels=uniqueSymbols_(candidates);
    evidence='行頭の解答記号';
  } else if (spans.length) {
    const primary = spans.find(x=>DOCUMENT_ANALYSIS.circle.includes(x[0])) || spans[0];
    labels=primary;
    evidence='連続記号を展開';
    if (spans.length>1 && /組み合わせ|組合せ/.test(answerText)) {
      evidence+='／組み合わせ：1枠に2項目';
    } else if (spans.length>1) status='要確認';
  } else {
    labels=uniqueSymbols_(collectAnswerLabels_(answerText,hasParentMarker?'('+rowLabel+')':'',false)
      .filter(x=>x!==ownCircle));
    evidence=labels.length ? '本文の記号' : '記号なし：記述欄';
  }
  const compound = spans.length>1 && /組み合わせ|組合せ/.test(answerText) &&
    labels.length === spans.find(x=>DOCUMENT_ANALYSIS.circle.includes(x[0]))?.length;
  if (/【選択肢】|\[選択肢\]/.test(text) && !labels.length) status='要確認';
  let frameCount=labels.length || 1;
  const countPhrase=/(?:([2-9２-９])\s*つ|二つ|２つ)の[^。\n]{0,32}?(?:答え|挙げ|書き)/.exec(answerText) ||
    /(?:([2-9２-９])\s*つ|二つ|２つ)の[^。\n]{0,18}/.exec(answerText);
  if (!labels.length && countPhrase) {
    frameCount=countPhrase[1] ? Number(countPhrase[1].normalize('NFKC')) : 2;
    evidence='「'+countPhrase[0]+'」から'+frameCount+'枠を推定';status='要確認';
  }
  const xy=/\bX\s*と\s*Y\b/.test(answerText);
  if (!labels.length && xy) {labels=['X','Y'];frameCount=2;evidence='XとYの2つの空欄';}
  if (labels.length>15) status='要確認'; // 33列では狭くなることが多い
  return {path,rowLabel,labels,weights:labels.map(()=>1),frameCount,expectedFrames:'',itemsPerFrame:compound?2:1,
    pattern:labels.length?'小問複合枠':'記述欄',height:2,
    evidence,status,sourceText:text.trim().slice(0,180)};
}

function uniqueSymbols_(symbols) {
  const seen=new Set();
  return symbols.filter(x=>{if(seen.has(x))return false;seen.add(x);return true;});
}

function expandAnswerRanges_(text) {
  const ranges=[];
  const re=/([①-⑳㉑-㉟]|[（(][a-zA-Zａ-ｚＡ-Ｚｦ-ﾟァ-ヶ][）)]|[ｦ-ﾟァ-ヶ])\s*[〜～~－-]\s*([①-⑳㉑-㉟]|[（(][a-zA-Zａ-ｚＡ-Ｚｦ-ﾟァ-ヶ][）)]|[ｦ-ﾟァ-ヶ])/g;
  let m;
  // 「（ア〜ケ）」「（A〜I）」は範囲全体を括弧で囲む表記。
  const enclosing=/[（(]([a-zA-Zｦ-ﾟァ-ヶ])\s*[〜～~－-]\s*([a-zA-Zｦ-ﾟァ-ヶ])[）)]/g;
  let enclosed;
  while ((enclosed=enclosing.exec(text))) {
    const a=rangeSequence_('('+enclosed[1]+')'), b=rangeSequence_('('+enclosed[2]+')');
    if (!a || !b || a.kind!==b.kind) continue;
    const first=a.sequence.indexOf(a.value),last=a.sequence.indexOf(b.value);
    if(first>=0 && last>=first && last-first<=35) ranges.push(a.sequence.slice(first,last+1).map(v=>'('+v+')'));
  }
  while ((m=re.exec(text))) {
    const a=rangeSequence_(m[1]), b=rangeSequence_(m[2]);
    if (!a || !b || a.kind!==b.kind || a.wrapped!==b.wrapped) continue;
    const first=a.sequence.indexOf(a.value), last=a.sequence.indexOf(b.value);
    if (first<0 || last<first || last-first>35) continue;
    ranges.push(a.sequence.slice(first,last+1).map(v=>a.wrapped?'('+v+')':v));
  }
  return ranges;
}

function rangeSequence_(token) {
  if (DOCUMENT_ANALYSIS.circle.includes(token)) return {kind:'circle',value:token,
    sequence:DOCUMENT_ANALYSIS.circle,wrapped:false};
  const wrapped=/^[（(].+[）)]$/.test(token);
  const inner=(wrapped?token.slice(1,-1):token).normalize('NFKC');
  if (/^[ア-ン]$/.test(inner)) return {kind:'kana',value:inner,
    sequence:DOCUMENT_ANALYSIS.kana,wrapped};
  if (wrapped && /^[A-Z]$/.test(inner)) return {kind:'upper',value:inner,
    sequence:Array.from('ABCDEFGHIJKLMNOPQRSTUVWXYZ'),wrapped};
  if (wrapped && /^[a-z]$/.test(inner)) return {kind:'lower',value:inner,
    sequence:Array.from('abcdefghijklmnopqrstuvwxyz'),wrapped};
  return null;
}

/** 印刷予定の物理枠数を、描画と同じ小問単位で集計する。 */
function answerFrameReport_(parts, expected) {
  const byQuestion={};
  parts.filter(p=>p.kind==='question').forEach(p=>{
    const rows=p.rows.map(r=>({path:r.path,frames:r.pattern==='小問複合枠'?r.labels.length:r.frameCount,
      itemsPerFrame:r.itemsPerFrame}));
    const total=rows.reduce((sum,r)=>sum+r.frames,0);
    const target=expected && expected[p.number];
    byQuestion[p.number]={total,rows,expected:target===''||target==null?null:Number(target),
      matches:target===''||target==null?null:Number(target)===total};
  });
  return byQuestion;
}

/** 選んだ大問を作業データへ配置する。設定シートは変更しない。 */
function importAnalyzedDocument(payload) {
  return withOperationLog_('place_questions', function () {
    return importAnalyzedDocumentLogged_(payload);
  }, function (result) { return { questionCount: payload && Array.isArray(payload.questions) ? payload.questions.length : 0, conflictCount: result && result.conflict ? result.numbers.length : 0 }; });
}

function importAnalyzedDocumentLogged_(payload) {
  return withDraftLock_(function () {
    const questions = payload && payload.questions;
    if (!Array.isArray(questions) || !questions.length || questions.length > 12) {
      throw new Error('配置する大問を1～12件選んでください。');
    }
    const previous = withOperationLog_('placement_read_draft', function () { return readDraft_(); });
    const draft = JSON.parse(JSON.stringify(previous));
    const seen = new Set();
    const existingByNumber = new Map(previous.parts.filter(function (part) { return part.kind === 'question'; })
      .map(function (part) { return [part.number, part]; }));
    let order = previous.parts.reduce(function (max, part) { return Math.max(max, part.order); }, 0);
    if (payload.positions != null && (typeof payload.positions !== 'object' || Array.isArray(payload.positions))) {
      throw new Error('配置位置を確認してください。');
    }
    const prepared = questions.map(function (question) {
      const number = String(question.number || '').trim();
      if (!/^(?:[1-9]|1[0-2])$/.test(number) || seen.has(number)) {
        throw new Error('大問番号は1～12で重複なく入力してください。');
      }
      seen.add(number);
      const width = requiredInteger_(question.width, '大問' + number + 'の幅', 6, 33);
      if (!Array.isArray(question.groups) || !question.groups.length) {
        throw new Error('大問' + number + 'には小問を1件以上含めてください。');
      }
      const rows = question.groups.map(function (group, index) {
        if (group.status === '要確認' && group.reviewed !== true) {
          throw new Error('大問' + number + ' ' + String(group.path || index + 1) + 'を確認してください。');
        }
        const labels = Array.isArray(group.labels) ? group.labels.map(function (x) { return String(x).trim(); }) : [];
        const weights = Array.isArray(group.weights) ? group.weights.map(Number) : [];
        const pattern = answerPattern_(group.pattern, labels.join(','), index + 1);
        const frameCount = requiredInteger_(group.frameCount, '小問の枠数', 1, 30);
        if (group.expectedFrames !== '' && group.expectedFrames != null &&
            requiredInteger_(group.expectedFrames, '小問の照合枠数', 1, 30) !== frameCount) {
          throw new Error('大問' + number + 'の小問' + (index + 1) + 'の枠数が原文と一致しません。');
        }
        return {
          sheetRow: index + 1, index: index + 1,
          rowLabel: String(group.rowLabel || index + 1),
          labels: labels, weights: weights,
          height: requiredInteger_(group.height, '小問の高さ', 1, 10),
          pattern: pattern, path: String(group.path || number + '-' + (index + 1)),
          frameCount: frameCount,
          itemsPerFrame: requiredInteger_(group.itemsPerFrame, '記入項目数', 1, 10),
          evidence: String(group.evidence || '').slice(0, 250),
        };
      });
      const total = rows.reduce(function (sum, row) { return sum + row.frameCount; }, 0);
      if (question.expectedFrames !== '' && question.expectedFrames != null &&
          requiredInteger_(question.expectedFrames, '大問の照合枠数', 1, 500) !== total) {
        throw new Error('大問' + number + 'の枠数が原文と一致しません。');
      }
      const old = existingByNumber.get(number);
      const rawCell = payload.positions && Object.prototype.hasOwnProperty.call(payload.positions, number) ?
        payload.positions[number] : question.placementCell;
      const placementCell = String(rawCell || '').trim();
      return {
        id: old ? old.id : Utilities.getUuid(), kind: 'question', sheetRow: draft.parts.length + 1,
        order: question.order == null ? (old ? old.order : ++order) : requiredInteger_(question.order, '順番', 1, 999),
        ...(placementCell ? {placementCell: placementCell} : {}),
        number: number, title: String(question.title || ''),
        rubric: String(question.rubric || ''), points: String(question.points || ''),
        width: width, gap: requiredInteger_(question.gap, '後の空き', 0, 10),
        rows: rows, height: 1 + rows.reduce(function (sum, row) { return sum + row.height; }, 0),
      };
    });
    const existing = draft.parts.filter(function (part) {
      return part.kind === 'question' && seen.has(part.number);
    });
    const fingerprint = JSON.stringify(existing);
    if (existing.length && payload.replaceExisting !== true) {
      logEvent_('warn', 'place_questions', 'conflict', { conflictCount: existing.length });
      return { conflict: true, numbers: existing.map(function (part) { return part.number; }), fingerprint: fingerprint };
    }
    if (payload.replaceExisting === true && payload.existingFingerprint !== fingerprint) {
      logEvent_('warn', 'place_questions', 'conflict_changed', { conflictCount: existing.length });
      throw new Error('確認後に出力が変わりました。もう一度確認してください。');
    }
    if (existing.length) {
      logEvent_('info', 'place_questions', 'replace_confirmed', { conflictCount: existing.length });
    }
    draft.parts = draft.parts.filter(function (part) {
      return part.kind !== 'question' || !seen.has(part.number);
    }).concat(prepared);
    draft.placementBatch = prepared.map(function (part) { return part.id; });
    if (payload.previewOnly === true) {
      const output = withOperationLog_('placement_get_sheet', function () {
        return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAM.outputSheetName);
      });
      if (!output) throw new Error('「出力」シートがありません。');
      validateDraft_(draft);
      let plan;
      try {
        const targeted = outputTargetedPlacementRanges_(previous, draft);
        const replacement = targeted ? null : outputConfirmedReplacements_(previous, draft, null, true);
        const ranges = targeted || (replacement && replacement.additions.map(function (entry) { return entry.rect; }));
        const hasRequestedPosition = prepared.some(function (part) { return part.placementCell; });
        let grid = ranges ? outputReadRanges_(output, ranges) : null;
        if (!grid && !existing.length && !hasRequestedPosition) {
          grid = withOperationLog_('placement_get_cache', function () { return outputCachedGrid_(output); });
        }
        if (!grid) {
          grid = outputReadGrid_(output);
          outputCache_(output, grid);
        }
        plan = withOperationLog_('placement_plan', function () { return outputPlan_(previous, draft, grid); });
      } catch (error) {
        if (error.message === '新しいパーツを配置できる空きがありません。出力シートを確認してください。') {
          return { manualRequired: true, numbers: prepared.map(function (part) { return part.number; }) };
        }
        throw error;
      }
      if (!plan.additions.length) return '選んだ大問の内容と配置は出力済みです。';
      plan.additions.forEach(function (entry) {
        logEvent_('info', 'place_questions', 'candidate', { range: outputA1_(entry.rect) });
      });
      return { placementPreview: true, signature: outputPlanSignature_(plan),
        numbers: prepared.map(function (part) { return part.number; }),
        candidates: plan.additions.map(function (entry) {
          return { number: entry.part.number, range: outputA1_(entry.rect) };
        }) };
    }
    if (!payload.placementSignature) throw new Error('配置候補を確認してください。');
    commitDraft_(draft, payload.placementSignature, previous);
    return '大問' + prepared.map(function (part) { return part.number; }).join('、') +
      'を出力に配置しました。';
  });
}
