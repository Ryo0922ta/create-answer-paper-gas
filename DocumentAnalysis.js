/** 文書全体の解析。候補は保存前に必ずパネルで確認する。 */
const DOCUMENT_ANALYSIS = {
  circle: Array.from('①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟'),
  kana: Array.from('アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン'),
};

function analyzeDocumentText(text, fallbackNumber) {
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
  const actualLines = text.split('\n').map(line => {
    const m = /^[ \t　]*([①-⑳㉑-㉟]|[（(][0-9０-９a-zA-ZＡ-Ｚａ-ｚｦ-ﾟァ-ヶ]+[）)]|\[[0-9A-Za-z]+\])/.exec(line);
    return m ? m[1] : '';
  }).filter(Boolean).filter(x=>x!==ownCircle);
  const spans = expandAnswerRanges_(text);
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
    if (spans.length>1 && /組み合わせ|組合せ/.test(text)) {
      evidence+='／組み合わせ：1枠に2項目';
    } else if (spans.length>1) status='要確認';
  } else {
    labels=uniqueSymbols_(collectAnswerLabels_(text,hasParentMarker?'('+rowLabel+')':'',false)
      .filter(x=>x!==ownCircle));
    evidence=labels.length ? '本文の記号' : '記号なし：記述欄';
  }
  const compound = spans.length>1 && /組み合わせ|組合せ/.test(text) &&
    labels.length === spans.find(x=>DOCUMENT_ANALYSIS.circle.includes(x[0]))?.length;
  // 選択肢の記号は回答箇所に加えない。抽出した候補に混ざる場合は確認する。
  if (/【選択肢】|\[選択肢\]/.test(text) && !actualLines.length && !spans.length) status='要確認';
  let frameCount=labels.length || 1;
  const countPhrase=/(?:([2-9２-９])\s*つ|二つ|２つ)の[^。\n]{0,32}?(?:答え|挙げ|書き)/.exec(text) ||
    /(?:([2-9２-９])\s*つ|二つ|２つ)の[^。\n]{0,18}/.exec(text);
  if (!labels.length && countPhrase) {
    frameCount=countPhrase[1] ? Number(countPhrase[1].normalize('NFKC')) : 2;
    evidence='「'+countPhrase[0]+'」から'+frameCount+'枠を推定';status='要確認';
  }
  const xy=/\bX\s*と\s*Y\b/.test(text);
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

/** 確認済みの解析候補を、複数大問まとめて安全に設定へ登録する。 */
function importAnalyzedDocument(payload) {
  const lock=LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新中です。少し待ってください。');
  try { return importAnalyzedDocumentLocked_(payload || {}); }
  finally { lock.releaseLock(); }
}

function importAnalyzedDocumentLocked_(payload) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(EXAM.settingsSheetName);
  if (!sheet) throw new Error('先に設定シートを作成してください。');
  ensureDetailPatternColumn_(sheet);
  const input=payload.questions;
  if (!Array.isArray(input) || !input.length || input.length>12) {
    throw new Error('登録する大問を1～12件選んでください。');
  }
  const used=new Set();
  const prepared=input.map(question=>{
    const number=String(question.number||'').trim();
    if (!/^(?:[1-9]|1[0-2])$/.test(number) || used.has(number)) {
      throw new Error('大問番号は1～12で重複なく入力してください。');
    }
    used.add(number);
    const order=requiredInteger_(question.order,'大問'+number+'の順番',1,999);
    const width=requiredInteger_(question.width,'大問'+number+'の幅',6,33);
    const gap=requiredInteger_(question.gap,'大問'+number+'の後の空き',0,10);
    if (!Array.isArray(question.groups) || !question.groups.length) {
      throw new Error('大問'+number+'には少なくとも1グループが必要です。');
    }
    const details=question.groups.map((g,i)=>{
      if (g.status==='要確認' && g.reviewed!==true) {
        throw new Error('大問'+number+' '+String(g.path||i+1)+'は「要確認」です。内容を確認してチェックしてください。');
      }
      const labels=Array.isArray(g.labels)?g.labels.map(x=>String(x).trim()):[];
      if (labels.some(x=>!x || x.includes(','))) throw new Error('大問'+number+'の小問'+(i+1)+'のラベルが不正です。');
      const frames=requiredInteger_(g.frameCount,'大問'+number+'の小問'+(i+1)+'の枠数',1,30);
      if (g.expectedFrames!=='' && g.expectedFrames!=null &&
          requiredInteger_(g.expectedFrames,'小問の解答用紙の枠数',1,30)!==frames) {
        throw new Error('大問'+number+' '+String(g.path||i+1)+'：文書の'+frames+'枠と解答用紙の'+g.expectedFrames+'枠が一致しません。');
      }
      const pattern=answerPattern_(g.pattern,labels.join(','),EXAM.detailStartRow+i);
      if (pattern==='小問複合枠' && frames!==labels.length) {
        throw new Error('大問'+number+'の小問'+(i+1)+'はラベル数と枠数を一致させてください。');
      }
      if (pattern==='記述欄' && frames>width-2) {
        throw new Error('大問'+number+'の小問'+(i+1)+'は幅に対して枠が多すぎます。');
      }
      const weights=Array.isArray(g.weights) ? g.weights.map(Number) : labels.map(()=>1);
      if (weights.length!==labels.length && labels.length) {
        throw new Error('大問'+number+'の小問'+(i+1)+'はラベルと幅比を同じ数にしてください。');
      }
      if (weights.some(x=>!Number.isFinite(x)||x<=0)) throw new Error('小問の幅比は正数にしてください。');
      if (pattern==='小問複合枠') validateAnswerWidths_({width},{labels,weights,sheetRow:i+1});
      const label=String(g.rowLabel==null?i+1:g.rowLabel).trim();
      const path=String(g.path || number+'-'+(i+1)).trim();
      if (!path || /^[=+@-]/.test(path) || /^[=+@-]/.test(label)) {
        throw new Error('小問番号とグループ経路を確認してください。');
      }
      return [number,i+1,label,labels.join(','),weights.length?weights.join(','):'1',
        requiredInteger_(g.height,'小問の高さ',1,10),
        '文書から追加',pattern,path,frames,
        requiredInteger_(g.itemsPerFrame,'1枠の記入項目数',1,10),
        (g.status==='要確認'?'[確認済] ':'')+String(g.evidence||'').slice(0,250)];
    });
    const total=details.reduce((sum,row)=>sum+row[9],0);
    const expected=question.expectedFrames;
    if (expected!=='' && expected!==null && expected!==undefined &&
      requiredInteger_(expected,'大問'+number+'の照合枠数',1,500)!==total) {
      throw new Error('大問'+number+'：文書の'+total+'枠と指定された解答用紙の'+expected+'枠が一致しません。');
    }
    return {number,order,width,gap,title:String(question.title||''),rubric:String(question.rubric||''),
      points:String(question.points||''),details,total,expected};
  });
  const qRange=sheet.getRange(EXAM.questionStartRow,1,EXAM.questionRowCount,7);
  const dRange=sheet.getRange(EXAM.detailStartRow,10,EXAM.detailRowCount,EXAM.detailColumnCount);
  const titleRange=sheet.getRange(EXAM.questionStartRow,9,EXAM.questionRowCount,1);
  const qBefore=qRange.getValues(),dBefore=dRange.getValues(),titleBefore=titleRange.getValues();
  const fingerprint=documentSettingsFingerprint_(qBefore,dBefore);
  const conflicts=prepared.filter(p=>qBefore.some(r=>r[0]===true&&String(r[2]).trim()===p.number))
    .map(p=>p.number);
  if (conflicts.length && payload.replaceExisting!==true) {
    return {conflict:true,numbers:conflicts,fingerprint};
  }
  if (payload.replaceExisting===true && String(payload.existingFingerprint||'')!==fingerprint) {
    throw new Error('確認後に設定が変更されました。もう一度解析結果を確認してください。');
  }
  const qAfter=qBefore.map(r=>r.slice()), dAfter=dBefore.map(r=>r.slice());
  const titleAfter=titleBefore.map(r=>r.slice());
  prepared.forEach(p=>{
    let qi=qAfter.findIndex(r=>r[0]===true&&String(r[2]).trim()===p.number);
    if (qi<0) qi=qAfter.findIndex(r=>r.every(v=>v===''||v===false));
    if (qi<0) throw new Error('大問表に空き行がありません。');
    qAfter[qi]=[true,p.order,p.number,p.rubric,p.points,p.width,p.gap];
    titleAfter[qi]=[p.title];
    dAfter.forEach((r,i)=>{if(String(r[0]).trim()===p.number)dAfter[i]=Array(EXAM.detailColumnCount).fill('');});
    p.details.forEach(values=>{
      const di=dAfter.findIndex(r=>r.every(v=>v===''||v===false));
      if(di<0) throw new Error('小問行表に空き行がありません。');
      dAfter[di]=values;
    });
  });
  // すべて検査してからまとめて書き、失敗時は全行を戻す。
  try {
    sheet.getRange(EXAM.questionStartRow,3,EXAM.questionRowCount,1).setNumberFormat('@');
    sheet.getRange(EXAM.detailStartRow,10,EXAM.detailRowCount,EXAM.detailColumnCount).setNumberFormat('@');
    qRange.setValues(qAfter);dRange.setValues(dAfter);titleRange.setValues(titleAfter);
    SpreadsheetApp.flush();
    const parts=readParts_(sheet);
    calculatePlacements_(parts);
    const targets={};prepared.forEach(p=>{targets[p.number]=p.expected;});
    const report=answerFrameReport_(parts,targets);
    prepared.forEach(p=>{
      if(report[p.number].total!==p.total || report[p.number].matches===false) {
        throw new Error('大問'+p.number+'の描画予定の枠数が解析結果と一致しません。');
      }
    });
  } catch(error) {
    qRange.setValues(qBefore);dRange.setValues(dBefore);titleRange.setValues(titleBefore);
    throw new Error('登録できませんでした。元の設定に戻しました。詳細: '+(error.message||error));
  }
  generateAnswerSheet();
  return '大問'+prepared.map(p=>p.number).join('、')+'を追加し、合計'+
    prepared.reduce((sum,p)=>sum+p.total,0)+'枠の照合を完了しました。';
}

function documentSettingsFingerprint_(questions,details) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,JSON.stringify({questions,details})));
}