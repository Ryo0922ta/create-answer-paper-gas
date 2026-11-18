/** 入力内容を記録せず、操作の経過を実行ログに残す。 */
function logEvent_(level, operation, event, details) {
  const entry = { operation: operation, event: event };
  const allowed = ['durationMs', 'partCount', 'basicCount', 'questionCount',
    'detailCount', 'conflictCount', 'sheetRow', 'emptyCount', 'createdCount'];
  allowed.forEach(function (key) {
    if (details && Number.isFinite(details[key])) entry[key] = details[key];
  });
  if (details && typeof details.range === 'string' && /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/.test(details.range)) {
    entry.range = details.range;
  }
  try {
    const method = level === 'error' ? 'error' : level === 'warn' ? 'warn' : 'info';
    console[method](JSON.stringify(entry));
  } catch (_) {
    // ログの失敗で本来の処理を止めない。
  }
}

function withOperationLog_(operation, callback, summarize) {
  const started = Date.now();
  logEvent_('info', operation, 'start');
  try {
    const result = callback();
    let details = {};
    try { details = summarize ? summarize(result) || {} : {}; } catch (_) {}
    details.durationMs = Date.now() - started;
    logEvent_('info', operation, 'complete', details);
    return result;
  } catch (error) {
    logEvent_('error', operation, 'failed', { durationMs: Date.now() - started });
    throw error;
  }
}

function draftLogCounts_(draft) {
  const parts = draft && Array.isArray(draft.parts) ? draft.parts : [];
  const questions = parts.filter(function (part) { return part.kind === 'question'; });
  return {
    partCount: parts.length,
    questionCount: questions.length,
    basicCount: parts.filter(function (part) { return part.kind === 'basic'; }).length,
    detailCount: questions.reduce(function (sum, part) {
      return sum + (Array.isArray(part.rows) ? part.rows.length : 0);
    }, 0),
  };
}
