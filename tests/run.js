/**
 * Pure-function tests for Grader.js — run with `node tests/run.js` (or `npm test`).
 *
 * Grader.js is written for the Apps Script runtime and has no module exports, so
 * this harness loads the real source into a Node `vm` sandbox with the Apps
 * Script globals it touches stubbed out, then pulls the pure helpers off the
 * sandbox and exercises them. Testing the source directly (not a copy) means
 * these assertions can't drift from the code they cover.
 *
 * The pure helpers are tested directly. The last section also runs
 * gradeNewRows() end to end against in-memory stand-ins for SpreadsheetApp /
 * UrlFetchApp / PropertiesService / LockService, which is what covers the
 * failure paths (rejected key, retired model, a row that moved mid-run). Real
 * Sheets and network behavior still only run bound to a real spreadsheet.
 */
'use strict';

var fs = require('fs');
var path = require('path');
var vm = require('vm');

// --- Load Grader.js with Apps Script globals stubbed -----------------------
var source = fs.readFileSync(path.join(__dirname, '..', 'Grader.js'), 'utf8');
var logLines = [];
var sandbox = {
  Logger: { log: function (m) { logLines.push(String(m)); } },
  Utilities: { sleep: function () {} }
  // SpreadsheetApp / UrlFetchApp / PropertiesService / LockService are
  // installed per scenario by the end-to-end section below.
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: 'Grader.js' });

// Top-level `var`/`function` declarations land on the sandbox object.
var parseExcludeKeywords_ = sandbox.parseExcludeKeywords_;
var parseGradeResponse_ = sandbox.parseGradeResponse_;
var checkExcludeKeywords_ = sandbox.checkExcludeKeywords_;
var guessRowTitle_ = sandbox.guessRowTitle_;
var resolveColumns_ = sandbox.resolveColumns_;
var buildRunLogRow_ = sandbox.buildRunLogRow_;
var parseCriteriaValues_ = sandbox.parseCriteriaValues_;
var isGradableStatus_ = sandbox.isGradableStatus_;
var findRowTitle_ = sandbox.findRowTitle_;
var normalizeConfig_ = sandbox.normalizeConfig_;
var buildGradingPrompt_ = sandbox.buildGradingPrompt_;
var rowMatchesSnapshot_ = sandbox.rowMatchesSnapshot_;
var gradeNewRows = sandbox.gradeNewRows;
var RUN_LOG_HEADERS = sandbox.RUN_LOG_HEADERS;
var EXCLUDE_KEYWORDS_PLACEHOLDER = sandbox.EXCLUDE_KEYWORDS_PLACEHOLDER;

// --- Tiny assertion harness ------------------------------------------------
var passed = 0;
var failures = [];
function check(name, cond) {
  if (cond) { passed++; }
  else { failures.push(name); }
}
function eq(name, actual, expected) {
  check(name + ' (got ' + JSON.stringify(actual) + ')',
    JSON.stringify(actual) === JSON.stringify(expected));
}

// --- parseExcludeKeywords_ -------------------------------------------------
// Entries carry a pre-compiled RegExp, which JSON.stringify flattens to {}, so
// compare its source string instead -- that also pins the escaping behavior.
function kws(csv) {
  return parseExcludeKeywords_(csv).map(function (k) {
    return { text: k.text, titleOnly: k.titleOnly, pattern: k.pattern.source };
  });
}
// Whole-word edges are lookarounds, not \b, so punctuation-edged keywords work.
function ww(escaped) { return '(?<![A-Za-z0-9_])' + escaped + '(?![A-Za-z0-9_])'; }
eq('empty string -> []', parseExcludeKeywords_(''), []);
eq('whitespace-only entries dropped', parseExcludeKeywords_(' , ,  '), []);
eq('lowercases + trims', kws('  FoO , Bar '),
  [{ text: 'foo', titleOnly: false, pattern: ww('foo') },
   { text: 'bar', titleOnly: false, pattern: ww('bar') }]);
eq('title: prefix sets titleOnly', kws('title:remote, intern'),
  [{ text: 'remote', titleOnly: true, pattern: ww('remote') },
   { text: 'intern', titleOnly: false, pattern: ww('intern') }]);
eq('bare "title:" with no text is dropped', parseExcludeKeywords_('title:'), []);
eq('regex punctuation in a keyword is escaped, not interpreted',
  kws('node.js'), [{ text: 'node.js', titleOnly: false, pattern: ww('node\\.js') }]);
// \b needs a word character on its side of the edge, so these could never match.
eq('keyword ending in punctuation (c++) matches as a whole word',
  checkExcludeKeywords_({ body: 'senior c++ developer', _row: 2 }, parseExcludeKeywords_('c++')), 'c++');
eq('keyword starting with punctuation (.net) matches as a whole word',
  checkExcludeKeywords_({ body: 'hiring .net engineer', _row: 2 }, parseExcludeKeywords_('.net')), '.net');
eq('keyword ending in # (c#) matches as a whole word',
  checkExcludeKeywords_({ body: 'c# and azure', _row: 2 }, parseExcludeKeywords_('c#')), 'c#');
eq('punctuation-edged keyword still respects the word edge (c++ vs c++x)',
  checkExcludeKeywords_({ body: 'c++x toolkit', _row: 2 }, parseExcludeKeywords_('c++')), null);
check('escaped keyword matches literally, not as a wildcard',
  checkExcludeKeywords_({ body: 'node.js backend', _row: 2 }, parseExcludeKeywords_('node.js')) === 'node.js' &&
  checkExcludeKeywords_({ body: 'nodexjs backend', _row: 2 }, parseExcludeKeywords_('node.js')) === null);
// The compiled pattern is reused across rows, so it must not carry match state
// (a /g flag would advance lastIndex between calls and drop every other hit).
(function () {
  var reused = parseExcludeKeywords_('intern');
  var row = { body: 'summer intern role', _row: 2 };
  check('compiled pattern is stateless across repeated rows',
    checkExcludeKeywords_(row, reused) === 'intern' &&
    checkExcludeKeywords_(row, reused) === 'intern' &&
    checkExcludeKeywords_(row, reused) === 'intern');
})();

// --- parseGradeResponse_ ---------------------------------------------------
eq('null input -> null', parseGradeResponse_(null), null);
eq('no GRADE line -> null', parseGradeResponse_('nothing here'), null);
eq('valid grade + reasoning',
  parseGradeResponse_('GRADE: A+\nREASONING: Strong fit. Clear match.'),
  { grade: 'A+', reasoning: 'Strong fit. Clear match.' });
eq('lowercase grade is upcased',
  parseGradeResponse_('GRADE: a-\nREASONING: ok'),
  { grade: 'A-', reasoning: 'ok' });
eq('grade outside VALID_GRADES (E) -> null',
  parseGradeResponse_('GRADE: E\nREASONING: x'), null);
eq('missing reasoning gets default',
  parseGradeResponse_('GRADE: B'),
  { grade: 'B', reasoning: 'No reasoning provided.' });
// Markdown-bolded labels: models routinely emphasize a label they were told to
// emit verbatim, and the payload is still valid.
eq('bold label with colon inside the emphasis',
  parseGradeResponse_('**GRADE:** A\n**REASONING:** Strong fit. Clear match.'),
  { grade: 'A', reasoning: 'Strong fit. Clear match.' });
eq('bold label with colon outside the emphasis',
  parseGradeResponse_('**GRADE**: B-\n**REASONING**: Partial match. Some gaps.'),
  { grade: 'B-', reasoning: 'Partial match. Some gaps.' });
eq('grade token is not the prefix of a longer word',
  parseGradeResponse_('GRADE: Apple\nREASONING: x'), null);
eq('lowercase labels keep the reasoning (not the default text)',
  parseGradeResponse_('Grade: B\nReasoning: solid fit overall.'),
  { grade: 'B', reasoning: 'solid fit overall.' });
// Labels are anchored to a line start, and reasoning is only read after GRADE.
eq('a preamble mentioning "reasoning:" is not captured as the reasoning',
  parseGradeResponse_('Here is my reasoning: it fits.\nGRADE: B\nREASONING: Good.'),
  { grade: 'B', reasoning: 'Good.' });
eq('"Upgrade: A" is not read as a grade', parseGradeResponse_('Upgrade: A\nREASONING: x'), null);
eq('GRADE mid-line is not read as a grade',
  parseGradeResponse_('I would not grade: A here.\nREASONING: x'), null);
eq('bracketed grade echo (GRADE: [B]) is accepted',
  parseGradeResponse_('GRADE: [B]\nREASONING: ok'), { grade: 'B', reasoning: 'ok' });
eq('both labels on one line still parse',
  parseGradeResponse_('GRADE: B REASONING: ok'), { grade: 'B', reasoning: 'ok' });
eq('leading blank line and bold labels still parse',
  parseGradeResponse_('\n**GRADE**: A\n**REASONING**: Great.'), { grade: 'A', reasoning: 'Great.' });
// The grade pattern is built from VALID_GRADES, so a new scale needs no other edit.
(function () {
  var original = sandbox.GRADER_CONFIG.VALID_GRADES;
  sandbox.GRADER_CONFIG.VALID_GRADES = ['10', '5', '4', '3', '2', '1'];
  eq('numeric scale parses once VALID_GRADES is changed',
    parseGradeResponse_('GRADE: 4\nREASONING: ok'), { grade: '4', reasoning: 'ok' });
  eq('numeric scale: "10" is not read as "1"',
    parseGradeResponse_('**GRADE:** 10\nREASONING: ok').grade, '10');
  eq('numeric scale: out-of-scale number -> null',
    parseGradeResponse_('GRADE: 7\nREASONING: ok'), null);
  sandbox.GRADER_CONFIG.VALID_GRADES = ['Pass', 'Fail'];
  eq('word scale matched case-insensitively, returned as configured',
    parseGradeResponse_('GRADE: PASS\nREASONING: ok').grade, 'Pass');
  sandbox.GRADER_CONFIG.VALID_GRADES = original;
})();
(function () {
  var long = 'GRADE: F\nREASONING: ' + new Array(1000).join('x');
  var out = parseGradeResponse_(long);
  check('long reasoning truncated to <=800 with ellipsis',
    out.reasoning.length === 800 && out.reasoning.slice(-3) === '...');
})();

// --- checkExcludeKeywords_ (whole-word, skip-column, title-only) -----------
var kw = parseExcludeKeywords_('intern, title:remote');
eq('no keywords -> null', checkExcludeKeywords_({ body: 'anything', _row: 2 }, []), null);
eq('whole-word match returns the keyword',
  checkExcludeKeywords_({ body: 'summer intern role', _row: 2 }, kw), 'intern');
eq('word-boundary: "internal" does not match "intern"',
  checkExcludeKeywords_({ body: 'internal tooling team', _row: 2 }, kw), null);
eq('title-only keyword matches the title field',
  checkExcludeKeywords_({ title: 'Remote SWE', _row: 2 }, kw), 'remote');
eq('title-only keyword does NOT match body text',
  checkExcludeKeywords_({ name: 'X', body: 'fully remote', _row: 2 }, parseExcludeKeywords_('title:remote')), null);
eq('keyword in a skipped column (status) is ignored',
  checkExcludeKeywords_({ status: 'intern', _row: 2 }, kw), null);
// With no title/name column, a title-only keyword must not fall through to
// whatever column happens to come first (that produced false-positive Fs).
eq('title-only keyword does NOT match a body column when there is no title column',
  checkExcludeKeywords_({ description: 'permanently closed last year', _row: 2 }, parseExcludeKeywords_('title:closed')), null);
eq('title-only keyword sees the whole title, not an 80-char slice',
  checkExcludeKeywords_({ title: new Array(101).join('x') + ' closed', _row: 2 }, parseExcludeKeywords_('title:closed')), 'closed');
eq('findRowTitle_ -> null when no title column', findRowTitle_({ body: 'B', _row: 2 }), null);
eq('findRowTitle_ skips an empty title and uses name', findRowTitle_({ title: '', name: 'N', _row: 2 }), 'N');

// --- isGradableStatus_ (which rows a run picks up) -------------------------
check('"new" is gradable', isGradableStatus_('new') === true);
check('"regrade" is gradable', isGradableStatus_('regrade') === true);
check('padded + mixed-case status still matches', isGradableStatus_('  ReGrade ') === true);
check('"graded" is not re-picked up', isGradableStatus_('graded') === false);
check('empty status is not gradable', isGradableStatus_('') === false);
check('blank cell (empty string from the sheet) is not gradable', isGradableStatus_('   ') === false);
check('null/undefined cell does not throw', isGradableStatus_(null) === false && isGradableStatus_(undefined) === false);
// Config entries are normalized too, so a capitalized custom status isn't a silent no-op.
(function () {
  var original = sandbox.GRADER_CONFIG.STATUSES_TO_GRADE;
  sandbox.GRADER_CONFIG.STATUSES_TO_GRADE = ['new', ' Recheck '];
  check('capitalized/padded config entry still matches its cell',
    isGradableStatus_('recheck') === true && isGradableStatus_('RECHECK') === true);
  sandbox.GRADER_CONFIG.STATUSES_TO_GRADE = original;
})();

// --- normalizeConfig_ (config entries vs lowercased sheet headers) --------
(function () {
  var cfg = { SKIP_COLUMNS_IN_PROMPT: ['status', ' Notes '], TITLE_COLUMNS: ['Title'],
    STATUSES_TO_GRADE: ['New', ''], STATUS_COLUMN: 'Status', GRADE_COLUMN: 'grade',
    REASONING_COLUMN: 'reasoning', CRITERIA_KEY_HEADER: 'Field', CRITERIA_VALUE_HEADER: 'value' };
  normalizeConfig_(cfg);
  eq('normalizeConfig_ lowercases + trims lists and drops blanks',
    [cfg.SKIP_COLUMNS_IN_PROMPT, cfg.TITLE_COLUMNS, cfg.STATUSES_TO_GRADE],
    [['status', 'notes'], ['title'], ['new']]);
  eq('normalizeConfig_ lowercases column/header names', [cfg.STATUS_COLUMN, cfg.CRITERIA_KEY_HEADER], ['status', 'field']);

  var original = sandbox.GRADER_CONFIG.SKIP_COLUMNS_IN_PROMPT;
  sandbox.GRADER_CONFIG.SKIP_COLUMNS_IN_PROMPT = ['Notes'];
  normalizeConfig_(sandbox.GRADER_CONFIG);
  check('a capitalized SKIP_COLUMNS_IN_PROMPT entry keeps that column out of the prompt',
    buildGradingPrompt_('rubric', { notes: 'SECRET', body: 'B', _row: 2 }).indexOf('SECRET') === -1);
  sandbox.GRADER_CONFIG.SKIP_COLUMNS_IN_PROMPT = original;
})();

// --- parseCriteriaValues_ (header-named Criteria sheet read) ---------------
eq('reads key/value by header name',
  (function () {
    var c = parseCriteriaValues_([['field', 'value'], ['criteria_text', 'grade it'], ['exclude_keywords', 'a, b']]);
    return [c.criteria_text, c.exclude_keywords, c._keys];
  })(),
  ['grade it', 'a, b', ['criteria_text', 'exclude_keywords']]);
eq('a column inserted to the LEFT no longer blanks the rubric',
  parseCriteriaValues_([['notes', 'field', 'value'], ['x', 'criteria_text', 'grade it']]).criteria_text,
  'grade it');
eq('a column inserted BETWEEN field and value still resolves',
  parseCriteriaValues_([['field', 'notes', 'value'], ['criteria_text', 'x', 'grade it']]).criteria_text,
  'grade it');
eq('headers are matched case-insensitively and trimmed',
  parseCriteriaValues_([[' Field ', 'VALUE'], ['criteria_text', 'grade it']]).criteria_text, 'grade it');
eq('missing headers fall back to the first two columns',
  parseCriteriaValues_([['k', 'v'], ['criteria_text', 'grade it']]).criteria_text, 'grade it');
eq('header row only -> empty map with empty _keys',
  parseCriteriaValues_([['field', 'value']])._keys, []);
eq('rows with a blank key are skipped',
  parseCriteriaValues_([['field', 'value'], ['', 'orphan'], ['criteria_text', 'grade it']])._keys,
  ['criteria_text']);
eq('_keys reports what was found, for the empty-criteria error message',
  parseCriteriaValues_([['field', 'value'], ['critera_text', 'typo'], ['exclude_keywords', '']])._keys,
  ['critera_text', 'exclude_keywords']);

// --- guessRowTitle_ --------------------------------------------------------
eq('prefers title', guessRowTitle_({ title: 'T', name: 'N', _row: 2 }), 'T');
eq('falls back to name', guessRowTitle_({ name: 'N', body: 'B', _row: 2 }), 'N');
eq('then first non-skipped column', guessRowTitle_({ status: 'new', body: 'B', _row: 2 }), 'B');
eq('fallback to row number', guessRowTitle_({ status: 'new', _row: 7 }), 'row 7');

// --- resolveColumns_ (presence + duplicate-required hard-fail) -------------
eq('maps the three required columns',
  resolveColumns_({ status: 1, grade: 2, reasoning: 3 }),
  { gradeCol: 2, reasoningCol: 3, statusCol: 1 });
eq('missing a required column -> null',
  resolveColumns_({ status: 1, grade: 2 }), null);
eq('duplicated required column (grade) hard-fails -> null',
  resolveColumns_({ status: 1, grade: 2, reasoning: 3 }, { grade: true }), null);
eq('duplicate of a non-required column does not hard-fail',
  resolveColumns_({ status: 1, grade: 2, reasoning: 3 }, { body: true }),
  { gradeCol: 2, reasoningCol: 3, statusCol: 1 });

// --- buildRunLogRow_ (run-summary row shaping) -----------------------------
eq('run-log row matches header order + rounds elapsed',
  buildRunLogRow_({ total: 10, graded: 7, rejected: 2, errors: 1, skipped: 0 },
    '2026-06-18T00:00:00.000Z', 12.34),
  ['2026-06-18T00:00:00.000Z', 10, 7, 2, 1, 0, 12.3]);
eq('run-log row width equals header width',
  buildRunLogRow_({ total: 0, graded: 0, rejected: 0, errors: 0, skipped: 0 }, 'x', 0).length,
  RUN_LOG_HEADERS.length);

// --- rowMatchesSnapshot_ (moved-row guard) ---------------------------------
check('identical rows match', rowMatchesSnapshot_(['a', 1, ''], ['a', 1, '']) === true);
check('numbers and dates compare by content',
  rowMatchesSnapshot_([2, new Date(0)], [2, new Date(0)]) === true);
check('a changed cell does not match', rowMatchesSnapshot_(['a', 'new'], ['a', 'graded']) === false);
check('a different width (inserted column) does not match', rowMatchesSnapshot_(['a'], ['a', '']) === false);
check('missing current row does not match', rowMatchesSnapshot_(['a'], undefined) === false);

// --- gradeNewRows end to end (stubbed Sheets / fetch / properties / lock) ----
// Minimal in-memory Sheet: enough of getRange/getValues/setValues/appendRow
// for Grader.js. `grid` is a 2D array, 0-based; the API is 1-based.
function FakeSheet(grid) { this.grid = grid; }
FakeSheet.prototype.getLastRow = function () { return this.grid.length; };
FakeSheet.prototype.getLastColumn = function () {
  var w = 0;
  for (var i = 0; i < this.grid.length; i++) w = Math.max(w, this.grid[i].length);
  return w;
};
FakeSheet.prototype.set = function (r, c, v) {
  while (this.grid.length < r) this.grid.push([]);
  this.grid[r - 1][c - 1] = v;
};
FakeSheet.prototype.getRange = function (r, c, nr, nc) {
  var self = this; nr = nr || 1; nc = nc || 1;
  return {
    getValues: function () {
      var out = [];
      for (var i = 0; i < nr; i++) {
        var line = [];
        for (var j = 0; j < nc; j++) {
          var v = (self.grid[r - 1 + i] || [])[c - 1 + j];
          line.push(v === undefined ? '' : v);
        }
        out.push(line);
      }
      return out;
    },
    setValues: function (vals) {
      for (var i = 0; i < vals.length; i++)
        for (var j = 0; j < vals[i].length; j++) self.set(r + i, c + j, vals[i][j]);
    },
    setValue: function (v) { self.set(r, c, v); },
    setFontWeight: function () {}
  };
};
FakeSheet.prototype.appendRow = function (vals) { this.grid.push(vals.slice()); };
FakeSheet.prototype.setFrozenRows = function () {};
FakeSheet.prototype.setColumnWidth = function () {};

var cfg = sandbox.GRADER_CONFIG;
var HEADERS = ['title', 'body', 'status', 'grade', 'reasoning'];
function dataGrid() {
  return [
    HEADERS.slice(),
    ['Good job', 'great fit', 'new', '', ''],
    ['Intern role', 'summer', 'new', '', ''],
    ['Done already', 'x', 'graded', 'A', 'kept']
  ];
}
// Installs the Apps Script globals for one scenario and runs gradeNewRows.
// `respond(callNumber, sheets)` returns { code, body } for each fetch.
function runScenario(opts) {
  var sheets = {};
  sheets[cfg.DATA_SHEET] = new FakeSheet(opts.data || dataGrid());
  sheets[cfg.CRITERIA_SHEET] = new FakeSheet([
    ['field', 'value'], ['criteria_text', 'grade it'], ['exclude_keywords', opts.keywords || '']
  ]);
  var ss = {
    getSheetByName: function (n) { return sheets[n] || null; },
    insertSheet: function (n) { sheets[n] = new FakeSheet([]); return sheets[n]; }
  };
  var calls = 0;
  sandbox.SpreadsheetApp = { getActiveSpreadsheet: function () { return ss; } };
  sandbox.PropertiesService = { getScriptProperties: function () {
    return { getProperty: function () { return 'apiKey' in opts ? opts.apiKey : 'test-key'; } };
  } };
  sandbox.LockService = { getScriptLock: function () {
    return { tryLock: function () { return true; }, releaseLock: function () {} };
  } };
  sandbox.UrlFetchApp = { fetch: function () {
    calls++;
    var r = opts.respond(calls, sheets);
    return { getResponseCode: function () { return r.code; }, getContentText: function () { return r.body; } };
  } };
  logLines.length = 0;
  var threw = null;
  try { gradeNewRows(); } catch (e) { threw = e; }
  var log = sheets[cfg.LOG_SHEET];
  return {
    threw: threw, calls: calls, data: sheets[cfg.DATA_SHEET].grid,
    // Log-sheet row without run_at / elapsed_sec: [total, graded, rejected, errors, deferred]
    logRow: log && log.grid.length === 2 ? log.grid[1].slice(1, 6) : null
  };
}
function ok200(content) {
  return { code: 200, body: JSON.stringify({ choices: [{ message: { content: content } }] }) };
}

// Happy path: one row graded by the API, one auto-rejected for free, one left alone.
(function () {
  var r = runScenario({ keywords: 'intern', respond: function () { return ok200('GRADE: B\nREASONING: Fine.'); } });
  check('happy path: does not throw', r.threw === null);
  eq('happy path: API row gets grade/reasoning/status', r.data[1].slice(2), ['graded', 'B', 'Fine.']);
  check('happy path: keyword row auto-rejected with lowest grade and no API call',
    r.data[2][2] === 'graded' && r.data[2][3] === 'F' && r.calls === 1);
  eq('happy path: already-graded row untouched', r.data[3], ['Done already', 'x', 'graded', 'A', 'kept']);
  eq('happy path: Log row counts', r.logRow, [2, 1, 1, 0, 0]);
})();

// Missing key: the pre-existing config check still throws before any work.
(function () {
  var r = runScenario({ apiKey: null, respond: function () { return ok200('GRADE: B\nREASONING: Fine.'); } });
  check('missing key: throws naming the property',
    r.threw !== null && r.threw.message.indexOf(cfg.API_KEY_PROPERTY) !== -1);
  check('missing key: no API call, no Log row', r.calls === 0 && r.logRow === null);
})();

// Rejected key: a 401 is fatal -- stop after the first call, write the Log row, throw.
(function () {
  var r = runScenario({ keywords: '', respond: function () { return { code: 401, body: '{"error":"invalid api key"}' }; } });
  check('401: throws', r.threw !== null && /401/.test(r.threw.message));
  check('401: only one API call made (remaining rows deferred, not retried)', r.calls === 1);
  eq('401: Log row written first (1 error, 1 deferred)', r.logRow, [2, 0, 0, 1, 1]);
  check('401: rows keep their status', r.data[1][2] === 'new' && r.data[2][2] === 'new');
})();
(function () {
  var r = runScenario({ keywords: '', respond: function () { return { code: 404, body: 'model not found' }; } });
  check('404: throws naming the model/endpoint', r.threw !== null && /API_MODEL/.test(r.threw.message));
})();

// Retired model on a provider that answers 400: no single call is fatal, but a
// run where every API-graded row failed and none succeeded still throws
// (after its Log row), even when free auto-rejects happened alongside.
(function () {
  var r = runScenario({ keywords: 'intern', respond: function () { return { code: 400, body: 'model_decommissioned' }; } });
  check('all-failed: throws', r.threw !== null && /failed/.test(r.threw.message));
  eq('all-failed: Log row written first (1 rejected, 1 error)', r.logRow, [2, 0, 1, 1, 0]);
  check('all-failed: reject still landed, API row still new', r.data[2][2] === 'graded' && r.data[1][2] === 'new');
})();
// A run with at least one success and some per-row errors is not a broken run.
(function () {
  var data = dataGrid(); data[2][2] = 'new';
  var r = runScenario({ data: data, keywords: '', respond: function (n) {
    return n === 1 ? ok200('GRADE: B\nREASONING: Fine.') : { code: 500, body: 'oops' };
  } });
  check('partial failure: does not throw', r.threw === null);
  eq('partial failure: Log row (1 graded, 1 error)', r.logRow, [2, 1, 0, 1, 0]);
})();

// Sheet sorted mid-run: the API call succeeds, but the row at that number now
// holds different data, so the write is skipped and the row is deferred.
(function () {
  var r = runScenario({ keywords: 'intern', respond: function (n, sheets) {
    var g = sheets[cfg.DATA_SHEET].grid;
    var tmp = g[1]; g[1] = g[2]; g[2] = tmp;  // swap rows 2 and 3, as a sort would
    return ok200('GRADE: B\nREASONING: Fine.');
  } });
  check('moved rows: does not throw', r.threw === null);
  check('moved rows: nothing written to the wrong row',
    r.data[1][2] === 'new' && r.data[1][3] === '' && r.data[2][2] === 'new' && r.data[2][3] === '');
  eq('moved rows: both rows deferred in the Log row', r.logRow, [2, 0, 0, 0, 2]);
  check('moved rows: a SKIP line was logged', logLines.some(function (l) { return l.indexOf('[SKIP]') !== -1; }));
})();

// --- placeholder sentinel regression guard ---------------------------------
check('EXCLUDE_KEYWORDS_PLACEHOLDER has no comma (stays a single non-matching token)',
  EXCLUDE_KEYWORDS_PLACEHOLDER.indexOf(',') === -1);

// --- report ----------------------------------------------------------------
if (failures.length) {
  console.error('✗ ' + failures.length + ' failed, ' + passed + ' passed:');
  failures.forEach(function (f) { console.error('  - ' + f); });
  process.exit(1);
}
console.log('✓ all ' + passed + ' assertions passed');
