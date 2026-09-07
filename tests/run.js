/**
 * Pure-function tests for Grader.js — run with `node tests/run.js` (or `npm test`).
 *
 * Grader.js is written for the Apps Script runtime and has no module exports, so
 * this harness loads the real source into a Node `vm` sandbox with the Apps
 * Script globals it touches stubbed out, then pulls the pure helpers off the
 * sandbox and exercises them. Testing the source directly (not a copy) means
 * these assertions can't drift from the code they cover.
 *
 * Only the genuinely pure helpers are tested here — anything that calls
 * SpreadsheetApp / UrlFetchApp / PropertiesService is integration-only and can
 * only run bound to a real sheet.
 */
'use strict';

var fs = require('fs');
var path = require('path');
var vm = require('vm');

// --- Load Grader.js with Apps Script globals stubbed -----------------------
var source = fs.readFileSync(path.join(__dirname, '..', 'Grader.js'), 'utf8');
var sandbox = {
  Logger: { log: function () {} },
  Utilities: { sleep: function () {} }
  // SpreadsheetApp / UrlFetchApp / PropertiesService / LockService are only
  // referenced by functions we don't unit-test, so they stay undefined.
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
eq('empty string -> []', parseExcludeKeywords_(''), []);
eq('whitespace-only entries dropped', parseExcludeKeywords_(' , ,  '), []);
eq('lowercases + trims', kws('  FoO , Bar '),
  [{ text: 'foo', titleOnly: false, pattern: '\\bfoo\\b' },
   { text: 'bar', titleOnly: false, pattern: '\\bbar\\b' }]);
eq('title: prefix sets titleOnly', kws('title:remote, intern'),
  [{ text: 'remote', titleOnly: true, pattern: '\\bremote\\b' },
   { text: 'intern', titleOnly: false, pattern: '\\bintern\\b' }]);
eq('bare "title:" with no text is dropped', parseExcludeKeywords_('title:'), []);
eq('regex punctuation in a keyword is escaped, not interpreted',
  kws('node.js'), [{ text: 'node.js', titleOnly: false, pattern: '\\bnode\\.js\\b' }]);
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

// --- isGradableStatus_ (which rows a run picks up) -------------------------
check('"new" is gradable', isGradableStatus_('new') === true);
check('"regrade" is gradable', isGradableStatus_('regrade') === true);
check('padded + mixed-case status still matches', isGradableStatus_('  ReGrade ') === true);
check('"graded" is not re-picked up', isGradableStatus_('graded') === false);
check('empty status is not gradable', isGradableStatus_('') === false);
check('blank cell (empty string from the sheet) is not gradable', isGradableStatus_('   ') === false);
check('null/undefined cell does not throw', isGradableStatus_(null) === false && isGradableStatus_(undefined) === false);

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
