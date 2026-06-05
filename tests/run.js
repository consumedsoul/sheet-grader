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
eq('empty string -> []', parseExcludeKeywords_(''), []);
eq('whitespace-only entries dropped', parseExcludeKeywords_(' , ,  '), []);
eq('lowercases + trims', parseExcludeKeywords_('  FoO , Bar '),
  [{ text: 'foo', titleOnly: false }, { text: 'bar', titleOnly: false }]);
eq('title: prefix sets titleOnly', parseExcludeKeywords_('title:remote, intern'),
  [{ text: 'remote', titleOnly: true }, { text: 'intern', titleOnly: false }]);
eq('bare "title:" with no text is dropped', parseExcludeKeywords_('title:'), []);

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
  checkExcludeKeywords_({ name: 'X', body: 'fully remote', _row: 2 }, [{ text: 'remote', titleOnly: true }]), null);
eq('keyword in a skipped column (status) is ignored',
  checkExcludeKeywords_({ status: 'intern', _row: 2 }, kw), null);

// --- guessRowTitle_ --------------------------------------------------------
eq('prefers title', guessRowTitle_({ title: 'T', name: 'N', _row: 2 }), 'T');
eq('falls back to name', guessRowTitle_({ name: 'N', body: 'B', _row: 2 }), 'N');
eq('then first non-skipped column', guessRowTitle_({ status: 'new', body: 'B', _row: 2 }), 'B');
eq('fallback to row number', guessRowTitle_({ status: 'new', _row: 7 }), 'row 7');

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
