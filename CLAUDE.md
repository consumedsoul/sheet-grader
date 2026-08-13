# CLAUDE.md — sheet-grader

## What this is

A single-file Google Apps Script (`Grader.js`) that grades rows of a Google Sheet with
an LLM. Reads `status="new"` rows from a **Data** sheet, grades each against a rubric the
user writes in a **Criteria** sheet, writes `grade` + `reasoning` back, flips status to
`graded`. Provider-agnostic via the OpenAI-compatible `chat/completions` format (default:
Groq GPT OSS 20B).

This is a template/example repo meant to be copied into other people's sheets, so keep it
generic and keep the README honest. It is not wired up to a live spreadsheet here — there
is no test harness or runtime in this repo; correctness is reviewed by reading the code.

## Architecture (one file, top-down in `Grader.js`)

- `GRADER_CONFIG` (top) — all tunables: endpoint, model, key property name, delays,
  timer budget, sheet/column names, skip columns, truncation, valid grades. **Change
  behavior here, not in the functions.**
- `DEFAULT_CRITERIA` — placeholder text seeded into the Criteria sheet on first run.
- `gradeNewRows()` — the only entry point. Schedule it with a time-based trigger.
  Takes a `LockService` script lock, validates key + criteria, maps headers→columns
  (`resolveColumns_`), then loops rows calling `gradeOneRow_` and pacing against the rate
  limit. Respects a timer budget so it can resume across runs.
- `gradeOneRow_` — Stage 1 (regex reject, no API call) → Stage 2 (LLM) → write back for a
  single row. Returns `{ outcome, madeApiCall }`; the caller uses `outcome` for stats and
  `madeApiCall` to decide whether to sleep (auto-rejects don't sleep).
- Other helpers (all suffixed `_`, Apps Script's private convention):
  `resolveColumns_`, `getOrCreateCriteria_`, `parseExcludeKeywords_`, `getUngradedRows_`,
  `guessRowTitle_`, `checkExcludeKeywords_`, `callLlmApi_`, `buildGradingPrompt_`,
  `parseGradeResponse_`, `updateRowGrade_`, `buildRunLogRow_` / `appendRunLog_`
  (end-of-run summary to the `Log` sheet, gated on `ENABLE_RUN_LOG`).

## Conventions

- **Config-driven, not hardcoded.** New knobs go in `GRADER_CONFIG`.
- **Private helpers end in `_`** (Apps Script hides them from the Run menu).
- **The rubric lives in the sheet, not the code.** That keeps the user's prompt out of
  git. Never inline a real rubric into `DEFAULT_CRITERIA`.
- **Writes are per-row and immediate** (`updateRowGrade_`) on purpose — a mid-run crash
  must not lose completed work. The three cells of a single row may be batched into one
  `setValues` call (when contiguous), but never batch *across* rows into a write-at-the-end.
- ES5-flavored V8: uses `var`, `for` loops, no async/await (Apps Script `UrlFetchApp` is
  synchronous). Match the existing style.

## Gotchas (Apps Script specific)

- **6-minute execution cap.** `TIMER_BUDGET_SEC` (default 300) bails early; leftover rows
  stay `new` and resume next run. Don't write loops that can't be resumed.
- **Secrets live in Script Properties, never in code.** Required property:
  `LLM_API_KEY`. Set via Project Settings → Script Properties.
- **OAuth scopes are declared explicitly** in `appsscript.json`
  (`spreadsheets.currentonly` + `script.external_request`). If you add an API that needs a
  broader scope, add it there deliberately rather than deleting the array and letting Apps
  Script re-infer.
- **`.clasp.json` carries the script ID and is gitignored.** Only `.clasp.json.example`
  is committed. Never commit a real `.clasp.json`.
- **File extension rename:** clasp pushes `Grader.js` to Apps Script as a `.gs` file. When
  pasting manually, name the editor file `Grader.gs`.
- **Rate limits:** `API_DELAY_MS` (default 7000) paces calls for Groq's free tier, and
  only fires after an actual API call. One retry on HTTP 429 with a 5s backoff
  (`callLlmApi_`).
- **Overlap guard:** `gradeNewRows` takes a `LockService` script lock so a manual run and
  a trigger can't double-grade the same rows.

## Required Script Properties

| Property | Required | Purpose |
| --- | --- | --- |
| `LLM_API_KEY` | yes | Bearer token for the OpenAI-compatible endpoint (Groq/OpenAI/etc). |

## Known limitations / open items

- Untrusted row content is fenced in the prompt and the model is told to treat it as data,
  but this is a mitigation, not a guarantee against prompt injection. Fine for triage;
  don't trust grades from a hostile data source. Results only ever land in the sheet.
- `exclude_keywords` splits on `,` with no escape, so a single keyword can't contain a
  literal comma. Documented in the README.
- The pure helpers have a Node test harness (`npm test` / `node tests/run.js`); it stubs
  the Apps Script globals and exercises the real `Grader.js` source. The Apps Script
  *integration* (sheet I/O, API calls) has no runtime here — it only runs bound to a sheet.

## Safe-change checklist

- Touching the loop? Preserve resumability (timer budget) and per-row writes, and keep the
  rate-limit sleep gated on `madeApiCall` so auto-rejects stay free.
- Adding a provider? Only `API_ENDPOINT` / `API_MODEL` / key should need to change —
  plus `REASONING_EFFORT`, which is only sent when non-empty; blank it for providers
  that reject the parameter.
- Changing the grade scale? Update `VALID_GRADES` **and** the format string in
  `buildGradingPrompt_` **and** the regex in `parseGradeResponse_`.
- Changing `updateRowGrade_`? It takes a `cols` object (`{ gradeCol, reasoningCol,
  statusCol }`), not three separate args.

## Sync Policy (always — no prompting needed)

This is a **portfolio/example repo with no live bound script** — git is the only
deployment target. After ANY code/content change, finish the session by syncing
without being asked:

```
git add -A && git commit -m "<concise message>" && git push origin main
```

Never leave work uncommitted or unpushed, and never wait for Hun to say "push/commit/sync".
Do **not** run or offer `clasp push` — there is no `.clasp.json` and no live script to
deploy to. The `.clasp.json.example` / `.claspignore` files exist only to show how a
consumer of this template would wire it up to their own sheet.
