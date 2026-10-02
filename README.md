# sheet-grader

A small Google Apps Script that grades rows in a Google Sheet using an LLM.
You write a rubric, drop your data into a sheet, and on a schedule it reads
the unprocessed rows, grades each one A+ through F with a short reason, and
writes the result back next to the row.

It's two stages:

1. **Cheap regex pre-filter.** You can list dealbreaker keywords (e.g.
   `unpaid, internship, volunteer`). Any row that matches gets graded `F`
   immediately with no API call. On the project this was extracted from,
   this catches a meaningful fraction of rows for free.
2. **LLM grading.** Everything that survives the filter gets sent to an
   LLM with your rubric and the row's contents. The model replies in a
   fixed `GRADE: ... / REASONING: ...` format that the script parses
   strictly — anything malformed gets logged and skipped.

Uses the OpenAI-compatible `chat/completions` format, so it works with
[Groq](https://console.groq.com) (default, has a generous free tier),
OpenAI, OpenRouter, Together, Anyscale — anything that speaks that wire
format. Usually you just swap the endpoint, model name, and key — see
[Configuration](#configuration) for the one exception (newer reasoning-model
endpoints and the `max_tokens` field name).

This was extracted from a private project that uses the same pattern to
rank LA entertainment crew job listings against a working actor/producer's
profile. The generic version here strips out everything specific to that
use case so you can point it at anything.

## When you'd use this instead of ChatGPT

If you're going to grade two rows once, just use ChatGPT. This is for when:

- You have an ongoing flow of rows arriving in a sheet — scraped jobs,
  inbound leads, restaurant candidates, applications, listings, whatever
  — and you want them graded **without you having to babysit it**.
- You want the **rubric to stay consistent** across runs. The model reads
  the same criteria text every time, so grades are comparable week to
  week instead of drifting with whatever you happened to type into the
  chat.
- You want a **cheap filter in front of the LLM** so you're not paying
  (in tokens or rate limit) to have the model reject obvious junk.
- You want the output **back in the sheet** next to the source row, not
  in a chat transcript you have to copy-paste from.
- You want it to **run on a schedule** (daily, hourly) with zero hosting
  — Apps Script gives you triggers and a runtime for free.

If your data is high-stakes (medical, legal, hiring decisions) — don't
trust a small open-weight model's letter grade. Use this as a triage layer
that decides what's worth your own attention, not as the final word.

## How it works

```
Data sheet (status="new" or "regrade" rows)
         │
         ▼
  exclude-keyword regex  ──►  match? → grade F, no API call
         │
         ▼ no match
  build prompt: rubric + every column of the row
         │
         ▼
  POST /v1/chat/completions  (Groq / OpenAI / etc)
         │
         ▼
  parse "GRADE: X / REASONING: ..."
         │
         ▼
  write grade + reasoning back to the row, flip status to "graded"
```

A few details worth knowing:

- **Crash-safe.** Each row is written back to the sheet as soon as it's
  graded. If the script dies on row 47, rows 1–46 are already saved and
  the next run picks up from 47.
- **Respects the Apps Script 6-minute limit.** Defaults to bailing at 4.5
  minutes; remaining rows keep their existing status and get processed on
  the next scheduled run.
- **Sleeps between API calls** (default 7s — tuned for Groq's free tier
  6K tokens/minute). Drop this if you're on a paid tier.
- **One retry on HTTP 429** with a 5-second backoff.
- **Won't double-grade on overlap.** A script lock means if a manual run
  and a scheduled trigger fire at once, the second one exits immediately
  rather than grading the same rows twice.
- **Truncates long fields** in the prompt so token use stays predictable
  even if someone pastes a 50KB description into a cell.
- **Don't sort or insert rows in the Data tab while a run is in progress.**
  Each row is re-read right before its grade is written. If it has moved or
  changed since the run started, the write is skipped with a warning, the row
  counts as `deferred`, and it's graded on the next run — rather than landing
  on whatever row now sits at that number.

## Setup

You need a Google account and a free [Groq API key](https://console.groq.com)
(or any OpenAI-compatible key). Two ways to install — pick one.

### Option A: Copy-paste (no tools)

1. Open a new or existing Google Sheet.
2. Make a tab called `Data`. The first row must include at minimum these
   three columns: `status`, `grade`, `reasoning`. Any other columns are
   fair game — name them whatever makes sense (`title`, `description`,
   `url`, `company`, `address`, etc.). The grader will read all of them
   into the prompt.
3. Add rows with `status` set to `new` for anything you want graded.
4. **Extensions → Apps Script.** This opens the script editor bound to
   your sheet.
5. Delete the placeholder `Code.gs`. Create a new file called `Grader.gs`
   and paste the contents of [Grader.js](Grader.js) into it. (The repo file
   is named `Grader.js` so editors and tooling treat it as JavaScript; the
   Apps Script editor uses the `.gs` extension. clasp handles this rename
   automatically on push — when pasting manually, just name the editor file
   `Grader.gs`.)
6. In the Apps Script editor, **Project Settings (gear icon) → Script
   Properties → Add script property.** Set:
   - Name: `LLM_API_KEY`
   - Value: your Groq (or OpenAI) API key
7. Back in the editor, pick `gradeNewRows` from the function dropdown and
   click **Run**. Apps Script will ask for permission to access your
   sheet and make external HTTP calls — approve.
8. First run will create a `Criteria` sheet with placeholder text. Open
   that tab, edit the `criteria_text` cell to describe how you want
   things graded (see the example below), and edit `exclude_keywords` to
   list your dealbreakers (comma-separated).
9. Run `gradeNewRows` again. Your rows should fill in.
10. To make it automatic: in the editor, **Triggers (alarm clock icon) →
    Add Trigger.** Function: `gradeNewRows`. Event source: time-driven.
    Pick whatever cadence you want (e.g. daily, every hour). On a free
    (consumer) Google account, triggers get about 90 minutes of total run
    time per day. An idle run takes about a second, but a run working through
    a backlog uses its full ~4.5 minutes, so an hourly trigger can hit the cap
    mid-backlog and stop firing until the next day. During a big re-grade,
    every 2 hours is safer.

### Option B: clasp (recommended if you'll edit the code)

Install [clasp](https://github.com/google/clasp) and `clasp login` once,
then:

```bash
git clone https://github.com/YOUR_USERNAME/sheet-grader.git
cd sheet-grader

# In your Google Sheet: Extensions → Apps Script → Project Settings,
# copy the "Script ID".
cp .clasp.json.example .clasp.json
# Paste the Script ID into .clasp.json.

clasp push --force
```

Then continue from step 6 in Option A (set the API key, run the function).

## Writing the rubric

The `criteria_text` cell in the `Criteria` sheet is the single most
important thing in this whole project. The model is only as good as the
rubric.

**Concrete rubrics work better than vague ones.** Tell the model what
A+ looks like in your domain, what B looks like, what F looks like, and
what factors matter most. Example for grading restaurant candidates for
a date night:

```
You are grading restaurants for a Friday date night in Los Angeles.
Strict grading.

A+: walking distance from Silver Lake, $40-80/person, has cocktails,
    open past 10pm, vibe is "interesting" not "scene-y"
A:  same as A+ but missing one criterion
B:  decent backup, missing two criteria
C:  technically possible but underwhelming
F:  closed, too expensive, chain, or wrong neighborhood

Weight neighborhood and vibe more than price. Reject anywhere that
shows up on a "best of LA" listicle unless it's also genuinely good.
```

**`exclude_keywords` is your dealbreaker shortcut.** Comma-separated.
Anything matching gets `F` (the last entry in `VALID_GRADES`) instantly with
no API call. Matching is whole-word and case-insensitive, so `intern` won't
match `internal`; keywords that start or end with punctuation (`c++`, `.net`,
`c#`) work too. Use `title:` as a prefix to only match a row's title
(useful when a body description might mention the keyword in a benign way,
like "we are not an unpaid internship"). The title is the row's `title`
column, or `name` if `title` is empty — set `TITLE_COLUMNS` to use others.
**If your Data sheet has neither column, `title:` keywords never match**
(the run log warns about this). Example:

```
chain, fast casual, title:closed, title:permanently
```

The comma is the only separator and there's no escape for it, so a single
keyword can't contain a literal comma. Multi-word phrases without commas
(`fast casual`) are fine.

### Re-grading after a rubric change

Rubrics take a couple of passes to get right, and grades produced by the old
one are worth re-checking. To re-run a row, set its `status` cell back to
`regrade` and run `gradeNewRows` again — the row is picked up exactly like a
`new` one, and its `grade` and `reasoning` are overwritten with the new result.

To re-grade everything, select the whole `status` column below the header and
paste `regrade` down it. The usual limits still apply: a run stops at 4.5 minutes
and the rest carry over to the next one, so a large re-grade takes a few runs
(or a few trigger firings) to work through.

## Schema

### Data sheet (you create this)

Three required columns, anywhere in the header row:

| column | what goes in it |
| --- | --- |
| `status` | `new` for rows to grade. Becomes `graded` after. Set it back to `regrade` to run a row again — see [Re-grading after a rubric change](#re-grading-after-a-rubric-change). |
| `grade` | Filled in by the script (`A+` through `F`). |
| `reasoning` | Filled in by the script (~2 sentences). |

Any other columns are content for the LLM to evaluate. Name them
whatever — they get dumped into the prompt as `column_name: value`.

### Criteria sheet (auto-created on first run)

Two rows:

| field | value |
| --- | --- |
| `criteria_text` | Your rubric. |
| `exclude_keywords` | Comma-separated dealbreaker words. |

The two columns are found by their header names (`field` and `value`), not by
position, so you can insert your own columns around them. If you rename them,
change `CRITERIA_KEY_HEADER` / `CRITERIA_VALUE_HEADER` in `GRADER_CONFIG` to match
(the script seeds the sheet from those two values). If either name is missing,
the script falls back to reading the first two columns.

### Log sheet (auto-created on first run)

Each run appends one summary row — `run_at`, `total`, `graded`, `rejected`,
`errors`, `deferred`, `elapsed_sec` — so scheduled runs leave a trail without
opening the execution log. Disable with `ENABLE_RUN_LOG: false`.

A run that dies outright — no API key, a blank `criteria_text` on an existing
Criteria sheet, missing or duplicated required columns, an unhandled exception —
writes **no** Log row. It throws instead, so Apps Script marks the execution
failed and emails the trigger owner. A gap in the Log sheet means a broken run,
not a quiet one. (Setup step 8 — the first run that *creates* the Criteria
sheet — is expected, and exits quietly.)

A run whose API calls all fail — a revoked or rotated key, a retired model, a
wrong endpoint — is also treated as broken, not quiet: it writes its Log row
(so the `errors` column shows it) *and then* throws, so you still get the
email. A `401`, `403` or `404` stops the run at the first row instead of
retrying every row; a provider that answers with `400` or a server error is
caught at the end of the run, when no API-graded row succeeded.

## Configuration

All settings live in `GRADER_CONFIG` at the top of [Grader.js](Grader.js).
Change behavior there, not in the functions. Anything that names a column,
header, or status is matched case-insensitively.

**Provider and model**

| Key | Default | What it does |
| --- | --- | --- |
| `API_ENDPOINT` | Groq chat completions URL | Any OpenAI-compatible `chat/completions` endpoint. |
| `API_MODEL` | `openai/gpt-oss-20b` | Model name sent to that endpoint. |
| `API_KEY_PROPERTY` | `LLM_API_KEY` | Name of the Script Property holding your API key. |
| `REASONING_EFFORT` | `low` | `low`/`medium`/`high` for reasoning models like `openai/gpt-oss-*`, which spend hidden reasoning tokens out of `MAX_OUTPUT_TOKENS`. Set to `''` for providers that reject the parameter — it's only sent when non-empty. |
| `TEMPERATURE` | `0.3` | Lower = more consistent grades. |
| `MAX_OUTPUT_TOKENS` | `1200` | Ceiling on reasoning **and** visible output together. Too low and the reply gets clipped, which reads as a parse error. Sent as `max_tokens`; some newer reasoning-model endpoints (OpenAI's `o*`/`gpt-5*` among them) accept only `max_completion_tokens` — if a provider swap fails on that, rename the field in `callLlmApi_`. |

**Pacing**

| Key | Default | What it does |
| --- | --- | --- |
| `API_DELAY_MS` | `7000` | Pause after each API call, sized for Groq's free tier. Lower on paid tiers. |
| `TIMER_BUDGET_SEC` | `270` | A run stops starting new rows after this, leaving room under Apps Script's 6-minute cap for the last row's API call. Leftover rows carry over to the next run. |

**Sheets and columns**

| Key | Default | What it does |
| --- | --- | --- |
| `DATA_SHEET` / `CRITERIA_SHEET` / `LOG_SHEET` | `Data` / `Criteria` / `Log` | Tab names. |
| `STATUS_COLUMN` / `GRADE_COLUMN` / `REASONING_COLUMN` | `status` / `grade` / `reasoning` | Required Data sheet column headers. |
| `TITLE_COLUMNS` | `['title', 'name']` | Which column counts as a row's title for `title:` exclude keywords (first non-empty wins). |
| `CRITERIA_KEY_HEADER` / `CRITERIA_VALUE_HEADER` | `field` / `value` | Criteria sheet column headers. |
| `SKIP_COLUMNS_IN_PROMPT` | `status`, `grade`, `reasoning`, `id`, `scraped_at`, `graded_at` | Columns never sent to the LLM or checked by exclude keywords. Add anything you don't want leaving the sheet. |
| `MAX_FIELD_CHARS` | `600` | Per-field truncation in the prompt. Raise it if your rows have important long-form content. |
| `ENABLE_RUN_LOG` | `true` | Append a per-run summary to the `Log` sheet. |

**Statuses and grades**

| Key | Default | What it does |
| --- | --- | --- |
| `STATUSES_TO_GRADE` | `['new', 'regrade']` | Which `status` values a run picks up. Add your own for another entry point. |
| `STATUS_GRADED` | `graded` | Written to `status` after a row is graded. |
| `VALID_GRADES` | `A+` … `F` | The grading scale, best first. The prompt and the parser are both built from this list, so changing it here is the whole change (e.g. `['5', '4', '3', '2', '1']`). The **last** entry is what exclude-keyword rejects get. Update your rubric in the Criteria sheet to match. |

## What it doesn't do

On purpose, to keep the example clean:

- **No retry queue.** Errors get logged and the row keeps its current
  status, so the next run picks it up again. That's the whole retry strategy.
- **No fancy follow-up actions.** The source project also generates
  cover letters for A-grade matches and emails them. That's intentionally
  not here — once you have grades in your sheet, write your own
  `Apps Script` function that reads grade=A rows and does whatever you
  want with them.
- **No embeddings / RAG.** It's a single LLM call per row with truncated
  text. If you need semantic similarity as a pre-filter, that's a sensible
  thing to add (replace stage 1), but it's not in here.
- **No multi-model voting.** One model, one grade. If you need higher
  reliability, call this twice with different models and reconcile in
  your own code.
- **Not hardened against prompt injection.** Row content is fenced in the
  prompt and the model is told to treat it as data, not instructions, but
  that's a mitigation, not a guarantee. A determined adversarial row could
  still influence its own grade. Fine for triage; don't point it at a
  hostile data source and trust the grades blindly. Results only ever land
  in the sheet — the script takes no action on them.

## Development

The pure helpers (keyword parsing, response parsing, whole-word exclude
matching, row titles, config normalization) have a dependency-free Node test harness:

```
npm test        # or: node tests/run.js
```

It loads the real `Grader.js` into a sandbox with the Apps Script globals
stubbed, so the assertions run against the actual source. It also runs
`gradeNewRows` end to end against in-memory stand-ins for the Sheets and fetch
services, covering the happy path, a rejected key, a retired model and a row
that moves mid-run. Real Sheets and network behavior still only run bound to a
real spreadsheet.

## License

MIT. See [LICENSE](LICENSE).
