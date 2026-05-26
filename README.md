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
format. Just swap the endpoint, model name, and key.

This was extracted from a real project ([entertainment-jobs-tracker](https://github.com/consumedsoul/entertainment-jobs-tracker))
that uses the same pattern to rank LA entertainment crew job listings
against a working actor/producer's profile. The generic version here strips
out everything specific to that use case so you can point it at anything.

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
trust an 8B model's letter grade. Use this as a triage layer that decides
what's worth your own attention, not as the final word.

## How it works

```
Data sheet (status="new" rows)
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
- **Respects the Apps Script 6-minute limit.** Defaults to bailing at 5
  minutes; remaining rows stay `status="new"` and get processed on the
  next scheduled run.
- **Sleeps between API calls** (default 7s — tuned for Groq's free tier
  6K tokens/minute). Drop this if you're on a paid tier.
- **One retry on HTTP 429** with a 5-second backoff.
- **Truncates long fields** in the prompt so token use stays predictable
  even if someone pastes a 50KB description into a cell.

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
   and paste the contents of [Grader.js](Grader.js) into it.
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
    Pick whatever cadence you want (e.g. daily, every hour).

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
Anything matching gets `F` instantly with no API call. Use `title:` as
a prefix to only match a row's title field (useful when a body description
might mention the keyword in a benign way, like "we are not an unpaid
internship"). Example:

```
chain, fast casual, title:closed, title:permanently
```

## Schema

### Data sheet (you create this)

Three required columns, anywhere in the header row:

| column | what goes in it |
| --- | --- |
| `status` | `new` for rows to grade. Becomes `graded` after. |
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

## Configuration

Most things you'd want to change are in `GRADER_CONFIG` at the top of
[Grader.js](Grader.js):

- `API_ENDPOINT` / `API_MODEL` — switch providers or models here.
- `API_DELAY_MS` — sleep between calls. Lower on paid tiers.
- `TIMER_BUDGET_SEC` — when to bail before the 6-minute Apps Script
  limit. Default 300s.
- `VALID_GRADES` — change the grading scale (e.g. to a 1-5 numeric scale
  — also update the prompt format string in `buildGradingPrompt_`).
- `SKIP_COLUMNS_IN_PROMPT` — columns that won't be sent to the LLM
  (already metadata, not content).
- `MAX_FIELD_CHARS` — per-field truncation. Default 600. Bump up if
  your rows have important long-form content.

## What it doesn't do

On purpose, to keep the example clean:

- **No retry queue.** Errors get logged and the row stays `status="new"`
  for the next run to pick up. That's the whole retry strategy.
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

## License

MIT. See [LICENSE](LICENSE).
