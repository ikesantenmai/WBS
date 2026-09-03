# wbsgen — WBS Generator

*[日本語](README.md)*

Creates an **empty WBS (Gantt chart sheet)** in Excel for a given period, and
**shows a filled-in Excel file back as a Gantt chart**.

The layout — the month and week headings, the colours, the columns — is taken
from an existing Japanese WBS tool and reproduced as-is.

Japanese and English are both available. **Japanese is the default.**

```
wbsgen new --start 2026-04-01 --months 12 --lang en -o FY2026.xlsx
wbsgen serve                                    # create / view in a browser
```

How it is meant to be used:

1. Decide a period and hand out an empty WBS (`wbsgen new`, or "Download Excel")
2. Everyone fills in tasks, dates and progress in Excel
3. "Import Excel" in the browser turns it into a Gantt chart
4. "Export with Gantt chart" gives you an Excel file with the bars drawn in

---

## What you get

An Excel file with three sheets.

- **Schedule** — the table headings, the calendar and blank rows (frozen at `S4`)
- **Members** — the owners you listed, plus blank rows to add more
- **Settings** — the period, unit, working days and the holidays in that period

The Schedule sheet has the same columns as the original tool.

| Col | Contents | | Col | Contents |
|-----|----------|-|-----|----------|
| B | Group | | K | Actual end |
| C | Sub-group | | L | Delay |
| D | No. | | M | Progress |
| E | Task | | N | Effort |
| F | Planned start | | O | Predecessor |
| G | Planned days | | Q | Owner |
| H | Planned end | | R | Status |
| I | Actual start | | S… | Calendar |
| J | Actual days | | | |

Blank rows carry borders and number formats (dates as `m/dd`, days as `0 d`,
progress as `0%`), so you can open the file and type straight into it. The
planned columns are tinted.

### The calendar heading

The heading is two rows, as in the original file. The top row is not merged: a
date is placed only where the grouping changes, and a number format makes it
read as "Apr".

| Unit | Top row | Bottom row |
|------|---------|------------|
| Daily | Month `mmm` | Day `d` plus a weekday row |
| Weekly | Month `mmm` | Week start `m/d` |
| Monthly | Year `yyyy` | Month `mmm` |

Week columns step **seven days from the start date**; they are not snapped to a
weekday. Move the start date if you want the weeks to break elsewhere.

In the daily view, Saturdays, Sundays and holidays are shaded.

---

## Language

The switch at the top right moves between Japanese and English. **Japanese is
the default**, and the choice is remembered in the browser. On the command line,
pass `--lang en`.

The language applies to all of:

- the interface, the date headings (4月 / Apr) and weekdays (月 / Mon)
- the Excel file: sheet names (スケジュール / Schedule), column headings,
  number formats (`m"月"` / `mmm`), and the settings sheet
- error messages, and the command-line help and output

**Importing does not care about the language.** A file written in Japanese
opens fine in English and the other way round: columns are found by their
heading text, and both languages' wordings are known. Status colours work for
both "完了 / Done" and "実行中 / In progress".

---

## Installing

```bash
pip install -e .          # command line only (openpyxl)
pip install -e '.[web]'   # with the web application (adds FastAPI / uvicorn)
```

Python 3.9 or later.

---

## Commands

```
wbsgen new --start YYYY-MM-DD [--end YYYY-MM-DD | --period-days N | --months N]
           [-o out.xlsx] [--unit day|week|month] [--rows N] [--title NAME]
           [--member OWNER] [--workdays mon,tue,wed,thu,fri] [--holiday YYYY-MM-DD]
           [--no-jp-holidays] [--lang ja|en]
wbsgen serve [--host 127.0.0.1] [--port 8000] [--reload] [--lang ja|en]
```

| Option | Meaning |
|--------|---------|
| `--start` | start date (required) |
| `--end` / `--period-days` / `--months` | the period (12 months by default) |
| `--unit` | calendar unit (`week` by default) |
| `--rows` | number of blank rows (40 by default) |
| `--title` | project name |
| `--member` | add an owner to the members sheet (repeatable) |
| `--workdays` | working weekdays (`mon,tue,wed,thu,fri` by default) |
| `--holiday` | add a non-working day (repeatable) |
| `--no-jp-holidays` | do not treat Japanese public holidays as days off |
| `--lang` | `ja` / `en` (`ja` by default) |

```bash
# A financial year, weekly, 40 blank rows
wbsgen new --start 2026-04-01 --end 2027-03-31 --lang en -o FY2026.xlsx

# Six months, daily, 80 blank rows
wbsgen new --start 2026-04-01 --months 6 --unit day --rows 80 --lang en -o H1.xlsx

# Saturdays are working days; the new year break is off
wbsgen new --start 2026-04-01 --workdays mon,tue,wed,thu,fri,sat \
           --holiday 2026-12-30 --holiday 2026-12-31 --lang en -o wbs.xlsx
```

`serve` uses the `PORT` environment variable when it is set and listens on
`0.0.0.0`, so on Render and similar platforms `wbsgen serve` alone is enough.

---

## The web application

```bash
wbsgen serve                       # http://127.0.0.1:8000
wbsgen serve --host 0.0.0.0 --port 8080
```

### Creating an empty WBS

The period, sheet and working-day settings are on the left; a preview with
**the same columns and headings as the Excel file** is on the right. The
calendar is built by exactly the same code as the command line, so you can
check it before downloading.

### While it is working

Importing and exporting take a moment, so a line at the top says what is
happening ("Reading sbi.xlsx…", "Building the Excel file with the Gantt
chart…"). The buttons are held while it runs, so nothing starts twice.

### On a phone

The layout follows the width of the screen. It is the **same URL** — there is
no app to install.

- **760px and narrower** (a phone held upright): there is no room to put the
  settings next to the calendar, so the **Settings / Preview** tabs at the top
  switch between them. Buttons are sized for a finger, and the fields use a
  type size iOS will not zoom into.
- **900px and narrower** (a phone on its side, or a small window): the full
  table would leave no room for the calendar, so a **Columns** control appears
  in the toolbar.

| Columns | What you get |
|---------|--------------|
| **Name only** (the default when narrow) | Just the task name, pinned to the left while you swipe the calendar sideways |
| **Key columns** | Name, planned and actual dates, progress and status |
| **All** | The same 16 columns as the Excel file |

On a wide screen nothing changes: the settings sit next to the table and every
column is shown.

### Viewing a filled-in file as a Gantt chart

"Import Excel" takes a WBS made by this tool and filled in, and shows the table
and the Gantt chart side by side.

- **Planned** and **actual** bars, one above the other, coloured per owner
- Progress is painted inside the planned bar
- Unfinished actual bars are drawn semi-transparent
- A dashed red line marks **today**
- Status (Done / In progress / Delayed …) keeps the original colours
- Hovering a bar shows the plan, actuals, progress, owner and status
- The **unit** can be switched between daily, weekly and monthly

Columns are found by their heading text, so files with extra or reordered
columns still read. The period, working days and holidays come from the
Settings sheet and the owners from the Members sheet.

Because these sheets are filled in by hand, the following are also accepted.

- Marks that mean "nothing here" (`-` `—` `未` `なし` `N/A` …) count as empty
- Full-width digits and signs (`８０％`, `２０２６/４/１０`) are folded to ASCII
- Dates as `2026-04-10`, `2026/4/10`, `2026.4.10`, `2026年4月10日`, `4/10`
- A date left as a plain number (an Excel serial) is read as a date
- Days as `10`, `10日`, `10 days`, `10d`
- Progress as `0.8`, `80%`, `80`

**Every sheet that was there on import is there on export.** The export does
not build a new file: it **starts from the original and replaces only the
sheets it rewrites under the same name**.

| Sheet | What the export does |
|-------|----------------------|
| Schedule / Members / Settings | Rebuilt under the same name |
| Workload sheets for the months it generates | Rebuilt under the same name |
| Everything else | **Left completely untouched** |

"Everything else" covers notes, issue lists, cover pages, a schedule sheet you
renamed, a hand-made workload sheet for a month outside the plan, and hidden
sheets. They are never read, so charts, images, conditional formatting, data
validation, autofilters, comments, print areas and tab colours all survive. The
sheet order is preserved too (the Gantt shapes find the schedule sheet wherever
it sits).

**A renamed schedule sheet is still found.** If none of the known sheet names
match, the first sheet with a Task column is used.

**Indentation in a task name is kept.** Names are often indented with leading
spaces (full-width or ASCII) to show the hierarchy, so no whitespace is
trimmed: the name reaches the screen and the exported Excel exactly as it was
written. A cell holding only spaces still counts as empty.

**Vertically merged cells are read too.** Excel keeps a merged cell's value
only in its top-left cell, but shows it across the whole range, so the rows
below read the same value (dates and group names are often merged this way).

**A cell that cannot be read does not discard the row.** Only that field is
left empty, and the row and column are reported on screen, so a single typo
never makes a whole row disappear.

**You are also told *why* a cell looks empty.** When a date or a progress
value you filled in does not show up, it is usually one of these two. Both are
reported on screen when the file is read.

- **The formula result is not stored in the file.** Excel saves the result of
  a formula alongside the formula itself, but files written by a script do
  not. Such a cell reads as empty, so an actual end date written as, say,
  `=I5+20` never marks the row Done. **Open the file in Excel and save it
  again** and the result is stored, after which it reads fine.
- **A heading was changed, so the column is not found.** Renaming "End" to
  something else makes that whole column read as empty. The names of the
  columns that could not be found are listed.

#### How days and progress are decided

Values are taken from what is filled in, as follows. Derived values are shown
in a lighter font, with the reason on hover.

**The number written in a Days cell is never used.** It is cleared, then
counted again from the start and end dates. This applies to both plan and
actual.

| What is filled in | What follows |
|-------------------|--------------|
| Planned start and end | **Days are counted again from those two** (working days, both ends included) |
| Planned start and days only | The end date is derived from the days |
| No planned start | **Days are left empty** (there is nothing to count from) |
| Actual start and end | **Actual days are counted again from those two** |
| Actual start and days only (no end) | **Actual days are cleared**; no end date is derived |
| No actual start | **Actual days are left empty** |
| An actual end date | **Progress becomes 100% and the status becomes Done**; this wins over written values |
| Planned dates | **The status and delay are calculated** (see below) |

An end date is derived from the days **for the plan only**. "Start plus n
days" is the normal way to write a plan, and counting back from the derived
end gives the same number.

**Actuals never derive one.** Days left on work that has not finished are a
mistake in the source data, so they are thrown away. Inventing an end date
there would make both the bar and the progress look further along than they
are. The actual start date is kept, so the row still reads as started
(status "Remaining n d" and so on). The actual bar is not drawn until an end
date is filled in.

**A row with an actual end date written in is always 100% and Done.** That
holds even if it finished after the planned end, even with no planned dates at
all, and whatever progress or status the file says. A row with an empty name is
kept too, as long as it has actual dates.

Only an actual end date that is *written in* counts as finished. An end date
**derived** from the actual days does not (otherwise a task in progress would
silently become complete). For the same reason, a derived actual end date is
not written back into the cell when exporting.

If a row with an end date does not become Done, that cell is being **read as
empty**. The reason is listed in the warnings shown when the file is read (no
stored formula result, a heading that does not match, or contents that cannot
be read as a date).

#### Status

The status is worked out from a base date (today by default) and written into
the Status column. The first matching rule wins.

| Condition | Status |
|-----------|--------|
| An actual end date is filled in | **Done** |
| Past the planned start with no actual start, or past the planned end unfinished | **Delayed n d** |
| Started (an actual start date is filled in) | **Remaining n d** — until the planned end |
| Not started | **Starts in n d** — until the planned start |

Counts are in working days, and the day itself counts as zero (if today is the
planned end date, it reads "Remaining 0 d"). The delay is also written into the
Delay column.

An actual end date makes a row Done even with no planned dates. Otherwise a
row with no planned dates cannot be judged, so whatever status was written
stays. The colours match the original file (done grey, delayed pink,
remaining orange, upcoming light blue).

### Exporting with the Gantt chart

"**Export with Gantt chart**" writes an Excel file with the bars drawn in.

- Bars are **floating shapes (DrawingML)**, so their ends sit inside a cell in
  the weekly and monthly views — the same approach the original tool used
- Planned, actual and progress bars and the current-date line are placed
  exactly where they are on screen
- A row that had no end date gets the derived value written in
- The file is exported **in the unit you are viewing**, so pick daily, weekly
  or monthly first

The exported file can be imported again; the tests check that the contents,
period and totals round-trip.

### The workload check

The exported file carries a **workload sheet for each month**. It counts, day
by day, the planned periods of every row whose Owner column names the person,
so you can see at a glance who is free and who is stacked up.

|  | What it holds |
|---|---------------|
| Rows | The names that appear in the Owner column, in the order they appear |
| Columns | Every day the plan covers, one sheet per month |
| Values | How many tasks are assigned to that person on that day |
| Totals | Working days / days with tasks / free days / which days are free |

The colours match the original sheet (**red** = a working day with no task,
**green** = 1-2, **orange** = 3 or more, **grey** = a non-working day). They are
conditional formatting, so they follow along if you edit the numbers. Working
days come from the Settings sheet.

An Owner cell often names several people ("Yoshida/Kikuchi", "Sasaki (Takase)"),
so it is split into individual names. Counting then asks whether the name is
*contained* in the cell, so "Sasaki (Takase)" counts for both.

No sheet is added when nobody is named. Very long plans stop after 24 months.

### API

| Method | Path | Purpose |
|--------|------|---------|
| `GET` | `/api/meta` | choices (units, weekdays, defaults) |
| `POST` | `/api/preview` | the calendar headings and dimensions |
| `POST` | `/api/build` | an Excel file |
| `POST` | `/api/import` | read a filled-in Excel file, return the chart model |
| `POST` | `/api/export` | read a filled-in Excel file, return it with the chart |

The body is JSON in this shape. Anything wrong comes back as `422` with a
message.

```json
{
  "title": "FY2026 Schedule",
  "start": "2026-04-01",
  "end": "2027-03-31",
  "unit": "week",
  "rows": 40,
  "members": ["Design", "Build"],
  "workdays": ["mon", "tue", "wed", "thu", "fri"],
  "japanese_holidays": true,
  "holidays": ["2026-12-30"],
  "language": "en"
}
```

`period_days` or `months` work in place of `end`.

`/api/import` and `/api/export` take the file as `file` in a
`multipart/form-data` body. `?unit=day` overrides the unit written in the file.

Every endpoint accepts `?lang=en` (Japanese when omitted); `/api/build` also
reads `language` from the body.

```bash
curl -X POST "http://127.0.0.1:8000/api/export?lang=en&unit=month" \
     -F file=@filled.xlsx -o gantt.xlsx
```

---

## Development

```bash
pip install -e '.[web]' pytest
pytest -q
```

`tests/test_web_ui.py` runs only where Playwright and Chromium are available
and is skipped otherwise; it checks that the JavaScript actually runs.

### Layout

```
src/wbsgen/
  i18n.py        Japanese and English wordings and formats
  blank.py       the sheet spec (period, unit, blank rows, working days, language)
  timeline.py    the calendar axis and its two-row heading
  workcal.py     working-day rules and Japanese public holidays
  workload.py    the workload check per member
  style.py       colours and formats (taken from the original file)
  workbook.py    writing Excel
  importer.py    reading a filled-in Excel file
  chart.py       the Gantt chart model (for the screen)
  drawing.py     DrawingML (the bars drawn into Excel)
  inject.py      adding the drawing part to the xlsx
  cli.py         command line
  web/
    app.py       FastAPI endpoints
    static/      the interface (index.html / app.js / style.css)
```

The interface is plain HTML and JavaScript — no build step, no external CDN.

---

## History

Earlier versions scheduled tasks from their predecessors and computed delays
and statuses. That was removed when the tool was narrowed to "only create the
empty one" (it is still in commit `4735ef1`).

Empty sheets contain no shapes. Shapes appear only in files produced by
importing a filled-in sheet and choosing "Export with Gantt chart".
