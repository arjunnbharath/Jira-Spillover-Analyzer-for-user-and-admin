# Jira Spillover Analyzer — Dashboard

A **single-file, client-side** web app for analyzing Jira Excel exports, flagging **spillover** work by sprint values, tracking **bugs** (PROD / RCA), and exporting an enriched **multi-sheet workbook**—all in the browser, with **no server** and no data upload.

| | |
|---|---|
| **Entry point** | `velocity.html` (open in a modern browser) |
| **Stack** | HTML, CSS, JavaScript |
| **Libraries (CDN)** | [SheetJS (xlsx)](https://sheetjs.com/), [FileSaver](https://github.com/eligrey/FileSaver.js) |
| **Fonts** | Inter (Google Fonts) |

---

## Features

- **Upload & configure** — Load `.xlsx` / `.xls` (e.g. Jira export), set **base version** and **sprint threshold**, then **Proceed** to analyze.
- **Spillover rule** — A row is **SpillOver = Yes** when **any** “Sprint” column value **equals** the threshold (e.g. `14.1` matches `14.1`, not `14.2`). Float matching uses a small epsilon for Excel rounding.
- **Dashboard** — Summary metrics, **cumulative spillover chart**, and **issue type** breakdown (when an **Issue Type** column exists).
- **Assignees** — Sidebar list of assignees and issue keys; click to open **ticket detail** (full row from the sheet).
- **Spillover** — List of spillover issues; per-row **Spillover reason** and **category** (dropdown: Defect Blocker, Dependency, Scope Creep, Owner Specific, Estimation, Carryover Work). Exported on the **SpillOver** sheet only.
- **Bugs** — Bug rows with **PROD** (Yes/No) and **RCA** (text). **PROD** and **RCA** appear **only** on the **Bug** sheet in the export—not on TicketIntake or SpillOver.
- **Reports** — Sprint & export report: stats, column detection, issue mix, enrichment progress.
- **Analytics** — Spillover split, issue-type bars, top assignees.
- **Exports** — **Export preview**: every output sheet with **column chips** (app-added columns highlighted) and **scrollable table previews** (first rows) before download.
- **Search** — Quick search by issue key, assignee, summary (`/` keyboard shortcut when not typing in a field).
- **Export** — One-click **Download** / **Export** produces a structured Excel file (see below).

---

## How to run

1. Clone or download this repository.
2. Open **`velocity.html`** in **Chrome, Edge, or Firefox** (local file or any static host).
3. No `npm install`, no build step.

> **Note:** Browsers may restrict `file://` in some cases; if anything fails, serve the folder with a simple static server (e.g. `npx serve .`) and open the served URL.

---

## Workflow

1. **Choose file** — Pick your Jira / spreadsheet export.
2. **Base version** — e.g. `14.0` (used in sheet names and output filename).
3. **Sprint threshold** — Value that any Sprint cell must **match** for SpillOver Yes (e.g. `14.1`).
4. **Proceed** — Loads analysis, opens the main dashboard, and populates Spillover, Bugs, Reports, Analytics, and Export preview.
5. **Optional** — Fill spillover reasons/categories and bug PROD/RCA, then use **Export** or **Download Excel** from the **Exports** page or other headers.

---

## Output Excel file

Default filename pattern:

`SprintVelocity_{baseVersion}&{baseVersion+0.1}.xlsx`

### Sheets

| Sheet | Contents |
|--------|----------|
| **TicketIntake_{version}** | All data rows with original columns + **`SpillOver`** (Yes/No). **No PROD/RCA** here. |
| **SpillOver_{version}** | Only rows where **SpillOver = Yes**, plus **`SpillOverReason`** and **`SpillOverCategory`**. **No PROD/RCA** here. |
| **Bug_{version}** | Rows where **Issue Type = Bug**, same columns as intake for those rows + **`SpillOver`** + **`PROD`** + **`RCA`** (last two are bug-only). If there are no bugs, a **placeholder** row is still written so the sheet exists. |

**Sprint columns** are detected from the header row: columns whose header is **Sprint** (from a fixed starting column index in the parser—see `getSprintColumnsIndices` in the source).

---

## Project layout

```
.
├── velocity.html   # Full application (UI + logic + styles)
└── README.md       # This file
```

Other files in the folder (if any) are optional / unrelated unless documented separately.

---

## Privacy & security

- All processing runs **in your browser**.
- The file is read with the **File API**; nothing is sent to a server by this app.
- CDNs are used only for **xlsx**, **FileSaver**, and **Google Fonts** (see `<script>` / `<link>` tags in `velocity.html`).

---

## Browser support

- Modern browsers with **ES6+** (e.g. `const` / `let`, `Object.assign`, template usage as in the file).
- **JavaScript enabled** (required).

---

## Customization

- Styling: CSS variables and classes at the top of `velocity.html`.
- Behavior: search for functions such as `runProceed`, `processExcel`, `hasSpilloverRaw`, `computeExportPlan` in `velocity.html`.

---

## License

Add a `LICENSE` file in the repository if you need a formal license; this README does not specify one by default.

---

## Credits

- [SheetJS](https://sheetjs.com/) for Excel read/write.
- [FileSaver.js](https://github.com/eligrey/FileSaver.js) for client-side download.
