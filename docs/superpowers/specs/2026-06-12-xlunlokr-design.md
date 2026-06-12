# xlunlokr — Design

**Date:** 2026-06-12
**Status:** Approved

## Purpose

A fully client-side web application that removes sheet and workbook protection
from Excel files. Users drop one or more files into the page and download
unlocked copies. No server-side code, no external APIs — files never leave the
browser.

## Scope

**In scope (v1):**

- Remove `<sheetProtection>` elements from all sheet parts (worksheets,
  chartsheets, dialogsheets, macrosheets).
- Remove `<workbookProtection>` from `xl/workbook.xml`.
- Supported input formats: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` (all OPC ZIP
  containers; treated identically, original extension preserved on output).
- Multiple files per drop; each processed independently.

**Out of scope:**

- Open-password encryption (files are truly encrypted; detected and explained,
  not unlocked).
- Legacy binary `.xls` (detected and explained).
- `<fileSharing>` / modify-password removal.
- VBA project password removal.

## Approach decision

An `.xlsx` file is a ZIP of XML parts, so the unlock operation is: unzip,
strip protection tags, re-zip. Three approaches were considered:

1. **fflate** (chosen) — tiny (~8 KB gzipped), fast, battle-tested ZIP
   library. One runtime dependency, minimal code.
2. Zero-dep native `CompressionStream` + hand-rolled ZIP container parsing —
   no dependencies but ~200 lines of fragile binary-format code to own.
3. JSZip — works, but larger and slower than fflate with no benefit.

XML modification uses **targeted string surgery** (regex removal of the
specific elements), not `DOMParser` re-serialization. Re-serializing can
rewrite namespaces and formatting document-wide; surgical removal guarantees
only the targeted bytes change. Both `sheetProtection` and
`workbookProtection` are childless elements in the OOXML schema
(CT_SheetProtection / CT_WorkbookProtection), so element removal is safe. The
regex handles both self-closing (`<sheetProtection …/>`) and paired
(`<sheetProtection …></sheetProtection>`) forms.

## Architecture

**Stack:** Vite + vanilla TypeScript. Single page. One runtime dependency:
fflate. No framework, no Web Worker (processing takes tens of milliseconds).

### Core module — `src/unlock.ts`

Pure function, no DOM dependencies, testable in Node:

```ts
unlockWorkbook(data: Uint8Array): UnlockResult
```

`UnlockResult` is a discriminated union:

- `{ kind: "unlocked", data, sheetProtectionsRemoved, workbookProtectionsRemoved }`
- `{ kind: "already-unlocked" }` — valid workbook, zero protection tags found
- `{ kind: "error", reason: "encrypted-or-legacy" | "not-excel" | "invalid-zip" }`

Data flow per file:

1. **Sniff magic bytes.** `PK` (50 4B) → ZIP, proceed. `D0 CF 11 E0` →
   OLE/CFB container (legacy `.xls` or open-password-encrypted workbook) →
   `encrypted-or-legacy` error. Anything else → `not-excel` error.
2. **Unzip** with fflate. Unparseable ZIP → `invalid-zip` error.
3. **Strip tags** from every entry matching `xl/**/*.xml`: decode UTF-8,
   remove `sheetProtection` and `workbookProtection` elements, count
   removals. Entries without matches pass through byte-identical.
4. **Re-zip** all entries and return the result.

### UI layer — `src/main.ts` + small helpers

- Drop zone (drag-and-drop or click to pick; `multiple` enabled).
- File list state: one row per file with name, status, and per-file download
  button (Blob URL).
- Output filename: `<original>-unlocked.<ext>`.
- Talks to the core module only through `unlockWorkbook`.

## UI design

Clean minimal utility, one screen:

- Title + one-line explanation of what the tool does.
- Prominent privacy note: "files never leave your browser".
- Large drop zone.
- Results list, one row per file:
  - ✅ **Unlocked** — "removed N sheet protections, M workbook protections" +
    download button.
  - ℹ️ **Already unlocked** — no protection found; no download button.
  - ❌ **Error** — per-file message; one bad file never blocks others.
- "Download all" button when 2+ files succeeded. It triggers each file's
  individual download in sequence (no ZIP bundling — users expect workbook
  files back, not an archive).

## Error handling

| Condition | Detection | User-facing behavior |
|---|---|---|
| Open-password encrypted | CFB magic bytes | "Encrypted with an open password (or a legacy .xls) — this tool can't unlock it." |
| Legacy `.xls` | CFB magic bytes | Same message as above (both are CFB containers). |
| Not an Excel/ZIP file | No known magic | "Not an Excel file." |
| Corrupt ZIP | fflate throws | "File appears corrupted." |
| No protection present | Zero tags removed | Info row: "already unlocked", no download. |

Errors are per-file; processing of other files continues.

## Testing

**Vitest**, unit tests against `unlockWorkbook`:

- Strips `sheetProtection` across multiple worksheets.
- Strips `workbookProtection` from `workbook.xml`.
- Handles paired-form (non-self-closing) protection elements.
- Leaves all other ZIP entries byte-identical.
- Returns `already-unlocked` for clean files.
- Returns `encrypted-or-legacy` for CFB bytes.
- Returns `not-excel` for arbitrary bytes.

Fixtures are **generated programmatically** by a test helper that zips
handcrafted XML into minimal locked workbooks — no binary fixtures committed.

One integration test globs `fixtures/*.xlsx` (real locked files kept locally,
e.g. `lp-wiskunde.xlsx`) and auto-skips when the directory is empty or absent,
so it runs locally but never requires the files in CI.

## Repository & deployment

- Public GitHub repo `xlunlokr`, created via `gh`.
- `.gitignore` includes `fixtures/` and `.idea/` so real locked files never
  enter version control.
- GitHub Actions workflow on push to `main`:
  1. Install dependencies, run Vitest.
  2. Build with Vite, `base: '/xlunlokr/'`.
  3. Deploy via official `actions/configure-pages` →
     `actions/upload-pages-artifact` → `actions/deploy-pages` flow.
