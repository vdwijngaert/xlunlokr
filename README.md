# xlunlokr 🔓

Remove sheet and workbook protection from Excel files — entirely in your browser.

**Use it here: https://vdwijngaert.github.io/xlunlokr/**

## How it works

An `.xlsx` file is a ZIP archive of XML parts. Excel's "protect sheet" and
"protect workbook" features just add `<sheetProtection>` / `<workbookProtection>`
elements (with a password *hash*) to that XML — the data itself is not
encrypted. xlunlokr unzips the file in your browser, removes those elements,
re-zips it, and hands the unlocked copy back. Nothing is uploaded anywhere:
there is no server.

Supported formats: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`.

## What it can't do

- **Open-password encryption** — files that require a password to *open* are
  truly encrypted; xlunlokr detects them and tells you, but cannot unlock them.
- **Legacy `.xls`** — the old binary format is not supported.
- **VBA project passwords** — out of scope.

Use it only on files you're allowed to modify.

## Development

```bash
npm install
npm run dev    # local dev server
npm test       # vitest suite
npm run build  # production build to dist/
```

Real locked workbooks can be placed in `fixtures/` (git-ignored) — the
integration test picks them up automatically and skips when none are present.

Deployed to GitHub Pages by `.github/workflows/deploy.yml` on every push to
`main`.
