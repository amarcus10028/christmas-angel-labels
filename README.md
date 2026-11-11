# Christmas Angel Label Maker

A free website for the Church of St. Ignatius Loyola. It turns the Christmas Angel spreadsheet into one print-ready Word document of gift labels.

**Volunteer workflow:** open the bookmarked page → drop in the spreadsheet → choose the return deadline → click **Generate Word Document** → print `Christmas_Angel_Labels_<year>.docx` on Avery 5168 labels.

Everything happens inside the browser. The spreadsheet is never uploaded: the published page carries a Content-Security-Policy with `connect-src 'none'`, so the browser blocks every network request the page might try to make. There are no accounts, no database, no analytics and no paid services.

## Supported spreadsheet

The app finds the heading row on its own (title or instruction rows above it are fine) and matches the columns by their heading text:

| Field | Headings it recognises | Required |
| --- | --- | --- |
| Angel number | `Angel #`, `Angel No.`, `#`, `Number` | Yes |
| Name | `first name` + `Last name`, or a single `Name` column | Yes |
| Gender | `gender`, `sex` (M/F become Male/Female; anything else prints as typed) | No |
| Age | `age` (printed exactly as typed, e.g. `12` or `adult`) | No |
| Gifts | every column whose heading contains `gift`, `choice`, `wish` or `request`, in left-to-right order. There is no limit on how many. | No |

Blank rows and lone notes such as "Total: 199" are ignored. Labels are sorted by angel number, never renumbered, and gaps such as #429 are allowed.

**Stops generation:** the file can't be read, no recipients are found, an angel number is duplicated or missing, a name is missing, or no deadline is chosen. Each message gives the Excel row number.
**Warnings only:** a missing age, gender or gift choices; text long enough that a label is printed smaller; gaps in the numbering (shown as a note).

## Label sheets

| Option | Label | Per sheet |
| --- | --- | --- |
| **Avery 5168 / 8168** (default) | 3½″ × 5″ | 4 |
| Avery 5164 / 8164 | 3⅓″ × 4″ | 6 |
| Avery 5126 / 8126 | 5½″ × 8½″ half sheet | 2 |
| Avery 5165 / 8165 | 8½″ × 11″ full sheet | 1 |

The 5168 geometry is copied from the parish's own document: 0.5″ top and side margins, label columns of 5055/630/5115 twips, 5″ rows. Each page is one borderless, fixed-width table. Rows have an exact height and cannot split across pages, so labels land in the same place on every page. The layout uses no tabs, spaces or manual page breaks. If a label's text is too long, only that label's type is made smaller so nothing gets cut off.

Smaller sheets such as the 30-up 5160 are left out on purpose, because a full label doesn't fit legibly on them. To add a size, add an entry to `src/lib/layouts.ts`. The tests check every layout automatically.

Parish name and return wording can be edited under **Parish Settings**. Those settings and the chosen label sheet are remembered in this browser only.

## Running it locally

Requires Node 22+.

```bash
npm ci
npm run dev        # http://localhost:5173
npm test           # automated checks
npm run build      # production build in dist/
npm run preview    # serve the production build (with the privacy policy active)
```

## Tests

`npm test` builds spreadsheets in memory, parses them, generates the `.docx`, then **reads the Word XML back** and compares every label with the spreadsheet: number, name, gender, age, each gift, and exactly one deadline line. It covers the December 6 and December 13 runs, missing ages (#379–#383), the #429 gap, duplicates, missing names and numbers, an empty sheet, an unreadable file, missing gifts, long names and gifts, extra gift columns, other numbering ranges, and the page geometry of every label sheet.

The test data is synthetic, because the real spreadsheet contains children's names and must never be committed. To also test against the real files, copy them into the git-ignored `fixtures-private/` folder:

```
fixtures-private/OLL.xlsx
fixtures-private/Christmas_Angel_Labels_2026_December_6.docx
```

With those present, the tests also confirm that the generated document matches the parish's December 6 document label by label. The one intended difference is that it no longer prints `Age: None`.

To write sample documents for checking by eye:

```bash
OUT=/some/folder npx vitest run scripts/sample-docs.test.ts
```

## Deploying

Pushing to `main` runs `.github/workflows/deploy.yml`: install, test, build, then publish `dist/` to GitHub Pages. To turn this on once, go to **Settings → Pages → Build and deployment → Source: GitHub Actions**. The site works at any Pages path, so no custom domain is needed.

**Updating:** change the code, run `npm test`, and push or merge to `main`. The site updates in about a minute.

**Making next year's labels:** nothing to change. Open the site, upload the new spreadsheet, and pick the new date.

## Dependencies

All are free for this use: React (MIT), `docx` (MIT), and SheetJS Community Edition (Apache-2.0). SheetJS is installed from its official CDN tarball because the npm registry copy is outdated. Versions are pinned exactly and locked in `package-lock.json`.
