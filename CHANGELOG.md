# Version history

The version shown at the bottom of the site comes from `"version"` in `package.json`.
For each update: bump that number, add an entry at the top of this file, and push to `main`.

## v1.2 — November 16, 2026

- **Recognizes more column headings**, such as `GEN`, `M/F`, `Boy/Girl`, `Yrs`, `Angel Sheet #`, `Tag #` and `ID`. Capitalization and punctuation don't matter.
- **Recognizes a gender column from its values** (M/F, Male/Female, Boy/Girl) even when the heading is something else.
- **Ignores phone, address, parent, notes, table and volunteer columns**, and lists them as "Not printed on labels". A "Parent name" column is never mistaken for the child's name.
- **Shows which columns it used**, for example "GEN → gender".
- **One note for a missing column** instead of a warning for every angel (for example, "No gender column was found").

## v1.1 — November 16, 2026

- **Label text is centered top to bottom** on each label. Before, it started at the very top edge with the spare space left at the bottom.
- **Downloadable blank spreadsheet template** with the right column headings and a sheet of instructions.
- **"How to set up your spreadsheet" instructions** with an example row. They start collapsed.
- **Clearer upload box:** shows "✓ Using *filename*" with a "choose a different spreadsheet" link.
- **Church photo banner** at the top of the page (photo by Jim.henderson, CC BY 4.0, via Wikimedia Commons).
- **Wider page,** so all of the example column headings are visible.
- **Footer credit and version number.**

## v1.0 — November 11, 2026

- First release. Upload the Christmas Angel spreadsheet, choose a return deadline, and download one Word document of labels.
- Label layout matches the parish's existing December 6 document on Avery 5168 (4 per sheet). Also supports Avery 5164, 5126 and 5165.
- Checks the spreadsheet for duplicate or missing angel numbers and missing names, with Excel row numbers. Notes missing ages, genders and gifts.
- Leaves out blank age lines instead of printing "Age: None".
- Runs entirely in the browser; the page is blocked from sending data anywhere.
