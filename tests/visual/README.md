# Visual compatibility regression

This suite generates ten deterministic DOCX fixtures, converts them with
LibreOffice, renders the resulting PDFs with Poppler, and compares every page
with a committed PNG baseline.

## Coverage

| Fixture | Compatibility surface |
| --- | --- |
| `01-typography-and-styles` | Inline styling, paragraph spacing, borders |
| `02-lists-and-links` | Nested lists, bookmarks, internal/external links |
| `03-tables-and-spans` | Fixed tables, row spans, borders, cell styles |
| `04-sections-and-headers` | Section breaks, landscape pages, headers/footers |
| `05-annotations-and-math` | Checkboxes, comments, footnotes, Office Math |
| `06-revisions-and-textbox` | Tracked changes, text boxes, thematic breaks |
| `07-theme-and-preset` | Theme tokens and built-in presets |
| `08-callouts-and-data` | Callout and data-table plugins |
| `09-business-plugins` | Letterhead, timeline, and signature plugins |
| `10-images-and-columns` | Images, multi-column layout, column breaks |

## Commands

Build the workspace before generating fixtures:

```shell
pnpm run build
pnpm run visual:check
```

When an intentional visual change has been reviewed, regenerate the baselines:

```shell
pnpm run build
pnpm run visual:update
```

Never update baselines just to make CI green. Inspect every changed page and
record the reason in the pull request. Failed comparisons are written to
`output/visual-regression/` and uploaded by CI.

## Microsoft Word release check

LibreOffice regression is automated. Before a tagged release, open all
generated files from `tmp/visual-regression/docx/` in a currently supported
desktop version of Microsoft Word and record:

| Check | Expected |
| --- | --- |
| Page count and orientation | Match the fixture inventory |
| Headers, footers, and section transitions | Correct variant and position |
| Tables and columns | No clipping, overlap, or broken spans |
| Comments, footnotes, math, and revisions | Native editable Word objects |
| Images and plugin output | Sharp, aligned, and legible |
| Fields | Refresh successfully when Word requests an update |

The release is blocked by unexplained clipping, overlap, missing content,
corrupt-file warnings, or a page-count change.
