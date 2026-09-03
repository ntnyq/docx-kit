# Visual compatibility regression

This suite generates ten deterministic DOCX fixtures, converts them with
LibreOffice, renders the resulting PDFs with Poppler, and compares every page
with a committed PNG baseline.

The default renderer runs in Docker on `linux/amd64`, using the Ubuntu image
digest and package snapshot pinned in
`scripts/visual-regression/docker/Dockerfile`. This keeps LibreOffice, Poppler,
font substitution, and font metrics identical on developer machines and CI.
Docker must be running; the first invocation builds the image and later runs
reuse its cached layers. Rendering itself runs without network access.
The image includes LibreOffice Math for Office Math objects, Carlito for Calibri
fallback, and an explicit Liberation Serif fallback for the Ocean theme's
`Georgia, serif` family.

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
`output/visual-regression/` and uploaded by CI, including the actual page PNGs,
diffs, and renderer version report.

When updating the renderer image or package snapshot, review all regenerated
pages and commit the Dockerfile and baselines together. Page count, nonblank-page
checks, and the 1% pixel-difference limit remain enforced.

### Baseline review: 2026-09-03

The baselines include the document-style fixes from `b8dd1e1`. All twelve
pages were compared with the previous baselines using the unchanged pinned
renderer (`docx-kit-visual:2a5e50ea2b30`, LibreOffice 24.2.7.2). The reviewed
changes are:

- Borders use eighth-points instead of twips, producing the requested widths
  for heading rules, paragraph boxes, theme accents, and table borders.
- Text defaults now reach headings, lists, links, semantic nodes, table cells,
  and header/footer strings, making their typography consistent with the
  configured font and size.
- Table headers honor their configured font, color, and weight.
- Explicit line-spacing rules preserve numeric multipliers and absolute
  heights, with corresponding paragraph and theme-spacing changes.

The ten fixtures still produce twelve pages. No new clipping or missing
content was found. The existing page-count, nonblank-page, and 1% difference
checks remain in place.

To investigate a locally installed LibreOffice version, use:

```shell
pnpm run visual:generate
pnpm run visual:render --system
```

`LIBREOFFICE_BIN`, `PDFTOPPM_BIN`, and `PDFINFO_BIN` override the native binaries
only in this diagnostic mode. Native renders may differ from the pinned
environment and cannot update the committed baselines.

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
