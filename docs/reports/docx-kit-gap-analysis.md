# docx-kit vs docxjs (dolanmiu/docx) — Feature Gap Analysis

> Updated: 2026-07-26 | Baseline: docx-kit `main` at `92de980` | docxjs v9.7.1

## Executive Summary

The feature gaps tracked in the original v0.3.0 audit are now closed. The
current release candidate exposes the underlying Word capabilities through
typed configuration, DSL nodes, builder methods, or focused built-in plugins.

| Original priority | Closed | Remaining |
| ----------------- | -----: | --------: |
| P0 — Critical     |      4 |         0 |
| P1 — High         |      8 |         0 |
| P2 — Medium       |     12 |         0 |
| P3 — Low/Niche    |      8 |         0 |
| **Total**         | **32** |     **0** |

This does not mean that every constructor option exported by `docx` has a
dedicated CSS-like property. The `style.docx` escape hatch remains available
for uncommon upstream options. It means every product gap explicitly recorded
by this audit now has a supported public path.

## Closed Capability Gaps

### Document Structure

| Capability                        | Public surface                                     | Status |
| --------------------------------- | -------------------------------------------------- | ------ |
| Multiple sections                 | `.section(config)`                                 | ✅     |
| Rich headers and footers          | `HeaderFooterContent.children` accepts block nodes | ✅     |
| First, even, and default variants | `HeaderFooterConfig`                               | ✅     |
| Header/footer distance and gutter | `PageConfig`                                       | ✅     |
| Page borders                      | `PageConfig.borders`                               | ✅     |
| Multi-column sections             | `SectionConfig.columns`                            | ✅     |
| Column breaks                     | `.columnBreak()`                                   | ✅     |
| Line numbering                    | `SectionConfig.lineNumbers`                        | ✅     |
| Page-number restart and format    | `PageConfig.pageNumber`                            | ✅     |
| Section break types               | `SectionConfig.type`                               | ✅     |
| Page-break-before behavior        | `style.pageBreakBefore`                            | ✅     |

### Content and Semantic Nodes

| Capability                   | Public surface                        | Status |
| ---------------------------- | ------------------------------------- | ------ |
| Bulleted and numbered lists  | `.bulletList()` / `.numberedList()`   | ✅     |
| Nested numbering             | Per-list and per-item levels (0–8)    | ✅     |
| External hyperlinks          | `.hyperlink()`                        | ✅     |
| Internal links and bookmarks | `.internalLink()` / `.bookmark()`     | ✅     |
| Table of contents            | `@docxkit/plugin-toc`                 | ✅     |
| Horizontal rules             | `.thematicBreak()` and divider plugin | ✅     |
| Positioned text boxes        | `.textBox()`                          | ✅     |
| Tab stops                    | `style.tabStops`                      | ✅     |
| Footnotes                    | `.footnote()`                         | ✅     |
| Comments                     | `.comment()`                          | ✅     |
| Checkboxes                   | `.checkbox()`                         | ✅     |
| Office Math                  | `.math()` with typed expressions      | ✅     |
| Tracked insertions/deletions | `.insertedText()` / `.deletedText()`  | ✅     |

### Tables

| Capability                          | Public surface                               | Status |
| ----------------------------------- | -------------------------------------------- | ------ |
| Row and column spans                | `rowSpan`, `colSpan`, and per-row span hints | ✅     |
| Cell shading and vertical alignment | Cell style rules                             | ✅     |
| Per-cell styles                     | Static styles and `TableCellStyleResolver`   | ✅     |
| Native table styles and look flags  | `styleName` / `tableLook`                    | ✅     |
| Explicit table layout               | `layout: 'autofit' \| 'fixed'`               | ✅     |
| Outer and inner borders             | `TableBordersConfig`                         | ✅     |
| Floating and side-by-side tables    | `TableFloatingOptions`                       | ✅     |
| Right-to-left table grids           | `visuallyRightToLeft`                        | ✅     |

### Text and Paragraph Styling

| Capability                            | Public surface                           | Status |
| ------------------------------------- | ---------------------------------------- | ------ |
| Highlight, superscript, and subscript | Text style rules                         | ✅     |
| Small caps and double strike          | Text style rules                         | ✅     |
| RTL runs                              | `rightToLeft`                            | ✅     |
| Character and CSS-unit letter spacing | `characterSpacing` / `letterSpacing`     | ✅     |
| Emboss and imprint effects            | Text style rules                         | ✅     |
| Keep-lines and keep-next              | Paragraph style rules                    | ✅     |
| Widow control and outline level       | Paragraph style rules                    | ✅     |
| Paragraph borders                     | Side-specific and shorthand border rules | ✅     |

### Assets, Metadata, and Output

| Capability                        | Public surface                                  | Status |
| --------------------------------- | ----------------------------------------------- | ------ |
| Core OOXML metadata               | `DocxKitConfig.metadata`                        | ✅     |
| Custom document properties        | `metadata.customProperties`                     | ✅     |
| Embedded TrueType/OpenType fonts  | `DocxKitConfig.fonts`                           | ✅     |
| Browser Blob output               | `.toBlob()`                                     | ✅     |
| Cross-platform byte/base64 output | `.toUint8Array()`, `.toBuffer()`, `.toBase64()` | ✅     |
| Node filesystem save              | Node entry `save` method / `saveDocument()`     | ✅     |
| Node streaming                    | `.toStream()` / `streamDocument()`              | ✅     |

### Built-in Plugins

The workspace now ships 19 built-in plugins:

`badge`, `barcode`, `callout`, `changelog`, `codeBlock`, `coverPage`,
`dataTable`, `divider`, `echarts`, `invoice`, `letterhead`, `meetingMinutes`,
`pageNumber`, `propertyTable`, `qrcode`, `signatureBlock`, `timeline`, `toc`,
and `watermark`.

The final plugin gap from the original audit, linear barcode generation, is
implemented through `@docxkit/plugin-barcode`. It supports Node and browser PNG
renderers through the optional `bwip-js` peer dependency.

## Platform Boundaries

These are intentional platform contracts, not open feature gaps:

| Surface                     | Browser                                                  | Node.js                       |
| --------------------------- | -------------------------------------------------------- | ----------------------------- |
| Core document generation    | ✅                                                       | ✅                            |
| Barcode rendering           | Canvas renderer                                          | Buffer renderer               |
| ECharts rendering           | Primary supported target                                 | Requires a server-side canvas |
| Filesystem save and streams | —                                                        | ✅                            |
| TOC and other Word fields   | Generated as fields; the editor may need to refresh them |

Media integrations keep their renderers as optional peer dependencies:
`bwip-js`, `echarts`, `highlight.js`, and `qrcode`. Consumers install only the
renderers they use.

## Verification Baseline

The gap-closure branch is protected by the same commands used for release:

| Gate                           | Result                                           |
| ------------------------------ | ------------------------------------------------ |
| Workspace build                | 35 non-docs workspaces passed                    |
| TypeScript                     | All workspace packages passed `tsc --noEmit`     |
| Tests                          | 76 files, 778 tests passed                       |
| Documentation examples         | 146 Markdown files validated                     |
| Documentation production build | VitePress build passed                           |
| Generated Monaco declarations  | Deterministic and complete for all 19 plugins    |
| Package artifact audit         | Every published entry point resolved             |
| Packed consumer smoke test     | Runtime imports and downstream TypeScript passed |
| CI platforms                   | Ubuntu and Windows                               |

## Quality Infrastructure

The quality-and-scale roadmap is now implemented:

1. LibreOffice/Poppler visual regression covers 10 DOCX fixtures and 12 pages,
   with a documented Microsoft Word release matrix.
2. Node and real-Chrome benchmarks cover paragraph, table, image, and
   plugin-heavy documents at small through large sizes.
3. The documentation playground lazy-loads Monaco, docx-kit, Sucrase, and the
   preview renderer behind user actions, with a bundle-size gate.
4. A third-party plugin fixture is built from TypeScript, discovered through
   registry metadata, validated by manifest, loaded locally, registered, and
   rendered into asserted OOXML.

These checks supplement the release verification workflow; performance remains
a non-blocking trend signal.

## docx-kit Advantages over Raw docxjs

| Feature                            | Value                                                             |
| ---------------------------------- | ----------------------------------------------------------------- |
| CSS-like styles                    | Familiar styling vocabulary with unit conversion                  |
| Type-safe builder                  | Fluent document construction with typed nodes                     |
| Plugin architecture                | Extensible `definePlugin()` / `.use()` / `.plugin()` model        |
| Style inheritance and theme tokens | Reusable design systems for documents                             |
| AI-friendly JSON DSL               | Declarative generation through `renderDocx()`                     |
| Platform entry points              | Explicit browser and Node contracts                               |
| Package validation                 | Generated declarations, artifact audits, and consumer smoke tests |
| Escape hatch                       | Direct upstream options remain available through `style.docx`     |
