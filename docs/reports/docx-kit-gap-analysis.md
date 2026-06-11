# docx-kit vs docxjs (dolanmiu/docx) — Feature Gap Analysis

> Generated: 2026-06-11 | docx-kit v? | docxjs v9.7.1

---

## 1. Executive Summary

`docx-kit` wraps the `docx` (docxjs) library, adding a CSS-like style system, plugin architecture, and type-safe builder API. This analysis identifies gaps between docx-kit's exposed surface and what the underlying `docx` library actually supports, and recommends which gaps to close.

### Summary by Priority

| Priority | Count | Impact |
|----------|-------|--------|
| **P0 — Critical** | 4 | ✅ All implemented |
| **P1 — High** | 8 | Common document features not yet exposed |
| **P2 — Medium** | 12 | Nice-to-have, possible with docx escape hatch |
| **P3 — Low/Niche** | 8 | Advanced/rarely-used features |

---

## 2. Document Structure Features

### 2.1 Headers & Footers

| Feature | docxjs | docx-kit | Gap | Priority |
|---------|--------|----------|-----|----------|
| Basic headers/footers | ✅ | ✅ | `.section({ header/footer })` | **P0** ✅ |
| Different first page | ✅ | ✅ | `HeaderFooterConfig.first` | P1 ✅ |
| Odd/even page headers | ✅ | ✅ | `HeaderFooterConfig.even` | P2 ✅ |
| Images in headers/footers | ✅ | ❌ | — | P1 |
| Header/footer margins | ✅ | ❌ | — | P2 |

### 2.2 Page & Section

| Feature | docxjs | docx-kit | Gap | Priority |
|---------|--------|----------|-----|----------|
| Page margins | ✅ | ✅ | — | — |
| Page orientation (portrait/landscape) | ✅ | ✅ | — | — |
| Page size presets (A3/A4/Legal/Letter) | ✅ | ✅ | — | — |
| Custom page size | ✅ | ✅ | `{width, height}` in PageConfig | — |
| Page borders | ✅ | ❌ | Not exposed in PageConfig | P2 |
| **Multiple sections** | ✅ | ✅ | `.section(config?)` | **P0** ✅ |
| Multiple columns | ✅ | ❌ | — | P2 |
| Column breaks | ✅ | ❌ | — | P3 |
| Line numbers | ✅ | ❌ | — | P3 |
| Page numbers | ✅ | ❌ | Not exposed | P1 |
| Restart page numbers | ✅ | ❌ | — | P2 |
| Section types (continuous/next page) | ✅ | ❌ | — | P1 |
| Page break before | ✅ | ❌ (only `pageBreak()` node) | — | P2 |

### 2.3 Content / DSL Nodes

| Feature | docxjs | docx-kit | Gap | Priority |
|---------|--------|----------|-----|----------|
| Paragraph (`p`) | ✅ | ✅ | — | — |
| Heading h1–h6 | ✅ | ✅ | — | — |
| Image | ✅ | ✅ (basic + floating) | — | — |
| Table | ✅ | ✅ (columns + data) | — | — |
| Page break | ✅ | ✅ | — | — |
| **Numbering / bullet lists** | ✅ | ✅ | `.bulletList()` / `.numberedList()` | **P0** ✅ |
| **Numbered (ordered) lists** | ✅ | ✅ | `.numberedList({ numberingFormat })` | **P0** ✅ |
| Nested numbering | ✅ | ❌ | — | P1 |
| **Hyperlinks** | ✅ | ❌ | Commonly needed | P1 |
| **Bookmarks** | ✅ | ❌ | — | P2 |
| Table of Contents | ✅ | ❌ | — | P2 |
| Horizontal rule / line | ✅ | ❌ | — | P3 |
| Text box / text frame | ✅ | ❌ | — | P3 |
| Tab stops | ✅ | ❌ | — | P2 |
| Footnotes | ✅ | ❌ | — | P2 |
| Comments | ✅ | ❌ | — | P3 |
| Checkboxes / form fields | ✅ | ❌ | — | P3 |
| Math equations (OMML) | ✅ | ❌ | — | P3 |

### 2.4 Table Features

| Feature | docxjs | docx-kit | Gap | Priority |
|---------|--------|----------|-----|----------|
| Basic table with header row | ✅ | ✅ | — | — |
| Striped rows | ✅ | ✅ (`striped` option) | — | — |
| Borders (`bordered` option) | ✅ | ✅ (on/off) | — | — |
| Custom column renderers | ✅ | ✅ (`TableColumn.render`) | — | — |
| Column width (percentage) | ✅ | ✅ | — | — |
| Column alignment | ✅ | ✅ | — | — |
| **Cell merging** | ✅ | ❌ | Common in reports | P1 |
| Cell shading / background | ✅ | Partial (via `cellStyle.backgroundColor`) | Style → cell mapping | P1 |
| Table style presets ("table look") | ✅ | ❌ | — | P2 |
| Side-by-side tables | ✅ | ❌ | — | P2 |
| Floating tables | ✅ | ❌ | — | P3 |
| Complex border styles per-cell | ✅ | ❌ (only global `bordered`) | — | P2 |
| Cell vertical alignment | ✅ | ✅ (`verticalAlign` in style) | — | — |

---

## 3. Style System (DocxStyleRule)

### 3.1 Text Style — Present & Working

| Property | docx-kit | Maps To | Status |
|----------|----------|---------|--------|
| `fontFamily` | ✅ | `TextRun.font` | Done |
| `fontSize` | ✅ | `TextRun.size` (half-pts) | Done |
| `fontWeight` | ✅ | `TextRun.bold` (≥600 → true) | Done |
| `fontStyle` | ✅ | `TextRun.italics` | Done |
| `color` | ✅ | `TextRun.color` | Done |
| `underline` | ✅ | `TextRun.underline` | Done |
| `strike` | ✅ | `TextRun.strike` | Done |
| `allCaps` | ✅ | `TextRun.allCaps` | Done |

### 3.2 Text Style — Missing

| Property | docxjs Support | Priority | Recommendation |
|----------|---------------|----------|----------------|
| `highlight` | `TextRun.highlight` | **P1** | Add with named colors (yellow, green, etc.) |
| `superScript` / `subScript` | `TextRun.superScript` / `TextRun.subScript` | P1 | Boolean + optional offset |
| `smallCaps` | `TextRun.smallCaps` | P2 | Boolean |
| `doubleStrike` | `TextRun.doubleStrike` | P3 | Boolean |
| `rightToLeft` | `TextRun.rightToLeft` | P2 | Boolean |
| `characterSpacing` | `TextRun.characterSpacing` | P2 | Number (twips) |
| `emboss` / `imprint` / `outline` / `shadow` | `TextRun` effects | P3 | Booleans |

### 3.3 Paragraph Style — Present & Working

| Property | docx-kit | Maps To | Status |
|----------|----------|---------|--------|
| `textAlign` | ✅ | `Paragraph.alignment` | Done |
| `textIndent` | ✅ | `Paragraph.indent.firstLine` | Done |
| `lineHeight` | ✅ | `Paragraph.spacing.line` | Done |
| `marginTop` | ✅ | `Paragraph.spacing.before` | Done |
| `marginBottom` | ✅ | `Paragraph.spacing.after` | Done |
| `marginLeft` | ✅ | `Paragraph.indent.left` | Done |
| `marginRight` | ✅ | `Paragraph.indent.right` | Done |
| `margin` (CSS shorthand) | ✅ | Parsed → 4-direction margins | Done |

### 3.4 Paragraph Style — Missing

| Property | docxjs Support | Priority | Recommendation |
|----------|---------------|----------|----------------|
| `keepLines` (keep together) | `Paragraph.keepLines` | P2 | Boolean |
| `keepNext` (keep with next) | `Paragraph.keepNext` | P2 | Boolean |
| `widowControl` | `Paragraph.widowControl` | P3 | Boolean (default true) |
| `pageBreakBefore` | `Paragraph.pageBreakBefore` | P2 | Boolean |
| `outlineLevel` | `Paragraph.outlineLevel` | P3 | Number (0–9) |
| `tabStops` | `Paragraph.tabStops` | P2 | Array of tab stop positions |
| `bullet` / `numbering` | `Paragraph.numbering` | **P0** | Numbering reference |
| Paragraph borders | `Paragraph.borders` | P1 | Already in type but compiles to cell border only |

### 3.5 Border System

| Feature | docx-kit | Status |
|---------|----------|--------|
| `border` (shorthand) | ✅ | Compiles to 4 sides |
| `borderTop/Right/Bottom/Left` | ✅ | Override individual sides |
| Border styles (dashed/dotted/double/none/single) | ✅ | Enum mapped |
| Border color | ✅ | Hex → docx border color |
| Border width | ✅ | pt → twips |
| **Paragraph borders** | ❌ | Currently compiles as cell border only — need `compileParagraphBorder` |

---

## 4. Plugin System

### 4.1 Current State

| Plugin | Status | Description |
|--------|--------|-------------|
| QRCode | ✅ | PNG QR code generation |
| ECharts | ✅ | Browser-only chart rendering |

### 4.2 Recommended Plugins

| Plugin | Priority | Rationale |
|--------|----------|-----------|
| `table` (built-in) | **P0** | Expose docx's native table as a plugin for advanced options |
| `numbering` (built-in) | **P0** | Bullet + ordered list support |
| `header` / `footer` | P1 | Expose header/footer as config + plugin |
| `toc` | P2 | Auto-generate table of contents from headings |
| `watermark` | P2 | Text or image watermark |
| `barcode` | P3 | Companion to QRCode |

---

## 5. Unit System

| Feature | docx-kit | Status |
|---------|----------|--------|
| `pt` unit | ✅ | All converters |
| `px` unit | ✅ | All converters |
| `mm` unit | ✅ | All converters |
| `cm` unit | ✅ | All converters |
| `in` unit | ✅ | All converters |
| `%` unit | ✅ | Column width only |
| Bare numbers (context-dependent) | ✅ | pt for dimensions, px for images |
| `em` / `rem` | ❌ | Not applicable to docx |
| `vh` / `vw` | ❌ | Not applicable |

---

## 6. Export System

| Feature | docxjs | docx-kit | Status |
|---------|--------|----------|--------|
| `toBlob()` | ✅ | ✅ | Browser |
| `toBuffer()` | ✅ | ✅ | Node.js |
| `toBase64()` | ✅ | ✅ | Cross-platform |
| `toUint8Array()` | Not native | ✅ | Cross-platform |
| `save()` (to disk) | Not native | ✅ | Node.js only |
| `toDocument()` | Not native | ✅ | Raw docx Document |
| `toJSON()` | ❌ | ✅ | Unique feature |
| Streaming output | ✅ (Node.js) | ❌ | P3 |

---

## 7. TypeScript Type System

| Feature | docxjs | docx-kit |
|---------|--------|----------|
| Type-safe stylesheet keys | ❌ | ✅ (ClassName, StyleSheet generics) |
| Type-safe plugin registry | ❌ | ✅ (accumulated generic) |
| Type-safe node union | ❌ | ✅ (BlockNode, InlineNode) |
| Identity helpers (`defineStyles`, etc.) | ❌ | ✅ (const inference) |
| Error taxonomy | ❌ | ✅ (DocxKitError + error codes) |
| UnitValue type | ❌ | ✅ (CSS unit union) |

---

## 8. Metadata & Properties

| Feature | docx-kit | Status |
|---------|----------|--------|
| OOXML core properties (title, creator, etc.) | ✅ | 6 fields |
| Custom document properties | ❌ | P3 |

---

## 9. Priority Roadmap

### P0 — Must Have ✅ All Done

| # | Feature | Effort | Approach | Status |
|---|---------|--------|----------|--------|
| 1 | Numbering / bullet lists | Medium | `BulletListNode` + `compileBulletList`; `.bulletList()` builder | ✅ Done |
| 2 | Numbered / ordered lists | Medium | Same infrastructure as bullet lists, with numbering config | ✅ Done |
| 3 | Multiple sections | Medium | `SectionConfig`; `SectionBreakNode`; `.section()` builder | ✅ Done |
| 4 | Headers & footers | Medium | `HeaderFooterConfig`; `compileHeaders/Footers`; per-section | ✅ Done |

### P1 — High Priority (commonly expected)

| # | Feature | Effort | Approach |
|---|---------|--------|----------|
| 5 | Hyperlinks | Small | `TextRun` with `link` property; new `HyperlinkNode` |
| 6 | Text highlighting | Small | Add `highlight` to `DocxStyleRule` + compile |
| 7 | SuperScript / subScript | Small | Add to `DocxStyleRule` + compile |
| 8 | Cell merging (rowspan/colspan) | Medium | Add `rowSpan`/`colSpan` to data objects or column config |
| 9 | Page numbers | Small | Expose via `PageConfig` or header/footer config |
| 10 | Nested numbering | Medium | Part of numbering system |
| 11 | Paragraph borders (compile fix) | Small | Add `compileParagraphBorder` function |
| 12 | Cell shading (compile fix) | Small | Map `backgroundColor` → cell fill correctly |

### P2 — Medium Priority

| # | Feature | Effort | Approach |
|---|---------|--------|----------|
| 13 | Bookmarks | Small | New `BookmarkNode` |
| 14 | Table of Contents | Medium | Collect headings, generate TOC field |
| 15 | Table style presets | Small | Pass through to docx table properties |
| 16 | Odd/even headers | Small | Config option |
| 17 | Keep with next / keep together | Small | Add to `DocxStyleRule` |
| 18 | Page break before (style) | Small | Add to `DocxStyleRule` |
| 19 | Tab stops | Medium | New `TabStop` type |
| 20 | Right-to-left text | Small | Add to `DocxStyleRule` |
| 21 | Page borders | Medium | Add to `PageConfig` |
| 22 | Character spacing | Small | Add to `DocxStyleRule` |
| 23 | Section types | Small | Add to section config |
| 24 | Side-by-side tables | Medium | Layout helper or section config |

### P3 — Low Priority / Niche

| # | Feature |
|---|---------|
| 25 | Footnotes |
| 26 | Comments |
| 27 | Track changes |
| 28 | Math equations |
| 29 | Checkboxes / form fields |
| 30 | Text boxes / frames |
| 31 | Custom fonts embedding |
| 32 | Watermark |

---

## 10. docx-kit Unique Advantages (over raw docxjs)

These are features docx-kit has that docxjs does NOT:

| Feature | Value |
|---------|-------|
| **CSS-like className system** | Write `docx-kit` styles, not raw docxjs options |
| **Plugin architecture** | Extensible via `definePlugin()` + `use()` + `plugin()` |
| **Style cascading** | `base → className(s) → inline` resolution with `resolveStyle()` |
| **Unit system** | CSS units (pt/px/mm/cm/in/%) with auto-conversion to OOXML twips |
| **Margin/padding shorthand** | `margin: '10pt 20pt'` parses into 4-direction values |
| **Type-safe stylesheet** | `defineStyles({ accent: { color: '#f00' } })` → `className: 'accent'` only accepts valid keys |
| **Type-safe plugins** | Plugin registry accumulates types; `doc.plugin('qrcode', ...)` is fully typed |
| **Platform separation** | `docx-kit`, `docx-kit/node`, `docx-kit/browser` entry points |
| **AI-friendly JSON entry** | `renderDocx({ content: [...], styles: {...} })` for declarative generation |
| **Debug JSON** | `doc.toJSON()` exports full builder state |
| **Error taxonomy** | 7 distinguishable error codes with structured `DocxKitError` |
| **Escape hatch** | `style.docx` property for any docxjs option not yet mapped |
| **ECharts + QRCode** plugins | Built-in chart + QR code generation |
