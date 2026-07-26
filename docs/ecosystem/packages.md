# Package Catalog

This page is the package-level map of the current docx-kit monorepo. It complements the plugin reference by showing how the workspace is split across runtime, tooling, presets, and themes.

## Core Runtime Packages

| Package | Purpose | When to use it |
| --- | --- | --- |
| `docx-kit` | Umbrella package | Default choice for app developers who want one import surface |
| `@docxkit/core` | Core builder, compiler, DSL, styles, utilities | Use when you want the minimal runtime or are building advanced integrations |
| `@docxkit/types` | Shared type definitions | Use in packages or apps that only need types |
| `@docxkit/renderer` | Browser preview rendering | Use when you need in-browser `.docx` preview or Office Online embedding |

## Extensibility And Tooling

| Package | Purpose | Key capability |
| --- | --- | --- |
| `@docxkit/loader` | Dynamic plugin loading | Inline, npm, URL, and local plugin sources |
| `@docxkit/pdk` | Plugin development kit | Plugin rendering tests and reusable test context |
| `@docxkit/registry` | Plugin registry search | Finds plugins published for docx-kit |
| `@docxkit/create-plugin` | Plugin scaffolder | Creates new plugin packages with the expected layout |

## AI And Agent Packages

| Package | Purpose | Key capability |
| --- | --- | --- |
| `@docxkit/ai` | AI-friendly template system | Report, invoice, resume, and letter templates plus prompt building |
| `@docxkit/mcp` | MCP server | Exposes document-generation tools to AI agents and assistants |

## Built-in Plugin Packages

| Package | Focus |
| --- | --- |
| `@docxkit/plugin-badge` | Status labels |
| `@docxkit/plugin-barcode` | Linear barcodes |
| `@docxkit/plugin-callout` | Admonition blocks |
| `@docxkit/plugin-changelog` | Release-note tables |
| `@docxkit/plugin-code-block` | Source code formatting |
| `@docxkit/plugin-cover-page` | Report and document covers |
| `@docxkit/plugin-data-table` | Structured object-array tables |
| `@docxkit/plugin-divider` | Section separators |
| `@docxkit/plugin-echarts` | Chart rendering |
| `@docxkit/plugin-invoice` | Billing documents |
| `@docxkit/plugin-letterhead` | Branded letter headers |
| `@docxkit/plugin-meeting-minutes` | Meeting notes |
| `@docxkit/plugin-page-number` | Page number fields |
| `@docxkit/plugin-property-table` | Key-value tables |
| `@docxkit/plugin-qrcode` | QR codes |
| `@docxkit/plugin-signature-block` | Approval and signature layouts |
| `@docxkit/plugin-timeline` | Milestone timelines |
| `@docxkit/plugin-toc` | Table of contents fields |
| `@docxkit/plugin-watermark` | Watermark text |

The full option-level reference for each plugin lives in the [Plugins section](/plugins/).

## Preset Packages

| Package | Style direction |
| --- | --- |
| `@docxkit/preset-classic` | Formal official-document styling |
| `@docxkit/preset-modern` | Business-friendly default styling |
| `@docxkit/preset-academic` | Thesis and paper formatting |

## Theme Packages

| Package | Theme direction |
| --- | --- |
| `@docxkit/theme-minimal` | Neutral minimal palette |
| `@docxkit/theme-ocean` | Blue-led color system |
| `@docxkit/theme-warm` | Amber and warm-tone color system |

## Recommended Entry Points

- Choose `docx-kit` if you are building an application.
- Choose `@docxkit/core` if you are building an adapter, plugin, or constrained runtime bundle.
- Choose `@docxkit/renderer` when previewing output in the browser.
- Choose `@docxkit/pdk` and `@docxkit/create-plugin` when extending the ecosystem.
- Choose `@docxkit/ai` or `@docxkit/mcp` when the document pipeline is AI-driven.
