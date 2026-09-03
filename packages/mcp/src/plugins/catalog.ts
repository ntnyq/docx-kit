import { badgePlugin } from '@docxkit/plugin-badge'
import { barcodePlugin } from '@docxkit/plugin-barcode'
import { calloutPlugin } from '@docxkit/plugin-callout'
import { changelogPlugin } from '@docxkit/plugin-changelog'
import { codeBlockPlugin } from '@docxkit/plugin-code-block'
import { coverPagePlugin } from '@docxkit/plugin-cover-page'
import { dataTablePlugin } from '@docxkit/plugin-data-table'
import { dividerPlugin } from '@docxkit/plugin-divider'
import { invoicePlugin } from '@docxkit/plugin-invoice'
import { letterheadPlugin } from '@docxkit/plugin-letterhead'
import { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'
import { pageNumberPlugin } from '@docxkit/plugin-page-number'
import { propertyTablePlugin } from '@docxkit/plugin-property-table'
import { qrcodePlugin } from '@docxkit/plugin-qrcode'
import { signatureBlockPlugin } from '@docxkit/plugin-signature-block'
import { timelinePlugin } from '@docxkit/plugin-timeline'
import { tocPlugin } from '@docxkit/plugin-toc'
import { watermarkPlugin } from '@docxkit/plugin-watermark'
import type { DocxPlugin, PluginSource } from '@docxkit/core'

// Options are supplied by schema nodes, not invoked through this factory map.
const BUILTIN_FACTORIES: Partial<
  Record<string, () => DocxPlugin<string, never>>
> = {
  badge: badgePlugin,
  barcode: barcodePlugin,
  callout: calloutPlugin,
  changelog: changelogPlugin,
  codeBlock: codeBlockPlugin,
  coverPage: coverPagePlugin,
  dataTable: dataTablePlugin,
  divider: dividerPlugin,
  invoice: invoicePlugin,
  letterhead: letterheadPlugin,
  meetingMinutes: meetingMinutesPlugin,
  pageNumber: pageNumberPlugin,
  propertyTable: propertyTablePlugin,
  qrcode: qrcodePlugin,
  signatureBlock: signatureBlockPlugin,
  timeline: timelinePlugin,
  toc: tocPlugin,
  watermark: watermarkPlugin,
}

export interface BuiltinPluginMetadata {
  available: boolean
  description: string
  name: string
  usageExample: string
}

/**
 * Canonical MCP-facing catalog for all built-in docx-kit plugins.
 */
export const BUILTIN_PLUGIN_CATALOG = [
  entry(
    'badge',
    '{ type: "plugin", name: "badge", options: { text: "New", color: "info" } }',
  ),
  entry(
    'barcode',
    '{ type: "plugin", name: "barcode", options: { text: "DOCX-KIT-2026", format: "code128" } }',
  ),
  entry(
    'callout',
    '{ type: "plugin", name: "callout", options: { type: "info", title: "Note", content: "Important message" } }',
  ),
  entry(
    'changelog',
    '{ type: "plugin", name: "changelog", options: { entries: [{ version: "1.0.0", date: "2026-06-12", type: "added", changes: "Initial release" }] } }',
  ),
  entry(
    'codeBlock',
    '{ type: "plugin", name: "codeBlock", options: { code: "console.log(\'hello\')", language: "javascript" } }',
  ),
  entry(
    'coverPage',
    '{ type: "plugin", name: "coverPage", options: { title: "Annual Report", author: "John Doe", date: "2026-06-12" } }',
  ),
  entry(
    'dataTable',
    '{ type: "plugin", name: "dataTable", options: { data: [{ name: "Alice", score: 95 }], labels: { name: "Name", score: "Score" }, striped: true } }',
  ),
  entry(
    'divider',
    '{ type: "plugin", name: "divider", options: { style: "solid", color: "D9D9D9" } }',
  ),
  entry(
    'echarts',
    '{ type: "plugin", name: "echarts", options: { option: { series: [] }, width: 400, height: 300 } }',
  ),
  entry(
    'invoice',
    '{ type: "plugin", name: "invoice", options: { invoiceNumber: "INV-001", date: "2026-06-12", from: { name: "Seller" }, to: { name: "Client" }, items: [] } }',
  ),
  entry(
    'letterhead',
    '{ type: "plugin", name: "letterhead", options: { companyName: "Acme Inc.", email: "hello@example.com" } }',
  ),
  entry(
    'meetingMinutes',
    '{ type: "plugin", name: "meetingMinutes", options: { title: "Team Meeting", date: "2026-06-12", attendees: ["Alice", "Bob"], agenda: [{ topic: "Roadmap", discussion: "Reviewed milestones" }] } }',
  ),
  entry('pageNumber', '{ type: "plugin", name: "pageNumber", options: {} }'),
  entry(
    'propertyTable',
    '{ type: "plugin", name: "propertyTable", options: { items: [{ key: "Name", value: "Alice" }] } }',
  ),
  entry(
    'qrcode',
    '{ type: "plugin", name: "qrcode", options: { text: "https://example.com", size: 128 } }',
  ),
  entry(
    'signatureBlock',
    '{ type: "plugin", name: "signatureBlock", options: { parties: [{ label: "Manager", name: "Alice" }] } }',
  ),
  entry(
    'timeline',
    '{ type: "plugin", name: "timeline", options: { events: [{ date: "2026-01", title: "Kickoff" }] } }',
  ),
  entry(
    'toc',
    '{ type: "plugin", name: "toc", options: { title: "Contents", maxLevel: 3 } }',
  ),
  entry(
    'watermark',
    '{ type: "plugin", name: "watermark", options: { text: "CONFIDENTIAL", color: "BFBFBF" } }',
  ),
] as const satisfies readonly BuiltinPluginMetadata[]

/**
 * Trusted, Node-compatible plugins available without loading external code.
 */
export function createBuiltinPluginSources(): PluginSource[] {
  return BUILTIN_PLUGIN_CATALOG.flatMap(metadata => {
    const create = BUILTIN_FACTORIES[metadata.name]
    // The schema boundary erases each plugin's specific options type.
    return create
      ? [{ plugin: create() as DocxPlugin, type: 'inline' as const }]
      : []
  })
}

export function findBuiltinPlugin(
  name: string,
): BuiltinPluginMetadata | undefined {
  return BUILTIN_PLUGIN_CATALOG.find(plugin => plugin.name === name)
}

function entry(name: string, usageExample: string): BuiltinPluginMetadata {
  return {
    available: Boolean(BUILTIN_FACTORIES[name]),
    name,
    usageExample,
    description:
      name === 'echarts'
        ? 'Browser-only ECharts plugin; unavailable in the Node.js MCP server.'
        : `Built-in docx-kit plugin: ${name}`,
  }
}
