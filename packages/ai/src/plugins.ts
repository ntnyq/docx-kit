/**
 * Shared AI and MCP usage examples. Node-compatible examples are rendered in
 * MCP integration tests so help stays aligned with the actual plugin contract.
 */
interface BuiltinPluginGuide {
  description: string
  name: string
  usageExample: string
}

/**
 * Built-in plugin descriptions and executable DSL examples, including ECharts
 * for browser callers. This catalog does not import plugin implementations.
 */
export const BUILTIN_PLUGIN_GUIDES: readonly BuiltinPluginGuide[] = [
  entry(
    'checklist',
    'Grouped checkbox tasks with owners and notes',
    '{ type: "plugin", name: "checklist", options: { items: [{ text: "Review release", checked: true, owner: "Alice", group: "Release" }] } }',
  ),
  entry(
    'figure',
    'Image with caption and source credit',
    '{ type: "plugin", name: "figure", options: { data: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==", caption: "Figure 1. Overview", width: 400 } }',
  ),
  entry(
    'metricCard',
    'Metric summaries with values and trends',
    '{ type: "plugin", name: "metricCard", options: { items: [{ label: "Revenue", value: 120, unit: "USD", trend: "up", change: "+12%" }], columns: 3 } }',
  ),
  entry(
    'badge',
    'Status badges with semantic colors',
    '{ type: "plugin", name: "badge", options: { text: "New", color: "info" } }',
  ),
  entry(
    'barcode',
    'Linear barcode generation',
    '{ type: "plugin", name: "barcode", options: { text: "DOCX-KIT-2026", format: "code128" } }',
  ),
  entry(
    'callout',
    'Highlighted information, warning, success, or danger box',
    '{ type: "plugin", name: "callout", options: { type: "info", title: "Note", content: "Important message" } }',
  ),
  entry(
    'changelog',
    'Versioned change entries',
    '{ type: "plugin", name: "changelog", options: { entries: [{ version: "1.0.0", date: "2026-06-12", type: "added", changes: "Initial release" }] } }',
  ),
  entry(
    'codeBlock',
    'Source code with optional syntax highlighting',
    '{ type: "plugin", name: "codeBlock", options: { code: "console.log(\'hello\')", language: "javascript" } }',
  ),
  entry(
    'coverPage',
    'Cover title, subtitle, author, organization, and date',
    '{ type: "plugin", name: "coverPage", options: { title: "Annual Report", author: "John Doe", date: "2026-06-12" } }',
  ),
  entry(
    'dataTable',
    'Data table with inferred columns and value formatting',
    '{ type: "plugin", name: "dataTable", options: { data: [{ name: "Alice", score: 95 }], labels: { name: "Name", score: "Score" }, striped: true } }',
  ),
  entry(
    'divider',
    'Horizontal divider',
    '{ type: "plugin", name: "divider", options: { style: "solid", color: "D9D9D9" } }',
  ),
  entry(
    'echarts',
    'Browser-only ECharts chart embedding',
    '{ type: "plugin", name: "echarts", options: { option: { series: [] }, width: 400, height: 300 } }',
  ),
  entry(
    'invoice',
    'Invoice with line items, tax, and totals',
    '{ type: "plugin", name: "invoice", options: { invoiceNumber: "INV-001", date: "2026-06-12", from: { name: "Seller" }, to: { name: "Client" }, items: [] } }',
  ),
  entry(
    'letterhead',
    'Company letterhead and contact information',
    '{ type: "plugin", name: "letterhead", options: { companyName: "Acme Inc.", email: "hello@example.com" } }',
  ),
  entry(
    'meetingMinutes',
    'Meeting title, date, attendees, and agenda',
    '{ type: "plugin", name: "meetingMinutes", options: { title: "Team Meeting", date: "2026-06-12", attendees: ["Alice", "Bob"], agenda: [{ topic: "Roadmap", discussion: "Reviewed milestones" }] } }',
  ),
  entry(
    'pageNumber',
    'Current page or page X of Y fields',
    '{ type: "plugin", name: "pageNumber", options: {} }',
  ),
  entry(
    'propertyTable',
    'Key-value property table',
    '{ type: "plugin", name: "propertyTable", options: { items: [{ key: "Name", value: "Alice" }] } }',
  ),
  entry(
    'qrcode',
    'QR code generation',
    '{ type: "plugin", name: "qrcode", options: { text: "https://example.com", size: 128 } }',
  ),
  entry(
    'signatureBlock',
    'Signature grid with party labels, names, and dates',
    '{ type: "plugin", name: "signatureBlock", options: { parties: [{ label: "Manager", name: "Alice" }] } }',
  ),
  entry(
    'timeline',
    'Chronological events',
    '{ type: "plugin", name: "timeline", options: { events: [{ date: "2026-01", title: "Kickoff" }] } }',
  ),
  entry(
    'toc',
    'Table of contents field',
    '{ type: "plugin", name: "toc", options: { title: "Contents", maxLevel: 3 } }',
  ),
  entry(
    'watermark',
    'Watermark text paragraph for a header or footer',
    '{ type: "plugin", name: "watermark", options: { text: "CONFIDENTIAL", color: "BFBFBF" } }',
  ),
] as const satisfies readonly BuiltinPluginGuide[]

function entry(
  name: string,
  description: string,
  usageExample: string,
): BuiltinPluginGuide {
  return {
    description,
    name,
    usageExample,
  }
}
