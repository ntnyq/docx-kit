import { BUILTIN_PLUGIN_GUIDES } from '@docxkit/ai'
import { badgePlugin } from '@docxkit/plugin-badge'
import { barcodePlugin } from '@docxkit/plugin-barcode'
import { calloutPlugin } from '@docxkit/plugin-callout'
import { changelogPlugin } from '@docxkit/plugin-changelog'
import { checklistPlugin } from '@docxkit/plugin-checklist'
import { codeBlockPlugin } from '@docxkit/plugin-code-block'
import { coverPagePlugin } from '@docxkit/plugin-cover-page'
import { dataTablePlugin } from '@docxkit/plugin-data-table'
import { dividerPlugin } from '@docxkit/plugin-divider'
import { figurePlugin } from '@docxkit/plugin-figure'
import { invoicePlugin } from '@docxkit/plugin-invoice'
import { letterheadPlugin } from '@docxkit/plugin-letterhead'
import { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'
import { metricCardPlugin } from '@docxkit/plugin-metric-card'
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
  checklist: checklistPlugin,
  codeBlock: codeBlockPlugin,
  coverPage: coverPagePlugin,
  dataTable: dataTablePlugin,
  divider: dividerPlugin,
  figure: figurePlugin,
  invoice: invoicePlugin,
  letterhead: letterheadPlugin,
  meetingMinutes: meetingMinutesPlugin,
  metricCard: metricCardPlugin,
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
export const BUILTIN_PLUGIN_CATALOG: readonly BuiltinPluginMetadata[] =
  BUILTIN_PLUGIN_GUIDES.map(plugin => ({
    ...plugin,
    available: Boolean(BUILTIN_FACTORIES[plugin.name]),
    description:
      plugin.name === 'echarts'
        ? 'Browser-only ECharts plugin; unavailable in the Node.js MCP server.'
        : `Built-in docx-kit plugin: ${plugin.name}. ${plugin.description}`,
  }))

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
