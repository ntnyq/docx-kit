export interface LinkItem {
  link: string
  text: string
}

export interface SidebarGroup {
  items: LinkItem[]
  text: string
}

const rootPrefix = ''

function mapLinks(prefix: string, items: LinkItem[]) {
  return items.map(item => ({
    ...item,
    link: withPrefix(prefix, item.link),
  }))
}

function section(
  prefix: string,
  text: string,
  items: LinkItem[],
): SidebarGroup[] {
  return [{ items: mapLinks(prefix, items), text }]
}

function withPrefix(prefix: string, link: string) {
  return prefix ? `${prefix}${link}` : link
}

export const englishNav: LinkItem[] = [
  { link: '/guide/getting-started', text: 'Guide' },
  { link: '/api/types', text: 'API' },
  { link: '/ecosystem/packages', text: 'Packages' },
  { link: '/plugins/', text: 'Plugins' },
  { link: '/examples/basic-report', text: 'Examples' },
  { link: '/playground', text: 'Playground' },
  { link: '/theme-studio', text: 'Theme Studio' },
]

export const simplifiedNav: LinkItem[] = [
  { link: '/guide/getting-started', text: '指南' },
  { link: '/ecosystem/monorepo', text: 'Monorepo' },
  { link: '/ecosystem/packages', text: '包总览' },
  { link: '/plugins/', text: '插件' },
  { link: '/playground', text: 'Playground' },
  { link: '/theme-studio', text: 'Theme Studio' },
]

export const traditionalNav: LinkItem[] = [
  { link: '/guide/getting-started', text: '指南' },
  { link: '/ecosystem/monorepo', text: 'Monorepo' },
  { link: '/ecosystem/packages', text: '套件總覽' },
  { link: '/plugins/', text: '外掛' },
  { link: '/playground', text: 'Playground' },
  { link: '/theme-studio', text: 'Theme Studio' },
]

export const apiLinks: LinkItem[] = [
  { link: '/api/config', text: 'Config' },
  { link: '/api/style', text: 'Styles' },
  { link: '/api/nodes', text: 'DSL Nodes' },
  { link: '/api/builder', text: 'Builder' },
  { link: '/api/plugins', text: 'Plugins' },
  { link: '/api/types', text: 'Types' },
]

export const ecosystemLinks: LinkItem[] = [
  { link: '/ecosystem/monorepo', text: 'Monorepo Structure' },
  { link: '/ecosystem/packages', text: 'Package Catalog' },
  { link: '/ecosystem/creating-plugins', text: 'Creating Plugins' },
  { link: '/playground', text: 'Online Playground' },
  { link: '/theme-studio', text: 'Theme Studio' },
]

export const exampleLinks: LinkItem[] = [
  { link: '/examples/basic-report', text: 'Basic Report' },
  { link: '/examples/preset-modern', text: 'Modern Preset' },
  { link: '/examples/preset-academic', text: 'Academic Preset' },
  { link: '/examples/ai-generated', text: 'AI Generated' },
  { link: '/examples/invoice', text: 'Invoice' },
  { link: '/examples/chart-report', text: 'Chart Report' },
  { link: '/examples/theme-system', text: 'Theme System' },
  { link: '/examples/rich-content', text: 'Rich Content' },
  { link: '/examples/style-inheritance', text: 'Style Inheritance' },
  { link: '/examples/preview', text: 'DOCX Preview' },
]

export const guideLinks: LinkItem[] = [
  { link: '/guide/getting-started', text: 'Getting Started' },
  { link: '/guide/builder-api', text: 'Builder API' },
  { link: '/guide/styling', text: 'CSS-like Styling' },
  { link: '/guide/presets', text: 'Style Presets' },
  { link: '/guide/themes', text: 'Themes' },
  { link: '/guide/tables', text: 'Tables' },
  { link: '/guide/images', text: 'Images' },
  { link: '/guide/preview', text: 'Browser Preview' },
  { link: '/guide/plugins', text: 'Plugins' },
  { link: '/guide/json-dsl', text: 'JSON DSL (renderDocx)' },
  { link: '/guide/platforms', text: 'Node.js & Browser' },
  { link: '/guide/errors', text: 'Error Handling' },
  { link: '/guide/plugin-security', text: 'Plugin Security' },
  { link: '/guide/ai-templates', text: 'AI Templates & Prompts' },
  { link: '/guide/mcp-server', text: 'MCP Server' },
]

export const pluginLinks: LinkItem[] = [
  { link: '/plugins/', text: 'Overview' },
  { link: '/plugins/badge', text: 'Badge' },
  { link: '/plugins/barcode', text: 'Barcode' },
  { link: '/plugins/callout', text: 'Callout' },
  { link: '/plugins/changelog', text: 'Changelog' },
  { link: '/plugins/code-block', text: 'Code Block' },
  { link: '/plugins/cover-page', text: 'Cover Page' },
  { link: '/plugins/data-table', text: 'Data Table' },
  { link: '/plugins/divider', text: 'Divider' },
  { link: '/plugins/echarts', text: 'ECharts' },
  { link: '/plugins/invoice', text: 'Invoice' },
  { link: '/plugins/letterhead', text: 'Letterhead' },
  { link: '/plugins/meeting-minutes', text: 'Meeting Minutes' },
  { link: '/plugins/page-number', text: 'Page Number' },
  { link: '/plugins/property-table', text: 'Property Table' },
  { link: '/plugins/qrcode', text: 'QR Code' },
  { link: '/plugins/signature-block', text: 'Signature Block' },
  { link: '/plugins/timeline', text: 'Timeline' },
  { link: '/plugins/toc', text: 'Table of Contents' },
  { link: '/plugins/watermark', text: 'Watermark' },
]

export const zhGuideLinks: LinkItem[] = [
  { link: '/guide/getting-started', text: '快速开始' },
  { link: '/guide/plugins', text: '插件系统' },
]

export const zhEcosystemLinks: LinkItem[] = [
  { link: '/ecosystem/monorepo', text: 'Monorepo 结构' },
  { link: '/ecosystem/packages', text: '包总览' },
]

export const zhPluginLinks: LinkItem[] = [
  { link: '/plugins/', text: '插件总览' },
  { link: '/plugins/badge', text: 'Badge' },
  { link: '/plugins/barcode', text: 'Barcode' },
  { link: '/plugins/callout', text: 'Callout' },
  { link: '/plugins/changelog', text: 'Changelog' },
  { link: '/plugins/code-block', text: 'Code Block' },
  { link: '/plugins/cover-page', text: 'Cover Page' },
  { link: '/plugins/data-table', text: 'Data Table' },
  { link: '/plugins/divider', text: 'Divider' },
  { link: '/plugins/echarts', text: 'ECharts' },
  { link: '/plugins/invoice', text: 'Invoice' },
  { link: '/plugins/letterhead', text: 'Letterhead' },
  { link: '/plugins/meeting-minutes', text: 'Meeting Minutes' },
  { link: '/plugins/page-number', text: 'Page Number' },
  { link: '/plugins/property-table', text: 'Property Table' },
  { link: '/plugins/qrcode', text: 'QR Code' },
  { link: '/plugins/signature-block', text: 'Signature Block' },
  { link: '/plugins/timeline', text: 'Timeline' },
  { link: '/plugins/toc', text: 'TOC' },
  { link: '/plugins/watermark', text: 'Watermark' },
]

export const zhTwEcosystemLinks: LinkItem[] = [
  { link: '/ecosystem/monorepo', text: 'Monorepo 結構' },
  { link: '/ecosystem/packages', text: '套件總覽' },
]

export function buildEnglishSidebar() {
  return {
    '/api/': section(rootPrefix, 'API Reference', apiLinks),
    '/ecosystem/': section(rootPrefix, 'Ecosystem', ecosystemLinks),
    '/examples/': section(rootPrefix, 'Examples', exampleLinks),
    '/guide/': section(rootPrefix, 'Guide', guideLinks),
    '/plugins/': section(rootPrefix, 'Plugins', pluginLinks),
  }
}

export function buildSimplifiedSidebar(prefix: string) {
  return {
    [`${prefix}/ecosystem/`]: section(prefix, '生态', zhEcosystemLinks),
    [`${prefix}/guide/`]: section(prefix, '指南', zhGuideLinks),
    [`${prefix}/plugins/`]: section(prefix, '插件', zhPluginLinks),
  }
}

export function buildTraditionalSidebar(prefix: string) {
  return {
    [`${prefix}/ecosystem/`]: section(prefix, '生態', zhTwEcosystemLinks),
    [`${prefix}/guide/`]: section(prefix, '指南', zhGuideLinks),
    [`${prefix}/plugins/`]: section(prefix, '外掛', zhPluginLinks),
  }
}

export function navWithPrefix(prefix: string, items: LinkItem[]) {
  return mapLinks(prefix, items)
}
