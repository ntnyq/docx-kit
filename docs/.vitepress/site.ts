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
  { link: '/api/types', text: 'API' },
  { link: '/ecosystem/packages', text: '包总览' },
  { link: '/plugins/', text: '插件' },
  { link: '/examples/basic-report', text: '示例' },
  { link: '/playground', text: 'Playground' },
  { link: '/theme-studio', text: '主题工作室' },
]

export const traditionalNav: LinkItem[] = [
  { link: '/guide/getting-started', text: '指南' },
  { link: '/api/types', text: 'API' },
  { link: '/ecosystem/packages', text: '套件總覽' },
  { link: '/plugins/', text: '外掛' },
  { link: '/examples/basic-report', text: '範例' },
  { link: '/playground', text: 'Playground' },
  { link: '/theme-studio', text: '主題工作室' },
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
  { link: '/guide/markdown', text: 'Markdown' },
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
  { link: '/plugins/checklist', text: 'Checklist' },
  { link: '/plugins/code-block', text: 'Code Block' },
  { link: '/plugins/cover-page', text: 'Cover Page' },
  { link: '/plugins/data-table', text: 'Data Table' },
  { link: '/plugins/divider', text: 'Divider' },
  { link: '/plugins/echarts', text: 'ECharts' },
  { link: '/plugins/figure', text: 'Figure' },
  { link: '/plugins/invoice', text: 'Invoice' },
  { link: '/plugins/letterhead', text: 'Letterhead' },
  { link: '/plugins/meeting-minutes', text: 'Meeting Minutes' },
  { link: '/plugins/metric-card', text: 'Metric Card' },
  { link: '/plugins/page-number', text: 'Page Number' },
  { link: '/plugins/property-table', text: 'Property Table' },
  { link: '/plugins/qrcode', text: 'QR Code' },
  { link: '/plugins/signature-block', text: 'Signature Block' },
  { link: '/plugins/timeline', text: 'Timeline' },
  { link: '/plugins/toc', text: 'Table of Contents' },
  { link: '/plugins/watermark', text: 'Watermark' },
]

export const zhApiLinks: LinkItem[] = [
  { link: '/api/config', text: '配置' },
  { link: '/api/style', text: '样式' },
  { link: '/api/nodes', text: 'DSL 节点' },
  { link: '/api/builder', text: '构建器' },
  { link: '/api/plugins', text: '插件' },
  { link: '/api/types', text: '类型' },
]

export const zhGuideLinks: LinkItem[] = [
  { link: '/guide/getting-started', text: '快速开始' },
  { link: '/guide/builder-api', text: 'Builder API' },
  { link: '/guide/styling', text: 'CSS 风格样式' },
  { link: '/guide/presets', text: '样式预设' },
  { link: '/guide/themes', text: '主题' },
  { link: '/guide/tables', text: '表格' },
  { link: '/guide/images', text: '图片' },
  { link: '/guide/preview', text: '浏览器预览' },
  { link: '/guide/plugins', text: '插件系统' },
  { link: '/guide/json-dsl', text: 'JSON DSL (renderDocx)' },
  { link: '/guide/markdown', text: 'Markdown' },
  { link: '/guide/platforms', text: 'Node.js 与浏览器' },
  { link: '/guide/errors', text: '错误处理' },
  { link: '/guide/plugin-security', text: '插件安全' },
  { link: '/guide/ai-templates', text: 'AI 模板与提示词' },
  { link: '/guide/mcp-server', text: 'MCP 服务端' },
]

export const zhEcosystemLinks: LinkItem[] = [
  { link: '/ecosystem/monorepo', text: 'Monorepo 结构' },
  { link: '/ecosystem/packages', text: '包总览' },
  { link: '/ecosystem/creating-plugins', text: '创建插件' },
  { link: '/playground', text: '在线 Playground' },
  { link: '/theme-studio', text: '主题工作室' },
]

export const zhExampleLinks: LinkItem[] = [
  { link: '/examples/basic-report', text: '基础报告' },
  { link: '/examples/preset-modern', text: '现代预设' },
  { link: '/examples/preset-academic', text: '学术预设' },
  { link: '/examples/ai-generated', text: 'AI 生成文档' },
  { link: '/examples/invoice', text: '发票' },
  { link: '/examples/chart-report', text: '图表报告' },
  { link: '/examples/theme-system', text: '主题系统' },
  { link: '/examples/rich-content', text: '富内容' },
  { link: '/examples/style-inheritance', text: '样式继承' },
  { link: '/examples/preview', text: 'DOCX 预览' },
]

export const zhPluginLinks: LinkItem[] = [
  { link: '/plugins/', text: '插件总览' },
  { link: '/plugins/badge', text: '徽章' },
  { link: '/plugins/barcode', text: '条形码' },
  { link: '/plugins/callout', text: '提示块' },
  { link: '/plugins/changelog', text: '更新日志' },
  { link: '/plugins/checklist', text: '任务清单' },
  { link: '/plugins/code-block', text: '代码块' },
  { link: '/plugins/cover-page', text: '封面' },
  { link: '/plugins/data-table', text: '数据表格' },
  { link: '/plugins/divider', text: '分隔线' },
  { link: '/plugins/echarts', text: 'ECharts' },
  { link: '/plugins/figure', text: '插图' },
  { link: '/plugins/invoice', text: '发票' },
  { link: '/plugins/letterhead', text: '信头' },
  { link: '/plugins/meeting-minutes', text: '会议纪要' },
  { link: '/plugins/metric-card', text: '指标卡' },
  { link: '/plugins/page-number', text: '页码' },
  { link: '/plugins/property-table', text: '属性表' },
  { link: '/plugins/qrcode', text: '二维码' },
  { link: '/plugins/signature-block', text: '签名区' },
  { link: '/plugins/timeline', text: '时间线' },
  { link: '/plugins/toc', text: '目录' },
  { link: '/plugins/watermark', text: '水印' },
]

export const zhTwApiLinks: LinkItem[] = [
  { link: '/api/config', text: '設定' },
  { link: '/api/style', text: '樣式' },
  { link: '/api/nodes', text: 'DSL 節點' },
  { link: '/api/builder', text: '建構器' },
  { link: '/api/plugins', text: '外掛' },
  { link: '/api/types', text: '型別' },
]

export const zhTwGuideLinks: LinkItem[] = [
  { link: '/guide/getting-started', text: '快速開始' },
  { link: '/guide/builder-api', text: 'Builder API' },
  { link: '/guide/styling', text: 'CSS 風格樣式' },
  { link: '/guide/presets', text: '樣式預設' },
  { link: '/guide/themes', text: '主題' },
  { link: '/guide/tables', text: '表格' },
  { link: '/guide/images', text: '圖片' },
  { link: '/guide/preview', text: '瀏覽器預覽' },
  { link: '/guide/plugins', text: '外掛系統' },
  { link: '/guide/json-dsl', text: 'JSON DSL (renderDocx)' },
  { link: '/guide/markdown', text: 'Markdown' },
  { link: '/guide/platforms', text: 'Node.js 與瀏覽器' },
  { link: '/guide/errors', text: '錯誤處理' },
  { link: '/guide/plugin-security', text: '外掛安全' },
  { link: '/guide/ai-templates', text: 'AI 範本與提示詞' },
  { link: '/guide/mcp-server', text: 'MCP 伺服器端' },
]

export const zhTwEcosystemLinks: LinkItem[] = [
  { link: '/ecosystem/monorepo', text: 'Monorepo 結構' },
  { link: '/ecosystem/packages', text: '套件總覽' },
  { link: '/ecosystem/creating-plugins', text: '建立外掛' },
  { link: '/playground', text: '線上 Playground' },
  { link: '/theme-studio', text: '主題工作室' },
]

export const zhTwExampleLinks: LinkItem[] = [
  { link: '/examples/basic-report', text: '基本報告' },
  { link: '/examples/preset-modern', text: '現代預設' },
  { link: '/examples/preset-academic', text: '學術預設' },
  { link: '/examples/ai-generated', text: 'AI 產生文件' },
  { link: '/examples/invoice', text: '發票' },
  { link: '/examples/chart-report', text: '圖表報告' },
  { link: '/examples/theme-system', text: '主題系統' },
  { link: '/examples/rich-content', text: '豐富內容' },
  { link: '/examples/style-inheritance', text: '樣式繼承' },
  { link: '/examples/preview', text: 'DOCX 預覽' },
]

export const zhTwPluginLinks: LinkItem[] = [
  { link: '/plugins/', text: '外掛總覽' },
  { link: '/plugins/badge', text: '徽章' },
  { link: '/plugins/barcode', text: '條碼' },
  { link: '/plugins/callout', text: '提示區塊' },
  { link: '/plugins/changelog', text: '更新記錄' },
  { link: '/plugins/checklist', text: '工作清單' },
  { link: '/plugins/code-block', text: '程式碼區塊' },
  { link: '/plugins/cover-page', text: '封面' },
  { link: '/plugins/data-table', text: '資料表格' },
  { link: '/plugins/divider', text: '分隔線' },
  { link: '/plugins/echarts', text: 'ECharts' },
  { link: '/plugins/figure', text: '插圖' },
  { link: '/plugins/invoice', text: '發票' },
  { link: '/plugins/letterhead', text: '信頭' },
  { link: '/plugins/meeting-minutes', text: '會議記錄' },
  { link: '/plugins/metric-card', text: '指標卡' },
  { link: '/plugins/page-number', text: '頁碼' },
  { link: '/plugins/property-table', text: '屬性表' },
  { link: '/plugins/qrcode', text: 'QR Code' },
  { link: '/plugins/signature-block', text: '簽名區' },
  { link: '/plugins/timeline', text: '時間軸' },
  { link: '/plugins/toc', text: '目錄' },
  { link: '/plugins/watermark', text: '浮水印' },
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
    [`${prefix}/api/`]: section(prefix, 'API 参考', zhApiLinks),
    [`${prefix}/ecosystem/`]: section(prefix, '生态', zhEcosystemLinks),
    [`${prefix}/examples/`]: section(prefix, '示例', zhExampleLinks),
    [`${prefix}/guide/`]: section(prefix, '指南', zhGuideLinks),
    [`${prefix}/plugins/`]: section(prefix, '插件', zhPluginLinks),
  }
}

export function buildTraditionalSidebar(prefix: string) {
  return {
    [`${prefix}/api/`]: section(prefix, 'API 參考', zhTwApiLinks),
    [`${prefix}/ecosystem/`]: section(prefix, '生態', zhTwEcosystemLinks),
    [`${prefix}/examples/`]: section(prefix, '範例', zhTwExampleLinks),
    [`${prefix}/guide/`]: section(prefix, '指南', zhTwGuideLinks),
    [`${prefix}/plugins/`]: section(prefix, '外掛', zhTwPluginLinks),
  }
}

export function navWithPrefix(prefix: string, items: LinkItem[]) {
  return mapLinks(prefix, items)
}
