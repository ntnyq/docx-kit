import { computed } from 'vue'
import type { Ref } from 'vue'

const englishMessages = {
  autoRender: 'Changes render automatically.',
  baseTheme: 'Base theme',
  buildingDescription: 'Resolving tokens and generating the DOCX…',
  buildingPreview: 'Building preview',
  chooseColor: 'Choose {key} color',
  colors: 'Colors',
  colorSystem: 'Color system',
  copied: 'Copied',
  copy: 'Copy',
  copyFailed: 'Could not copy. Select and copy the code manually.',
  description: 'Description',
  designTokens: 'Design tokens',
  documentLayout: 'Document layout',
  documentPreview: 'Document preview',
  download: '⤓ Download',
  downloadTitle: 'Download .docx file',
  downloadWord: 'Download Word',
  editorAriaLabel: 'Document TypeScript editor',
  editorError: 'Editor Error:',
  editorMissing: 'Editor container not found',
  editorTitle: 'Code Editor',
  emptyPreview: 'Click Run to generate a document.',
  emptyPreviewHint: 'Select an example above or write your own code.',
  exampleBasic: 'Basic',
  exampleChecklist: 'Checklist',
  exampleComprehensive: 'Comprehensive',
  exampleMultiSection: 'Multi-Section',
  examplePlugins: 'Plugins',
  exampleRichContent: 'Rich Content',
  exampleTheme: 'Theme',
  export: 'Export',
  fitWidth: 'Fit width',
  fontDescription: 'Font families for document roles.',
  fontRoles: 'Font roles',
  generationFailed: 'Theme generation failed',
  hideInspector: 'Hide inspector',
  layoutAcademic: 'Academic paper',
  layoutClassic: 'Chinese formal',
  layoutModern: 'Business report',
  name: 'Name',
  none: 'None',
  noPreview: 'No preview yet',
  previewError: 'Preview Error',
  previewLimitationsTitle: 'Preview limitations:',
  previewPending: 'Preview pending',
  previewRenderFailed: 'Preview render failed',
  previewTitle: 'Preview',
  previewZoom: 'Preview zoom',
  refreshingPreview: 'Refreshing preview…',
  rendering: 'Rendering',
  renderingPreview: 'Rendering preview…',
  renderPreview: 'Render preview',
  renderThemeHint: 'Render the theme to inspect the generated document.',
  reset: 'Reset',
  resetCodeTitle: 'Reset current example',
  run: '▶ Run',
  runError: 'Run Error:',
  running: 'Running...',
  showInspector: 'Show inspector',
  spacing: 'Spacing',
  spacingDescription: 'Spacing values for document rhythm.',
  spacingScale: 'Spacing scale',
  starterCode: 'Starter code',
  starterDescription: 'Copy this theme into your project.',
  studioTitle: 'Theme Studio',
  styleAcademic: 'Academic',
  styleAcademicDescription: 'Academic papers and theses with double spacing.',
  styleClassic: 'Classic',
  styleClassicDescription: 'Formal Chinese reports and correspondence.',
  styleLabel: 'Style:',
  styleModern: 'Modern',
  styleModernDescription: 'Clean business reports and professional documents.',
  themeDetails: 'Theme details',
  themeMinimal: 'Minimal',
  themeOcean: 'Ocean',
  themeWarm: 'Warm',
  tokenCategories: 'Token categories',
  typeDescription: 'The type scale used by styles.',
  typeScale: 'Type scale',
  typography: 'Type',
  updatedAt: 'Updated {time}',
  colorDescription:
    'Semantic colors used across text, borders, surfaces, and feedback.',
  downloadAfterError:
    'You can still download the .docx file using the button above.',
  layoutAcademicDescription:
    'Double spacing, justified paragraphs, and table captions.',
  layoutClassicDescription:
    'Centered title, formal sections, and first-line indentation.',
  layoutModernDescription:
    'Flush-left text, compact headings, and clear data tables.',
  previewLimitations:
    'Page numbers and the table of contents may be incomplete. Page breaks and page counts may differ from Word. Download the .docx file, open it in Word, and update fields to check the final layout and page numbers.',
}

export type ToolLocale = 'en-US' | 'zh-CN' | 'zh-TW'
type ToolMessageKey = keyof typeof englishMessages
type ToolMessages = Record<ToolMessageKey, string>

const simplifiedMessages: ToolMessages = {
  autoRender: '修改后自动更新预览。',
  baseTheme: '基础主题',
  buildingDescription: '正在解析令牌并生成 DOCX…',
  buildingPreview: '正在生成预览',
  chooseColor: '选择 {key} 颜色',
  colorDescription: '文字、边框、背景与状态提示使用的语义颜色。',
  colors: '颜色',
  colorSystem: '配色系统',
  copied: '已复制',
  copy: '复制',
  copyFailed: '复制失败，请选中代码后手动复制。',
  description: '说明',
  designTokens: '设计令牌',
  documentLayout: '文档版式',
  documentPreview: '文档预览',
  download: '⤓ 下载',
  downloadAfterError: '仍可通过上方按钮下载 .docx 文件。',
  downloadTitle: '下载 .docx 文件',
  downloadWord: '下载 Word',
  editorAriaLabel: '文档 TypeScript 编辑器',
  editorError: '编辑器错误：',
  editorMissing: '未找到编辑器容器',
  editorTitle: '代码编辑器',
  emptyPreview: '点击「运行」生成文档。',
  emptyPreviewHint: '选择上方示例，或编写自己的代码。',
  exampleBasic: '基础示例',
  exampleChecklist: '任务清单',
  exampleComprehensive: '综合示例',
  exampleMultiSection: '多节文档',
  examplePlugins: '插件',
  exampleRichContent: '富内容',
  exampleTheme: '主题',
  export: '导出',
  fitWidth: '适应宽度',
  fontDescription: '文档中不同用途使用的字体。',
  fontRoles: '字体用途',
  generationFailed: '主题文档生成失败',
  hideInspector: '隐藏设置面板',
  layoutAcademic: '学术论文',
  layoutAcademicDescription: '双倍行距、两端对齐段落与表格题注。',
  layoutClassic: '中文公文',
  layoutClassicDescription: '居中标题、正式章节与首行缩进。',
  layoutModern: '商务报告',
  layoutModernDescription: '正文左对齐、紧凑标题与清晰的数据表格。',
  name: '名称',
  none: '无',
  noPreview: '暂无预览',
  previewError: '预览错误',
  previewLimitationsTitle: '预览限制：',
  previewPending: '等待预览',
  previewRenderFailed: '预览渲染失败',
  previewTitle: '预览',
  previewZoom: '预览缩放',
  refreshingPreview: '正在更新预览…',
  rendering: '渲染中',
  renderingPreview: '正在渲染预览…',
  renderPreview: '渲染预览',
  renderThemeHint: '渲染主题以查看生成的文档。',
  reset: '重置',
  resetCodeTitle: '重置当前示例',
  run: '▶ 运行',
  runError: '运行错误：',
  running: '运行中…',
  showInspector: '显示设置面板',
  spacing: '间距',
  spacingDescription: '用于控制文档排版节奏的间距。',
  spacingScale: '间距体系',
  starterCode: '起始代码',
  starterDescription: '将此主题代码复制到项目中。',
  studioTitle: '主题工作室',
  styleAcademic: '学术论文',
  styleAcademicDescription: '采用双倍行距，适用于学术论文。',
  styleClassic: '经典公文',
  styleClassicDescription: '适用于中文报告与公文信函的正式样式。',
  styleLabel: '样式：',
  styleModern: '现代商务',
  styleModernDescription: '适用于商务报告与专业文档的简洁样式。',
  themeDetails: '主题详情',
  themeMinimal: '极简',
  themeOcean: '海洋',
  themeWarm: '暖色',
  tokenCategories: '令牌分类',
  typeDescription: '样式使用的字号层级。',
  typeScale: '字号体系',
  typography: '字体',
  updatedAt: '更新于 {time}',
  previewLimitations:
    '页码和目录可能不完整，分页位置和页数也可能与 Word 不同。请下载 .docx 文件，在 Word 中打开并更新域，检查最终排版和页码。',
}

const traditionalMessages: ToolMessages = {
  autoRender: '修改後自動更新預覽。',
  baseTheme: '基本主題',
  buildingDescription: '正在解析權杖並產生 DOCX…',
  buildingPreview: '正在產生預覽',
  chooseColor: '選擇 {key} 色彩',
  colorDescription: '文字、框線、背景與狀態提示使用的語意色彩。',
  colors: '色彩',
  colorSystem: '配色系統',
  copied: '已複製',
  copy: '複製',
  copyFailed: '複製失敗，請選取程式碼後手動複製。',
  description: '說明',
  designTokens: '設計權杖',
  documentLayout: '文件版面配置',
  documentPreview: '文件預覽',
  download: '⤓ 下載',
  downloadAfterError: '仍可透過上方按鈕下載 .docx 檔案。',
  downloadTitle: '下載 .docx 檔案',
  downloadWord: '下載 Word',
  editorAriaLabel: '文件 TypeScript 編輯器',
  editorError: '編輯器錯誤：',
  editorMissing: '找不到編輯器容器',
  editorTitle: '程式碼編輯器',
  emptyPreview: '按一下「執行」產生文件。',
  emptyPreviewHint: '選擇上方範例，或撰寫自己的程式碼。',
  exampleBasic: '基本範例',
  exampleChecklist: '工作清單',
  exampleComprehensive: '綜合範例',
  exampleMultiSection: '多節文件',
  examplePlugins: '外掛',
  exampleRichContent: '豐富內容',
  exampleTheme: '主題',
  export: '匯出',
  fitWidth: '符合寬度',
  fontDescription: '文件中不同用途使用的字型。',
  fontRoles: '字型用途',
  generationFailed: '主題文件產生失敗',
  hideInspector: '隱藏設定面板',
  layoutAcademic: '學術論文',
  layoutAcademicDescription: '雙倍行距、左右對齊段落與表格標號。',
  layoutClassic: '中文公文',
  layoutClassicDescription: '置中標題、正式章節與首行縮排。',
  layoutModern: '商務報告',
  layoutModernDescription: '正文靠左對齊、精簡標題與清楚的資料表格。',
  name: '名稱',
  none: '無',
  noPreview: '尚無預覽',
  previewError: '預覽錯誤',
  previewLimitationsTitle: '預覽限制：',
  previewPending: '等待預覽',
  previewRenderFailed: '預覽呈現失敗',
  previewTitle: '預覽',
  previewZoom: '預覽縮放',
  refreshingPreview: '正在更新預覽…',
  rendering: '呈現中',
  renderingPreview: '正在呈現預覽…',
  renderPreview: '呈現預覽',
  renderThemeHint: '呈現主題以查看產生的文件。',
  reset: '重設',
  resetCodeTitle: '重設目前範例',
  run: '▶ 執行',
  runError: '執行錯誤：',
  running: '執行中…',
  showInspector: '顯示設定面板',
  spacing: '間距',
  spacingDescription: '用於控制文件排版節奏的間距。',
  spacingScale: '間距尺度',
  starterCode: '起始程式碼',
  starterDescription: '將此主題程式碼複製到專案中。',
  studioTitle: '主題工作室',
  styleAcademic: '學術論文',
  styleAcademicDescription: '採用雙倍行距，適用於學術論文。',
  styleClassic: '經典公文',
  styleClassicDescription: '適用於中文報告與公文信函的正式樣式。',
  styleLabel: '樣式：',
  styleModern: '現代商務',
  styleModernDescription: '適用於商務報告與專業文件的簡潔樣式。',
  themeDetails: '主題詳細資訊',
  themeMinimal: '極簡',
  themeOcean: '海洋',
  themeWarm: '暖色',
  tokenCategories: '權杖分類',
  typeDescription: '樣式使用的字級層級。',
  typeScale: '字級尺度',
  typography: '字型',
  updatedAt: '更新於 {time}',
  previewLimitations:
    '頁碼與目錄可能不完整，分頁位置和頁數也可能與 Word 不同。請下載 .docx 檔案，在 Word 中開啟並更新功能變數，檢查最終排版與頁碼。',
}

export const messagesByLocale: Record<ToolLocale, ToolMessages> = {
  'en-US': englishMessages,
  'zh-CN': simplifiedMessages,
  'zh-TW': traditionalMessages,
}

const exampleKeys: Record<string, ToolMessageKey> = {
  Basic: 'exampleBasic',
  Checklist: 'exampleChecklist',
  Comprehensive: 'exampleComprehensive',
  'Multi-Section': 'exampleMultiSection',
  Plugins: 'examplePlugins',
  'Rich Content': 'exampleRichContent',
  Theme: 'exampleTheme',
}
const styleKeys: Record<string, ToolMessageKey> = {
  academic: 'styleAcademic',
  classic: 'styleClassic',
  modern: 'styleModern',
}
const styleDescriptionKeys: Record<string, ToolMessageKey> = {
  academic: 'styleAcademicDescription',
  classic: 'styleClassicDescription',
  modern: 'styleModernDescription',
}
const layoutKeys: Record<string, ToolMessageKey> = {
  academic: 'layoutAcademic',
  classic: 'layoutClassic',
  modern: 'layoutModern',
}
const layoutDescriptionKeys: Record<string, ToolMessageKey> = {
  academic: 'layoutAcademicDescription',
  classic: 'layoutClassicDescription',
  modern: 'layoutModernDescription',
}
const themeKeys: Record<string, ToolMessageKey> = {
  minimal: 'themeMinimal',
  ocean: 'themeOcean',
  warm: 'themeWarm',
}

/**
 * Derive UI copy from the documentation language without changing editable data.
 */
export function createToolI18n(language: Readonly<Ref<string>>) {
  const locale = computed(() => resolveToolLocale(language.value))
  const messages = computed(() => messagesByLocale[locale.value])

  function formatMessage(key: ToolMessageKey, values: Record<string, string>) {
    return messages.value[key].replace(
      /\{(\w+)\}/g,
      (match, name: string) => values[name] ?? match,
    )
  }

  function lookup(
    keys: Record<string, ToolMessageKey>,
    id: string,
    fallback: string,
  ) {
    const key = keys[id]
    return key ? messages.value[key] : fallback
  }

  return {
    formatMessage,
    locale,
    messages,
    exampleLabel: (label: string) => lookup(exampleKeys, label, label),
    layoutDescription(id: string, fallback: string) {
      return lookup(layoutDescriptionKeys, id, fallback)
    },
    layoutTitle(id: string, fallback: string) {
      return lookup(layoutKeys, id, fallback)
    },
    styleDescription(id: string, fallback: string) {
      return lookup(styleDescriptionKeys, id, fallback)
    },
    styleTitle(id: string, fallback: string) {
      return lookup(styleKeys, id, fallback)
    },
    themeTitle(id: string, fallback: string) {
      return lookup(themeKeys, id, fallback)
    },
  }
}

/**
 * Match the documentation languages, falling back to English for other locales.
 */
export function resolveToolLocale(language: string): ToolLocale {
  if (language === 'zh-CN' || language === 'zh-TW') {
    return language
  }
  return 'en-US'
}
