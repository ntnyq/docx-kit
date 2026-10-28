import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import {
  createThemeSample,
  DOCUMENT_LAYOUTS,
} from '../../docs/.vitepress/utils/documentSamples'
import {
  badgePlugin,
  BUILTIN_THEMES,
  calloutPlugin,
  changelogPlugin,
  codeBlockPlugin,
  coverPagePlugin,
  createDocx,
  dataTablePlugin,
  invoicePlugin,
  letterheadPlugin,
  meetingMinutesPlugin,
  modernPreset,
  propertyTablePlugin,
  timelinePlugin,
} from '../../packages/docx-kit/src/node'

const outputRoot = resolve(import.meta.dirname, '../../output/document-design')
await mkdir(outputRoot, { recursive: true })

for (const { preset } of DOCUMENT_LAYOUTS) {
  for (const [themeId, theme] of BUILTIN_THEMES) {
    await save(`${preset.id}-${themeId}`, createThemeSample(preset.id, theme))
  }
}

for (const [themeId, theme] of BUILTIN_THEMES) {
  await save(
    `theme-${themeId}`,
    createDocx({ theme })
      .use(calloutPlugin())
      .h1('Document design review')
      .p(
        'A theme alone supplies the Word font, heading hierarchy, and table palette. 中文与 English 使用独立字体。',
      )
      .h2('Review criteria')
      .table({
        striped: true,
        columns: [
          { key: 'item', title: 'Criterion' },
          { key: 'detail', title: 'Review note' },
        ],
        data: [
          {
            detail: 'Legible body text and restrained headings',
            item: 'Typography',
          },
          { detail: 'Consistent table padding and borders', item: 'Structure' },
        ],
      })
      .plugin('callout', {
        content: 'Use a clear type hierarchy and consistent spacing.',
        title: 'Review note',
        type: 'info',
      })
      .plugin('callout', {
        content: 'The document is ready for the next review.',
        title: 'Approved',
        type: 'success',
      })
      .plugin('callout', {
        content: 'Review tables and page breaks in the exported file.',
        title: 'Check before release',
        type: 'warning',
      })
      .plugin('callout', {
        content: 'Resolve missing content before sharing this document.',
        title: 'Action required',
        type: 'danger',
      }),
  )
}

const doc = createDocx(modernPreset.config)
  .use(badgePlugin())
  .use(calloutPlugin())
  .use(codeBlockPlugin())
  .use(coverPagePlugin())
  .use(dataTablePlugin())
  .use(letterheadPlugin())
  .use(propertyTablePlugin())
  .use(timelinePlugin())
  .plugin('coverPage', {
    author: '产品与体验团队',
    date: '2026年10月',
    organization: '文档管理办公室',
    subtitle: '排版与审核工作规范',
    title: '文档交付质量手册',
  })
  .pageBreak()
  .plugin('letterhead', {
    companyName: '文档管理办公室',
    email: 'documents@example.com',
    tagline: '文档交付质量手册',
  })
  .h1('内容组件')
  .plugin('badge', { color: 'success', text: '已完成复核' })
  .h2('数据与属性')
  .plugin('dataTable', {
    format: { 完成率: 'percent' },
    striped: true,
    data: [
      { 完成率: 0.95, 项目: '排版检查' },
      { 完成率: 0.88, 项目: '内容审核' },
    ],
  })
  .p('')
  .plugin('propertyTable', {
    items: [
      { key: '文档编号', value: 'DOC-2026-028' },
      { key: '责任部门', value: '产品与体验团队' },
    ],
  })
  .h2('提示与代码')
  .plugin('callout', {
    content: '发布前检查表格、分页与中英文混排。',
    title: '审核提示',
    type: 'warning',
  })
  .plugin('codeBlock', {
    code: "const report = createDocx(config)\nreport.h1('Quarterly report')\nawait report.toBlob()",
    showLineNumbers: true,
  })
  .h2('交付计划')
  .plugin('timeline', {
    layout: 'left',
    events: [
      {
        date: '10月28日',
        description: '核对主题与模板的视觉表现',
        title: '版式确认',
      },
      { date: '11月04日', description: '检查最终导出文件', title: '交付复核' },
    ],
  })
doc
  .use(changelogPlugin())
  .use(invoicePlugin())
  .use(meetingMinutesPlugin())
  .pageBreak()
  .plugin('meetingMinutes', {
    attendees: ['陈琳', '周宁'],
    date: '2026-10-28',
    title: '交付复核会议纪要',
    agenda: [
      {
        decision: '纳入交付检查',
        discussion: '确认主题与表格的可读性',
        owner: '陈琳',
        topic: '模板复核',
      },
      {
        decision: '保留最终 Word 样本',
        discussion: '检查字体与分页',
        owner: '周宁',
        topic: '跨平台检查',
      },
    ],
  })
  .h2('文档修订记录')
  .plugin('changelog', {
    entries: [
      {
        changes: '统一标题、表格与提示样式',
        date: '2026-10-28',
        type: 'changed',
        version: '1.1',
      },
      {
        changes: '建立文档模板',
        date: '2026-10-01',
        type: 'added',
        version: '1.0',
      },
    ],
  })
  .pageBreak()
  .plugin('invoice', {
    currency: 'CNY',
    date: '2026-10-28',
    from: { email: 'design@example.com', name: '文档设计团队' },
    invoiceNumber: 'DOC-2026-028',
    notes: 'Please retain the editable document for future updates.',
    taxRate: 0.06,
    to: { name: '交付审核团队' },
    items: [
      { description: 'Document layout review', quantity: 2, unitPrice: 1500 },
      { description: 'Template preparation', quantity: 1, unitPrice: 2400 },
    ],
  })
await save('plugin-showcase', doc)

async function save(
  name: string,
  document: Pick<ReturnType<typeof createDocx>, 'toUint8Array'>,
) {
  const path = resolve(outputRoot, `${name}.docx`)
  await writeFile(path, await document.toUint8Array())
  console.log(path)
}
