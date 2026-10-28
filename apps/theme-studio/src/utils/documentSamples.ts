import {
  academicPreset,
  calloutPlugin,
  classicPreset,
  createDocx,
  dataTablePlugin,
  modernPreset,
  propertyTablePlugin,
} from 'docx-kit'
import type { DocxKitConfig, DocxTheme } from 'docx-kit'

/**
 * Common Word layouts, independent of the editable color palette.
 */
export const DOCUMENT_LAYOUTS = [
  {
    description: 'Flush-left text, compact headings, and clear data tables.',
    preset: modernPreset,
    title: 'Business report',
  },
  {
    description: 'Centered title, formal sections, and first-line indentation.',
    preset: classicPreset,
    title: 'Chinese formal',
  },
  {
    description: 'Double spacing, justified paragraphs, and table captions.',
    preset: academicPreset,
    title: 'Academic paper',
  },
] as const

/**
 * Build the same native DOCX for the studio preview and its download.
 */
export function createThemeSample(layoutId: string, theme: DocxTheme) {
  const doc = createDocx(createThemeSampleConfig(layoutId, theme))
    .use(calloutPlugin())
    .use(dataTablePlugin())
    .use(propertyTablePlugin())

  if (layoutId === 'academic') {
    return doc
      .h1('Document Readability in Review')
      .p('Lin Chen · Document Systems Research · October 2026', {
        className: 'subtitle',
      })
      .h2('Abstract')
      .p(
        'This study examines how typography supports document review. A consistent type scale, predictable spacing, and legible tables help readers locate evidence and compare findings without interrupting the reading flow.',
      )
      .h2('Method and observations')
      .p(
        'Three layouts were reviewed using the same content structure. The comparison considered headings, paragraph spacing, and table alignment. The observations below illustrate review criteria rather than measured experimental results.',
      )
      .plugin('dataTable', {
        striped: true,
        data: [
          {
            Criterion: 'Hierarchy',
            Observation: 'Section headings remain easy to locate',
          },
          {
            Criterion: 'Tables',
            Observation: 'Labels and values share a stable baseline',
          },
          {
            Criterion: 'Pagination',
            Observation: 'Headings stay with the following paragraph',
          },
        ],
      })
      .p('Table 1  Criteria used in the document review', {
        className: 'caption',
      })
      .h2('Discussion')
      .p(
        'Layout should help readers follow the argument. Restrained color and clear spacing preserve that structure when the document is printed or edited in Word. 中文摘要与英文正文采用独立字体设置，保持混排时的阅读节奏。',
      )
  }

  if (layoutId === 'classic') {
    return doc
      .h1('关于推进文档规范化工作的通知')
      .p('文档管理办公室　2026年10月', { className: 'subtitle' })
      .p(
        '各部门：为提高业务材料的阅读与审核效率，现就文档编制、资料归档及交付检查等事项作出安排。各部门应结合实际工作落实要求，确保文档结构清晰、内容完整、格式统一。',
      )
      .h2('一、工作要求')
      .p(
        '正文使用统一字号与首行缩进，标题按层级设置。表格应完整呈现字段名称与业务数据，数值按列对齐，避免因排版差异影响信息理解。',
      )
      .plugin('propertyTable', {
        items: [
          { key: '适用范围', value: '业务报告、工作通知与项目交付材料' },
          { key: '牵头部门', value: '文档管理办公室' },
          { key: '完成时间', value: '2026年11月30日前' },
        ],
      })
      .h2('二、实施安排')
      .p(
        '各部门指定专人完成现有材料梳理，优先统一高频文档的标题、表格与落款格式。完成初步整理后，由牵头部门组织复核并汇总反馈。',
      )
      .plugin('callout', {
        content: '请同时提交可编辑文档及归档清单，便于后续审核与版本管理。',
        title: '报送要求',
        type: 'info',
      })
      .p('文档管理办公室', {
        style: { marginTop: 18, textAlign: 'right', textIndent: 0 },
      })
      .p('2026年10月28日', { style: { textAlign: 'right', textIndent: 0 } })
  }

  return doc
    .h1('产品体验季度报告')
    .p('2026年第三季度　产品与体验团队', { className: 'subtitle' })
    .p(
      '本季度完成了文档导出流程的梳理与核心页面的体验优化。下一阶段将优先解决跨平台排版一致性，并将审核流程纳入日常交付检查。',
    )
    .h2('本季度进展')
    .p(
      '围绕阅读效率、交付质量与协作体验三个方向，团队完成了以下工作。表格中的数量用于展示报告版式。',
    )
    .plugin('dataTable', {
      striped: true,
      data: [
        { 完成项: 12, 负责人: '陈琳', 项目: '导出模板整理' },
        { 完成项: 8, 负责人: '周宁', 项目: '文档质量检查' },
        { 完成项: 16, 负责人: '李明', 项目: '用户反馈跟进' },
      ],
    })
    .p('表 1　本季度主要交付', { className: 'caption' })
    .h2('重点结论')
    .plugin('callout', {
      content: '统一正文、表格与提示内容的阅读节奏，将版式检查放在导出前完成。',
      title: '交付重点',
      type: 'info',
    })
    .h2('下一步安排')
    .p(
      '将常用报告纳入模板库，完善长表格与中英文混排的检查，并在每次发布前对导出文档进行抽样复核。',
    )
}

/**
 * Combine a page layout with editable theme typography for studio samples.
 */
export function createThemeSampleConfig(
  layoutId: string,
  theme: DocxTheme,
): DocxKitConfig {
  const layout =
    DOCUMENT_LAYOUTS.find(item => item.preset.id === layoutId)
    ?? DOCUMENT_LAYOUTS[0]
  const typography = {
    fontFamily:
      theme.fonts?.body ?? layout.preset.config.defaults?.paragraph?.fontFamily,
    fontFamilyEastAsia:
      theme.fonts?.bodyEastAsia
      ?? layout.preset.config.defaults?.paragraph?.fontFamilyEastAsia,
    fontSize:
      theme.fontSize?.base
      ?? layout.preset.config.defaults?.paragraph?.fontSize,
  }
  const paragraphIndent =
    layout.preset.config.defaults?.paragraph?.textIndent
    && typeof typography.fontSize === 'number'
      ? typography.fontSize * 2
      : layout.preset.config.defaults?.paragraph?.textIndent
  const layoutStyles = Object.fromEntries(
    Object.entries(layout.preset.config.styles ?? {}).map(([key, style]) => [
      key,
      {
        ...style,
        ...(key === 'p'
          ? { ...typography, textIndent: paragraphIndent }
          : {
              fontFamily: theme.fonts?.heading ?? style.fontFamily,
              marginBottom: theme.spacing?.sm ?? style.marginBottom,
              fontFamilyEastAsia:
                theme.fonts?.headingEastAsia ?? style.fontFamilyEastAsia,
              fontSize:
                key === 'h1'
                  ? (theme.fontSize?.xl ?? style.fontSize)
                  : key === 'h2'
                    ? (theme.fontSize?.lg ?? style.fontSize)
                    : style.fontSize,
              marginTop:
                key === 'h1' ? 0 : (theme.spacing?.md ?? style.marginTop),
            }),
      },
    ]),
  )
  return {
    ...layout.preset.config,
    theme,
    defaults: {
      ...layout.preset.config.defaults,
      paragraph: { ...layout.preset.config.defaults?.paragraph, ...typography },
      text: { ...layout.preset.config.defaults?.text, ...typography },
    },
    styles: {
      ...layoutStyles,
      caption: {
        color: '$colors.muted',
        fontSize: 9,
        lineHeight: 1.3,
        marginBottom: 12,
        marginTop: 8,
        textIndent: 0,
      },
      subtitle: {
        color: '$colors.muted',
        fontSize: 10,
        marginBottom: 18,
        textAlign: layoutId === 'modern' ? 'left' : 'center',
        textIndent: 0,
      },
    },
  }
}
