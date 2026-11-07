import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { CheckBox, Paragraph, TextRun } from 'docx'
import { createGithubCheckbox } from './checkbox'
import type {
  DocxStyleRule,
  ParagraphStyleRule,
  PluginRenderContext,
  PluginStyleAdapter,
  TextStyleRule,
} from '@docxkit/core'
import type { IRunOptions } from 'docx'

const CHECKBOX_FONT = 'Segoe UI Symbol'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface ChecklistConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: ChecklistStyles
}

/**
 * A task with an optional owner, note, and group heading.
 */
export interface ChecklistItem {
  /**
   * Task label, also used in graphic alt text or as the native checkbox alias.
   */
  text: string
  /**
   * Initial checkbox state. Defaults to false.
   */
  checked?: boolean
  /**
   * Heading emitted whenever the group changes, preserving item order.
   */
  group?: string
  /**
   * Additional text displayed below the task.
   */
  note?: string
  /**
   * Responsible person, displayed after the task label.
   */
  owner?: string
}

/**
 * Options for a grouped checklist with graphic or native Word checkboxes.
 */
export interface ChecklistOptions {
  items: ChecklistItem[]
  /**
   * Static GitHub-style graphics by default; native enables toggling in Word.
   */
  checkboxStyle?: 'github' | 'native'
  /**
   * Per-part styles for this invocation.
   */
  styles?: ChecklistStyles
  title?: string
}

/**
 * Independently configurable Checklist parts.
 */
export interface ChecklistStyles {
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Group run typography.
   */
  group?: TextStyleRule
  /**
   * Group paragraph alignment, spacing and flow.
   */
  groupParagraph?: ParagraphStyleRule
  /**
   * Item run typography.
   */
  item?: TextStyleRule
  /**
   * Item paragraph alignment, spacing and flow.
   */
  itemParagraph?: ParagraphStyleRule
  /**
   * Note run typography.
   */
  note?: TextStyleRule
  /**
   * Note paragraph alignment, spacing and flow.
   */
  noteParagraph?: ParagraphStyleRule
  /**
   * Owner run typography.
   */
  owner?: TextStyleRule
  /**
   * Title run typography.
   */
  title?: TextStyleRule
  /**
   * Title paragraph alignment, spacing and flow.
   */
  titleParagraph?: ParagraphStyleRule
}

/**
 * Create a checklist plugin for acceptance criteria and inspection records.
 */
export function checklistPlugin(pluginConfig: ChecklistConfig = {}) {
  return definePlugin<'checklist', ChecklistOptions, Paragraph[]>({
    name: 'checklist',
    async render(options, context) {
      const commonFontSize = resolveTaskFontSize(context)
      const scales: Partial<Record<keyof ChecklistStyles, number>> = {
        group: 1.2,
        note: 0.8,
        owner: 0.85,
        title: 1.4,
      }
      // Common body sizes retain the existing metadata proportions. A part's
      // explicit size still overrides the scaled size through the normal cascade.
      const partStyle = (part: keyof ChecklistStyles) => {
        const scale = scales[part]
        const hasCommonSize =
          context.style?.fontSize !== undefined
          || context.style?.docx?.size !== undefined
        const style =
          scale !== undefined && hasCommonSize
            ? {
                ...context.style,
                fontSize: undefined,
                docx: {
                  ...context.style?.docx,
                  size: Math.max(1, Math.round(commonFontSize * 2 * scale)),
                },
              }
            : context.style
        return createPluginStyle(style, {
          defaults: [pluginConfig.styles?.[part]],
          overrides: [options.styles?.[part]],
          theme: context.config.theme,
        })
      }
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const groupStyle = partStyle('group')
      const groupParagraphStyle = partStyle('groupParagraph')
      const itemStyle = partStyle('item')
      const itemParagraphStyle = partStyle('itemParagraph')
      const noteStyle = partStyle('note')
      const noteParagraphStyle = partStyle('noteParagraph')
      const ownerStyle = partStyle('owner')
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
      const muted =
        normalizeHexColor(context.config.theme?.colors?.muted) ?? '666666'
      const fontSize = resolveTaskFontSize(context, itemStyle)
      const groupSize = Math.round(fontSize * 2 * 1.2)
      const titleSize = Math.round(fontSize * 2 * 1.4)
      const groupColor = normalizeHexColor(
        context.config.theme?.colors?.accent
          ?? context.config.theme?.colors?.primary
          ?? '#252525',
      )
      const titleColor = normalizeHexColor(
        context.config.theme?.colors?.primary ?? '#252525',
      )
      const noteSize = Math.max(1, Math.round(fontSize * 2 * 0.8))
      const ownerSize = Math.max(1, Math.round(fontSize * 2 * 0.85))
      const taskIndent = Math.round(fontSize * 25)
      const paragraphs: Paragraph[] = []

      if (options.title) {
        paragraphs.push(
          new Paragraph(
            titleParagraphStyle.paragraph({
              keepNext: true,
              spacing: { after: 160 },
              children: [
                createMetadataRun(titleStyle, {
                  bold: true,
                  color: titleColor,
                  size: titleSize,
                  sizeComplexScript: titleSize,
                  text: options.title,
                }),
              ],
            }),
          ),
        )
      }

      let previousGroup: string | undefined
      for (const item of options.items) {
        if (item.group && item.group !== previousGroup) {
          paragraphs.push(
            new Paragraph(
              groupParagraphStyle.paragraph({
                keepLines: true,
                keepNext: true,
                children: [
                  createMetadataRun(groupStyle, {
                    bold: true,
                    color: groupColor,
                    size: groupSize,
                    sizeComplexScript: groupSize,
                    text: item.group,
                  }),
                ],
                spacing: {
                  after: Math.round(fontSize * 10),
                  before: paragraphs.length ? Math.round(fontSize * 20) : 0,
                },
              }),
            ),
          )
        }
        previousGroup = item.group

        paragraphs.push(
          new Paragraph(
            itemParagraphStyle.paragraph({
              indent: { hanging: taskIndent, left: taskIndent },
              keepLines: true,
              keepNext: Boolean(item.note),
              spacing: { after: item.note ? 60 : 140 },
              children: [
                options.checkboxStyle === 'native'
                  ? new CheckBox({
                      alias: item.text,
                      checked: item.checked ?? false,
                      checkedState: { font: CHECKBOX_FONT, value: '2611' },
                      uncheckedState: { font: CHECKBOX_FONT, value: '2610' },
                    })
                  : await createGithubCheckbox(
                      item.checked ?? false,
                      item.text,
                      fontSize,
                      context.config.theme,
                    ),
                new TextRun(itemStyle.text({ text: ` ${item.text}` })),
                ...(item.owner
                  ? [
                      createMetadataRun(ownerStyle, {
                        color: muted,
                        size: ownerSize,
                        sizeComplexScript: ownerSize,
                        text: ` — ${item.owner}`,
                      }),
                    ]
                  : []),
              ],
            }),
          ),
        )

        if (item.note) {
          paragraphs.push(
            new Paragraph(
              noteParagraphStyle.paragraph({
                indent: { left: taskIndent },
                keepLines: true,
                spacing: { after: 160 },
                children: [
                  createMetadataRun(noteStyle, {
                    color: muted,
                    size: noteSize,
                    sizeComplexScript: noteSize,
                    text: item.note,
                  }),
                ],
              }),
            ),
          )
        }
      }

      if (!options.items.length) {
        paragraphs.push(
          new Paragraph(
            emptyParagraphStyle.paragraph({
              children: [
                new TextRun(
                  emptyTextStyle.text({ color: muted, text: '(no items)' }),
                ),
              ],
            }),
          ),
        )
      }

      return paragraphs
    },
  })
}

/**
 * Keep Latin and complex-script metadata sizes consistent after style overrides.
 */
function createMetadataRun(
  style: PluginStyleAdapter,
  options: IRunOptions,
): TextRun {
  const styled = style.text(options)
  return new TextRun({
    ...styled,
    sizeComplexScript: style.text({}).sizeComplexScript ?? styled.size,
  })
}

/**
 * Match Word's document defaults, then apply the invocation's run overrides.
 * The style adapter converts CSS units and preserves the docx size escape hatch.
 */
function resolveTaskFontSize(
  context: PluginRenderContext,
  taskStyle?: PluginStyleAdapter,
): number {
  const { config } = context
  const resolveSize = (style: DocxStyleRule) =>
    createPluginStyle(style, { theme: config.theme }).text({}).size
  const inheritedSize = resolveSize({
    fontSize: config.theme?.fontSize?.base ?? 11,
    ...config.defaults?.text,
    ...config.defaults?.paragraph,
  })
  const size =
    taskStyle?.text({}).size
    ?? resolveSize(context.style ?? {})
    ?? inheritedSize
  const pointsPerUnit: Record<string, number> = {
    cm: 72 / 2.54,
    in: 72,
    mm: 72 / 25.4,
    pc: 12,
    pi: 12,
    pt: 1,
  }
  const points =
    typeof size === 'string'
      ? Number(size.slice(0, -2)) * pointsPerUnit[size.slice(-2)]
      : typeof size === 'number'
        ? size / 2
        : undefined
  return typeof points === 'number' && Number.isFinite(points) && points > 0
    ? points
    : 11
}
