import { createPluginStyle, definePlugin } from '@docxkit/core'
import { CheckBox, Paragraph, TextRun } from 'docx'
import { createGithubCheckbox } from './checkbox'
import type { DocxStyleRule, PluginRenderContext } from '@docxkit/core'

const CHECKBOX_FONT = 'Segoe UI Symbol'

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
  title?: string
}

/**
 * Create a checklist plugin for acceptance criteria and inspection records.
 */
export function checklistPlugin() {
  return definePlugin<'checklist', ChecklistOptions, Paragraph[]>({
    name: 'checklist',
    async render(options, context) {
      const styles = createPluginStyle(context.style)
      const muted =
        context.config.theme?.colors?.muted?.replace(/^#/, '') ?? '666666'
      const fontSize = resolveTaskFontSize(context)
      const groupSize = Math.round(fontSize * 2 * 1.2)
      const titleSize = Math.round(fontSize * 2 * 1.4)
      const groupColor = (
        context.config.theme?.colors?.accent
        ?? context.config.theme?.colors?.primary
        ?? '#252525'
      ).replace(/^#/, '')
      const titleColor = (
        context.config.theme?.colors?.primary ?? '#252525'
      ).replace(/^#/, '')
      const noteSize = Math.max(1, Math.round(fontSize * 2 * 0.8))
      const ownerSize = Math.max(1, Math.round(fontSize * 2 * 0.85))
      const taskIndent = Math.round(fontSize * 25)
      const paragraphs: Paragraph[] = []

      if (options.title) {
        paragraphs.push(
          new Paragraph(
            styles.paragraph({
              keepNext: true,
              spacing: { after: 160 },
              children: [
                new TextRun({
                  ...styles.text({
                    bold: true,
                    color: titleColor,
                    text: options.title,
                  }),
                  size: titleSize,
                  sizeComplexScript: titleSize,
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
              styles.paragraph({
                keepLines: true,
                keepNext: true,
                children: [
                  new TextRun({
                    ...styles.text({
                      bold: true,
                      color: groupColor,
                      text: item.group,
                    }),
                    size: groupSize,
                    sizeComplexScript: groupSize,
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
            styles.paragraph({
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
                new TextRun(styles.text({ text: ` ${item.text}` })),
                ...(item.owner
                  ? [
                      new TextRun({
                        ...styles.text({ text: ` — ${item.owner}` }),
                        color: muted,
                        size: ownerSize,
                        sizeComplexScript: ownerSize,
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
              styles.paragraph({
                indent: { left: taskIndent },
                keepLines: true,
                spacing: { after: 160 },
                children: [
                  new TextRun({
                    ...styles.text({ color: muted, text: item.note }),
                    size: noteSize,
                    sizeComplexScript: noteSize,
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
            styles.paragraph({
              children: [
                new TextRun(styles.text({ color: muted, text: '(no items)' })),
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
 * Match Word's document defaults, then apply the invocation's run overrides.
 * The style adapter converts CSS units and preserves the docx size escape hatch.
 */
function resolveTaskFontSize(context: PluginRenderContext): number {
  const { config } = context
  const resolveSize = (style: DocxStyleRule) => {
    const fontSize = style.fontSize
    return createPluginStyle({
      ...style,
      fontSize:
        typeof fontSize === 'string' && fontSize.startsWith('$fontSize.')
          ? config.theme?.fontSize?.[fontSize.slice('$fontSize.'.length)]
          : fontSize,
    }).text({}).size
  }
  const inheritedSize = resolveSize({
    fontSize: config.theme?.fontSize?.base ?? 11,
    ...config.defaults?.text,
    ...config.defaults?.paragraph,
  })
  const size = resolveSize(context.style ?? {}) ?? inheritedSize
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
