import { createPluginStyle, definePlugin } from '@docxkit/core'
import { CheckBox, Paragraph, TextRun } from 'docx'

/**
 * A task with an optional owner, note, and group heading.
 */
export interface ChecklistItem {
  /**
   * Task label, also used as the checkbox's accessible alias.
   */
  text: string
  /**
   * Initial Word checkbox state. Defaults to false.
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
 * Options for a grouped checklist with native Word checkbox controls.
 */
export interface ChecklistOptions {
  items: ChecklistItem[]
  title?: string
}

/**
 * Create a checklist plugin for acceptance criteria and inspection records.
 */
export function checklistPlugin() {
  return definePlugin<'checklist', ChecklistOptions, Paragraph[]>({
    name: 'checklist',
    render(options, context) {
      const styles = createPluginStyle(context.style)
      const muted =
        context.config.theme?.colors?.muted?.replace(/^#/, '') ?? '666666'
      const paragraphs: Paragraph[] = []

      if (options.title) {
        paragraphs.push(
          new Paragraph(
            styles.paragraph({
              keepNext: true,
              spacing: { after: 160 },
              children: [
                new TextRun(styles.text({ bold: true, text: options.title })),
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
                keepNext: true,
                spacing: { after: 100, before: 160 },
                children: [
                  new TextRun(styles.text({ bold: true, text: item.group })),
                ],
              }),
            ),
          )
        }
        previousGroup = item.group

        paragraphs.push(
          new Paragraph(
            styles.paragraph({
              keepNext: Boolean(item.note),
              spacing: { after: 100 },
              children: [
                new CheckBox({
                  alias: item.text,
                  checked: item.checked ?? false,
                }),
                new TextRun(styles.text({ text: ` ${item.text}` })),
                ...(item.owner
                  ? [
                      new TextRun(
                        styles.text({ color: muted, text: ` — ${item.owner}` }),
                      ),
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
                indent: { left: 360 },
                spacing: { after: 120 },
                children: [
                  new TextRun(
                    styles.text({ color: muted, size: 18, text: item.note }),
                  ),
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
