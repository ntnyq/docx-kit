import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import {
  BorderStyle,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import type {
  CellStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Default part styles shared by this plugin instance.
 */
export interface GlossaryConfig {
  styles?: GlossaryStyles
}

/**
 * A term or abbreviation, with an optional expanded form and group.
 */
export interface GlossaryEntry {
  definition: string
  term: string
  fullForm?: string
  group?: string
}

/**
 * A three-column terminology table, optionally split into named groups.
 */
export interface GlossaryOptions {
  entries: GlossaryEntry[]
  /**
   * Group headings and separate tables. Defaults to none; unnamed entries come first.
   */
  groupBy?: 'group' | 'none'
  /**
   * Column labels. Defaults to Term, Full form, Definition.
   */
  labels?: { definition?: string; fullForm?: string; term?: string }
  /**
   * Intl.Collator locale for case-insensitive sorting. Defaults to en.
   */
  locale?: string
  /**
   * Sort terms and group headings, or preserve input order. Defaults to term.
   */
  sort?: 'none' | 'term'
  styles?: GlossaryStyles
  /**
   * Defaults to Glossary. An empty string omits the heading.
   */
  title?: string
}

/**
 * Independently configurable glossary parts.
 */
export interface GlossaryStyles {
  bodyCell?: Omit<CellStyleRule, 'height'>
  bodyParagraph?: ParagraphStyleRule
  definition?: TextStyleRule
  fullForm?: TextStyleRule
  group?: TextStyleRule
  groupParagraph?: ParagraphStyleRule
  headerCell?: Omit<CellStyleRule, 'height'>
  headerParagraph?: ParagraphStyleRule
  headerText?: TextStyleRule
  term?: TextStyleRule
  title?: TextStyleRule
  titleParagraph?: ParagraphStyleRule
}

/**
 * Create terminology and abbreviation tables without mutating the input catalog.
 */
export function glossaryPlugin(pluginConfig: GlossaryConfig = {}) {
  return definePlugin<'glossary', GlossaryOptions, (Paragraph | Table)[]>({
    name: 'glossary',
    render(options, context) {
      if (
        options.sort !== undefined
        && !['none', 'term'].includes(options.sort)
      ) {
        throw new Error('glossary: sort must be term or none')
      }
      if (
        options.groupBy !== undefined
        && !['group', 'none'].includes(options.groupBy)
      ) {
        throw new Error('glossary: groupBy must be group or none')
      }
      for (const entry of options.entries) {
        if (!entry.term.trim() || !entry.definition.trim()) {
          throw new Error('glossary: term and definition must not be empty')
        }
      }
      if (!options.entries.length) {
        return []
      }
      const partStyle = (part: keyof GlossaryStyles) =>
        createPluginStyle(context.style, {
          defaults: [pluginConfig.styles?.[part]],
          overrides: [options.styles?.[part]],
          theme: context.config.theme,
        })
      const collator = new Intl.Collator(options.locale ?? 'en', {
        sensitivity: 'base',
      })
      const entries = [...options.entries]
      if (options.sort !== 'none') {
        entries.sort((a, b) => collator.compare(a.term, b.term))
      }
      const groups = new Map<string, GlossaryEntry[]>()
      for (const entry of entries) {
        const group = options.groupBy === 'group' ? (entry.group ?? '') : ''
        const items = groups.get(group) ?? []
        items.push(entry)
        groups.set(group, items)
      }
      const groupedEntries = [...groups.entries()]
      if (options.sort !== 'none') {
        groupedEntries.sort(([a], [b]) => collator.compare(a, b))
      }
      const result: (Paragraph | Table)[] = []
      const primary =
        normalizeHexColor(context.config.theme?.colors?.primary) ?? '252525'
      const heading = (
        text: string,
        part: 'group' | 'title',
        paragraph: 'groupParagraph' | 'titleParagraph',
      ) =>
        new Paragraph(
          partStyle(paragraph).paragraph({
            keepNext: true,
            spacing: { after: 160 },
            children: [
              new TextRun(
                partStyle(part).text({ bold: true, color: primary, text }),
              ),
            ],
          }),
        )
      const title = options.title ?? 'Glossary'
      if (title) {
        result.push(heading(title, 'title', 'titleParagraph'))
      }
      const columns = ['term', 'fullForm', 'definition'] as const
      const defaults = {
        definition: 'Definition',
        fullForm: 'Full form',
        term: 'Term',
      }
      const border = {
        size: 4,
        style: BorderStyle.SINGLE,
        color:
          normalizeHexColor(context.config.theme?.colors?.border) ?? 'D5D9DF',
      }
      for (const [group, items] of groupedEntries) {
        if (group) {
          result.push(heading(group, 'group', 'groupParagraph'))
        }
        const header = new TableRow({
          tableHeader: true,
          children: columns.map(
            column =>
              new TableCell(
                partStyle('headerCell').cell({
                  children: [
                    new Paragraph(
                      partStyle('headerParagraph').paragraph({
                        spacing: { after: 0 },
                        children: [
                          new TextRun(
                            partStyle('headerText').text({
                              bold: true,
                              text:
                                options.labels?.[column] ?? defaults[column],
                            }),
                          ),
                        ],
                      }),
                    ),
                  ],
                  shading: {
                    fill:
                      normalizeHexColor(context.config.theme?.colors?.surface)
                      ?? 'F4F6F8',
                  },
                }),
              ),
          ),
        })
        result.push(
          new Table({
            margins: { bottom: 100, left: 140, right: 140, top: 100 },
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              bottom: border,
              insideHorizontal: border,
              insideVertical: border,
              left: border,
              right: border,
              top: border,
            },
            rows: [
              header,
              ...items.map(
                entry =>
                  new TableRow({
                    children: columns.map(
                      column =>
                        new TableCell(
                          partStyle('bodyCell').cell({
                            children: [
                              new Paragraph(
                                partStyle('bodyParagraph').paragraph({
                                  spacing: { after: 0 },
                                  children: [
                                    new TextRun(
                                      partStyle(column).text({
                                        bold: column === 'term',
                                        text: entry[column] ?? '',
                                      }),
                                    ),
                                  ],
                                }),
                              ),
                            ],
                          }),
                        ),
                    ),
                  }),
              ),
            ],
          }),
          // A paragraph between tables keeps adjacent groups from merging in Word.
          new Paragraph({ spacing: { after: 120, before: 0 } }),
        )
      }
      return result
    },
  })
}
