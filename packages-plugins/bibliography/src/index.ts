import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { Bookmark, InternalHyperlink, Paragraph, TextRun } from 'docx'
import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'
import type { ParagraphChild } from 'docx'

/**
 * Source catalogs may be provided per invocation for JSON/MCP callers.
 */
export interface BibliographyBaseOptions {
  /**
   * Defaults to bib_. Combined with every ID, must be a valid Word bookmark
   * name (letter first, then letters/digits/underscores, at most 40 characters).
   */
  bookmarkPrefix?: string
  /**
   * Replaces the factory catalog, including when empty. Input order sets numbers.
   */
  entries?: BibliographyEntry[]
  styles?: BibliographyStyles
}

/**
 * A citation paragraph. Repeated IDs are deduplicated in first-appearance order.
 */
export interface BibliographyCitationOptions extends BibliographyBaseOptions {
  ids: string[]
  mode: 'citation'
  /**
   * Link numbers to bibliography bookmarks. Defaults to true.
   */
  hyperlink?: boolean
  prefix?: string
  suffix?: string
}

/**
 * Shared source catalog and default styles. No citation state is accumulated.
 */
export interface BibliographyConfig {
  entries?: BibliographyEntry[]
  styles?: BibliographyStyles
}

/**
 * A source in the catalog. Its position determines its stable citation number.
 */
export interface BibliographyEntry {
  /**
   * Unique source ID, also used after bookmarkPrefix in its Word bookmark.
   */
  id: string
  title: string
  authors?: string[]
  doi?: string
  publisher?: string
  url?: string
  year?: string
}

/**
 * Numbered source paragraphs, with one bookmark per source.
 */
export interface BibliographyListOptions extends BibliographyBaseOptions {
  mode?: 'list'
  /**
   * Defaults to References. An empty string omits the heading.
   */
  title?: string
}

export type BibliographyOptions =
  BibliographyCitationOptions | BibliographyListOptions

/**
 * Independently configurable bibliography and citation parts.
 */
export interface BibliographyStyles {
  citation?: TextStyleRule
  citationParagraph?: ParagraphStyleRule
  entry?: TextStyleRule
  entryParagraph?: ParagraphStyleRule
  title?: TextStyleRule
  titleParagraph?: ParagraphStyleRule
}

/**
 * Create numeric references and citations using catalog-order numbering.
 */
export function bibliographyPlugin(pluginConfig: BibliographyConfig = {}) {
  return definePlugin<'bibliography', BibliographyOptions, Paragraph[]>({
    name: 'bibliography',
    render(options, context) {
      const entries = options.entries ?? pluginConfig.entries ?? []
      const bookmarkPrefix = options.bookmarkPrefix ?? 'bib_'
      const numbers = new Map<string, number>()
      for (const [index, entry] of entries.entries()) {
        const bookmarkName = `${bookmarkPrefix}${entry.id}`
        if (
          !entry.id
          || !/^[A-Z]\w{0,39}$/i.test(bookmarkName)
          || bookmarkName.trim() !== bookmarkName
        ) {
          throw new Error(
            'bibliography: bookmarkPrefix + id must be a valid Word bookmark name (1–40 characters)',
          )
        }
        if (numbers.has(entry.id)) {
          throw new Error(`bibliography: duplicate source id "${entry.id}"`)
        }
        if (!entry.title.trim()) {
          throw new Error('bibliography: source title must not be empty')
        }
        numbers.set(entry.id, index + 1)
      }
      const partStyle = (part: keyof BibliographyStyles) =>
        createPluginStyle(context.style, {
          defaults: [pluginConfig.styles?.[part]],
          overrides: [options.styles?.[part]],
          theme: context.config.theme,
        })

      if (options.mode === 'citation') {
        if (
          !Array.isArray(options.ids)
          || !options.ids.every(id => typeof id === 'string' && id.length > 0)
        ) {
          throw new Error(
            'bibliography: citation ids must be an array of nonempty strings',
          )
        }
        const ids = [...new Set(options.ids)]
        if (!ids.length) {
          return []
        }
        const style = partStyle('citation')
        const run = (text: string) => new TextRun(style.text({ text }))
        const children: ParagraphChild[] = [run(`${options.prefix ?? ''}[`)]
        ids.forEach((id, index) => {
          const number = numbers.get(id)
          if (number === undefined) {
            throw new Error(`bibliography: unknown source id "${id}"`)
          }
          if (index) {
            children.push(run(', '))
          }
          const numberRun = run(String(number))
          children.push(
            options.hyperlink === false
              ? numberRun
              : new InternalHyperlink({
                  anchor: `${bookmarkPrefix}${id}`,
                  children: [numberRun],
                }),
          )
        })
        children.push(run(`]${options.suffix ?? ''}`))
        return [
          new Paragraph(partStyle('citationParagraph').paragraph({ children })),
        ]
      }
      if (options.mode !== undefined && options.mode !== 'list') {
        throw new Error('bibliography: mode must be list or citation')
      }
      if (!entries.length) {
        return []
      }
      const paragraphs: Paragraph[] = []
      const title = options.title ?? 'References'
      if (title) {
        paragraphs.push(
          new Paragraph(
            partStyle('titleParagraph').paragraph({
              keepNext: true,
              spacing: { after: 160 },
              children: [
                new TextRun(
                  partStyle('title').text({
                    bold: true,
                    text: title,
                    color:
                      normalizeHexColor(context.config.theme?.colors?.primary)
                      ?? '252525',
                  }),
                ),
              ],
            }),
          ),
        )
      }
      for (const [index, entry] of entries.entries()) {
        const details = [
          entry.authors?.join(', '),
          entry.title,
          entry.year,
          entry.publisher,
          entry.doi ? `DOI: ${entry.doi}` : undefined,
          entry.url,
        ]
          .filter(Boolean)
          .join('. ')
        paragraphs.push(
          new Paragraph(
            partStyle('entryParagraph').paragraph({
              indent: { hanging: 420, left: 420 },
              spacing: { after: 120 },
              children: [
                new Bookmark({
                  id: `${bookmarkPrefix}${entry.id}`,
                  children: [
                    new TextRun(
                      partStyle('entry').text({
                        text: `[${index + 1}] ${details}`,
                      }),
                    ),
                  ],
                }),
              ],
            }),
          ),
        )
      }
      return paragraphs
    },
  })
}
