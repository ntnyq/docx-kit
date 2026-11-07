import { createPluginStyle, definePlugin } from '@docxkit/core'
import { BuilderElement, ImportedXmlComponent, Paragraph, TextRun } from 'docx'
import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Default part styles shared by this plugin instance.
 */
export interface CrossReferenceConfig {
  styles?: CrossReferenceStyles
}

/**
 * A block reference to an existing Word bookmark.
 */
export interface CrossReferenceOptions {
  /**
   * Existing bookmark name: a letter followed by letters, digits or underscores,
   * with at most 40 characters. The plugin does not create the target.
   */
  target: string
  /**
   * Result displayed until Word updates the field. Defaults to the target name.
   */
  fallback?: string
  /**
   * Make the field a hyperlink to the target. Defaults to true.
   */
  hyperlink?: boolean
  /**
   * Text uses REF, page uses PAGEREF, number uses REF with full numbering context.
   * Defaults to text. Number requires a bookmark in a numbered paragraph.
   */
  kind?: 'number' | 'page' | 'text'
  prefix?: string
  styles?: CrossReferenceStyles
  suffix?: string
}

/**
 * Formatting for the reference and its surrounding text.
 */
export interface CrossReferenceStyles {
  paragraph?: ParagraphStyleRule
  text?: TextStyleRule
}

/**
 * Create bookmark references whose results can be refreshed in Word.
 */
export function crossReferencePlugin(pluginConfig: CrossReferenceConfig = {}) {
  return definePlugin<'crossReference', CrossReferenceOptions, Paragraph>({
    name: 'crossReference',
    render(options, context) {
      if (
        typeof options.target !== 'string'
        || !/^[A-Z]\w{0,39}$/i.test(options.target)
        || options.target.trim() !== options.target
      ) {
        throw new Error(
          'crossReference: target must be a valid Word bookmark name (1–40 characters)',
        )
      }
      const kind = options.kind ?? 'text'
      if (!['number', 'page', 'text'].includes(kind)) {
        throw new Error('crossReference: kind must be text, number or page')
      }
      const partStyle = (part: keyof CrossReferenceStyles) =>
        createPluginStyle(context.style, {
          defaults: [pluginConfig.styles?.[part]],
          overrides: [options.styles?.[part]],
          theme: context.config.theme,
        })
      const textStyle = partStyle('text')
      const instruction = [
        kind === 'page' ? 'PAGEREF' : 'REF',
        options.target,
        ...(kind === 'number' ? ['\\w'] : []),
        ...(options.hyperlink === false ? [] : ['\\h']),
      ].join(' ')
      const instructionNode = new ImportedXmlComponent('w:instrText', {
        'xml:space': 'preserve',
      })
      instructionNode.push(instruction)
      const cachedRun = new TextRun(
        textStyle.text({
          text: options.fallback ?? options.target,
        }),
      )

      // Complex fields expose the styled cached result to browser readers that
      // ignore fldSimple. Word still refreshes the REF/PAGEREF instruction.
      return new Paragraph(
        partStyle('paragraph').paragraph({
          children: [
            new TextRun(textStyle.text({ text: options.prefix ?? '' })),
            fieldMarker('begin'),
            new TextRun({ children: [instructionNode] }),
            fieldMarker('separate'),
            cachedRun,
            fieldMarker('end'),
            new TextRun(textStyle.text({ text: options.suffix ?? '' })),
          ],
        }),
      )
    },
  })
}

/**
 * Word's complex-field delimiters must be in separate runs from cached text.
 */
function fieldMarker(type: 'begin' | 'end' | 'separate'): TextRun {
  return new TextRun({
    children: [
      new BuilderElement({
        attributes: { type: { key: 'w:fldCharType', value: type } },
        name: 'w:fldChar',
      }),
    ],
  })
}
