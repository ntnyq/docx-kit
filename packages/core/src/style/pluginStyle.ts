import { TextRun } from 'docx'
import {
  compileCellStyle,
  compileParagraphStyle,
  compileTextStyle,
} from '../compiler/compileStyle'
import type { DocxStyleRule } from '@docxkit/types'
import type { IParagraphOptions, IRunOptions, ITableCellOptions } from 'docx'

/**
 * Apply invocation styles to native docx options while preserving plugin defaults.
 */
export interface PluginStyleAdapter {
  /**
   * Override cell borders, shading, alignment and explicit margin sides.
   */
  cell(options: ITableCellOptions): ITableCellOptions
  /**
   * Override paragraph formatting and style shorthand text content.
   */
  paragraph(options: IParagraphOptions): IParagraphOptions
  /**
   * Override run typography without changing its text, fields or breaks.
   */
  text(options: IRunOptions): IRunOptions
}

/**
 * Compile the resolved `context.style` supplied to a plugin render hook.
 * Only explicit overrides replace defaults; nested spacing, borders and indents
 * retain the plugin's other settings. Content and opaque docx objects stay intact.
 */
export function createPluginStyle(
  style: DocxStyleRule = {},
): PluginStyleAdapter {
  const compiledParagraph = compileParagraphStyle(style)
  const compiledCell = compileCellStyle(style)
  const paragraph = {
    ...compiledParagraph,
    alignment:
      style.textAlign === undefined ? undefined : compiledParagraph.alignment,
  }
  const cell = {
    ...compiledCell,
    margins: {
      bottom:
        style.margin !== undefined || style.marginBottom !== undefined
          ? compiledCell.margins.bottom
          : undefined,
      left:
        style.margin !== undefined || style.marginLeft !== undefined
          ? compiledCell.margins.left
          : undefined,
      right:
        style.margin !== undefined || style.marginRight !== undefined
          ? compiledCell.margins.right
          : undefined,
      top:
        style.margin !== undefined || style.marginTop !== undefined
          ? compiledCell.margins.top
          : undefined,
    },
    verticalAlign:
      style.verticalAlign === undefined
        ? undefined
        : compiledCell.verticalAlign,
  }
  const text = compileTextStyle(style)
  // Older text compilers emit implicit false values; an invocation without
  // typography overrides must still preserve the plugin's native defaults.
  if (style.bold === undefined && style.fontWeight === undefined) {
    delete text.bold
  }
  if (style.italic === undefined && style.fontStyle === undefined) {
    delete text.italics
  }
  return {
    cell: options => mergeStyleOptions(options, cell),
    text: options => mergeStyleOptions(options, text),
    paragraph(options) {
      if (options.text === undefined) {
        return mergeStyleOptions(options, paragraph)
      }
      const { text: content, ...rest } = options
      return mergeStyleOptions(
        {
          ...rest,
          children: [
            new TextRun(mergeStyleOptions({ text: content }, text)),
            ...(options.children ?? []),
          ],
        },
        paragraph,
      )
    },
  }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === 'object'
    && value !== null
    && Object.getPrototypeOf(value) === Object.prototype
  )
}

function mergeStyleOptions<TOptions extends object>(
  defaults: TOptions,
  overrides: Record<string, unknown>,
): TOptions {
  const base = defaults as Record<string, unknown>
  return Object.assign(
    {},
    defaults,
    Object.fromEntries(
      Object.entries(overrides)
        .filter(([, value]) => value !== undefined)
        .map(([key, value]) => [
          key,
          isPlainObject(value)
            ? mergeStyleOptions(
                isPlainObject(base[key]) ? base[key] : {},
                value,
              )
            : value,
        ]),
    ),
  )
}
