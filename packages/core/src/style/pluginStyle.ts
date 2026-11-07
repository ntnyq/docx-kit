import { TextRun } from 'docx'
import {
  compileBorderRule,
  compileCellStyle,
  compileColumnWidth,
  compileParagraphStyle,
  compileTextStyle,
} from '../compiler/compileStyle'
import { resolveThemeTokens } from './theme'
import type { BorderRule, DocxStyleRule, DocxTheme } from '@docxkit/types'
import type { IParagraphOptions, IRunOptions, ITableCellOptions } from 'docx'

/**
 * Apply invocation styles to native docx options while preserving plugin defaults.
 */
export interface PluginStyleAdapter {
  /**
   * Override cell borders, shading, width, alignment and explicit margin sides.
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
 * Ordered style layers around the common invocation style.
 */
export interface PluginStyleOptions {
  /**
   * Plugin initialization defaults, from general to specific.
   */
  defaults?: readonly (DocxStyleRule | undefined)[]
  /**
   * Per-part invocation overrides, from general to specific.
   */
  overrides?: readonly (DocxStyleRule | undefined)[]
  /**
   * Resolve tokens only in style data, never in plugin content.
   */
  theme?: DocxTheme
}

/**
 * Apply initialization defaults, common invocation style, then part overrides.
 * Compile layers separately so shorthand, aliases and explicit false/zero values
 * respect precedence. Plain nested options merge; arrays replace as a whole.
 */
export function createPluginStyle(
  style: DocxStyleRule = {},
  options: PluginStyleOptions = {},
): PluginStyleAdapter {
  const layers = [
    ...(options.defaults ?? []),
    style,
    ...(options.overrides ?? []),
  ]
  const { cell, paragraph, text } = layers.reduce((result, layer) => {
    if (!layer) {
      return result
    }
    const next = compilePluginStyle(resolveThemeTokens(layer, options.theme))
    return {
      cell: mergeStyleOptions(result.cell, next.cell),
      paragraph: mergeStyleOptions(result.paragraph, next.paragraph),
      text: mergeTextStyleOptions(result.text, next.text),
    }
  }, compilePluginStyle({}))
  return {
    text: options => mergeTextStyleOptions(options, text),
    cell(options) {
      const merged = mergeStyleOptions(options, cell)
      return cell.borders
        ? {
            ...merged,
            borders: completePluginBorders(merged.borders, options.borders),
          }
        : merged
    },
    paragraph(options) {
      if (options.text === undefined) {
        const merged = mergeStyleOptions(options, paragraph)
        return paragraph.border
          ? {
              ...merged,
              border: completePluginBorders(merged.border, options.border),
            }
          : merged
      }
      const { text: content, ...rest } = options
      const merged = mergeStyleOptions(
        {
          ...rest,
          children: [
            new TextRun(mergeTextStyleOptions({ text: content }, text)),
            ...(options.children ?? []),
          ],
        },
        paragraph,
      )
      return paragraph.border
        ? {
            ...merged,
            border: completePluginBorders(merged.border, options.border),
          }
        : merged
    },
  }
}

/**
 * Border compilation must not introduce implicit width/style defaults into an
 * override layer, otherwise a color-only override would reset the native border.
 */
function compilePluginBorders(style: DocxStyleRule) {
  const sides = ['bottom', 'left', 'right', 'top'] as const
  const rules = {
    bottom: style.borderBottom,
    left: style.borderLeft,
    right: style.borderRight,
    top: style.borderTop,
  }
  if (!style.border && sides.every(side => !rules[side])) {
    return undefined
  }
  const compileSide = (override?: BorderRule) => {
    const rule = mergeStyleOptions(style.border ?? {}, { ...override })
    if (
      rule.color === undefined
      && rule.width === undefined
      && rule.style === undefined
    ) {
      return undefined
    }
    const compiled = compileBorderRule(rule)
    return {
      color: rule.color === undefined ? undefined : compiled.color,
      size: rule.width === undefined ? undefined : compiled.size,
      style: rule.style === undefined ? undefined : compiled.style,
    }
  }
  return {
    bottom: compileSide(rules.bottom),
    left: compileSide(rules.left),
    right: compileSide(rules.right),
    top: compileSide(rules.top),
  }
}

function compilePluginStyle(style: DocxStyleRule) {
  const compiledParagraph = compileParagraphStyle(style)
  const compiledCell = compileCellStyle(style)
  const borders = compilePluginBorders(style)
  const paragraph = {
    ...compiledParagraph,
    border: borders,
    alignment:
      style.textAlign === undefined ? undefined : compiledParagraph.alignment,
  }
  const cell = {
    ...compiledCell,
    borders,
    width: compileColumnWidth(style.width),
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
  return {
    cell,
    paragraph,
    text,
  }
}

/**
 * Supply required border defaults after the full cascade, so a new partial
 * border renders while color-only changes retain an existing width and style.
 */
function completePluginBorders<TBorders extends object>(
  borders: TBorders | undefined,
  nativeBorders: TBorders | undefined,
): TBorders | undefined {
  if (!borders) {
    return borders
  }
  const defaults = compileBorderRule({})
  return Object.assign(
    {},
    borders,
    Object.fromEntries(
      Object.entries(borders).map(([side, rule]) => [
        side,
        isPlainObject(rule)
        && (!nativeBorders || !isPlainObject(Reflect.get(nativeBorders, side)))
          ? mergeStyleOptions(defaults, rule)
          : rule,
      ]),
    ),
  )
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
        .filter(([key, value]) => value !== undefined && key !== '__proto__')
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

/**
 * Expand a native font shorthand before merging script-specific font overrides.
 */
function mergeTextStyleOptions<TOptions extends object>(
  defaults: TOptions,
  overrides: Record<string, unknown>,
): TOptions {
  const font: unknown = Reflect.get(defaults, 'font')
  const normalized =
    typeof font === 'string' && isPlainObject(overrides.font)
      ? {
          ...defaults,
          font: { ascii: font, cs: font, eastAsia: font, hAnsi: font },
        }
      : defaults
  return mergeStyleOptions(normalized, overrides)
}
