import { normalizeThemeColor } from './normalizeThemeColor'
import type { EditableTheme, EditableThemeCategory } from './themeStudio'

export type TokenError = 'fontRequired' | 'invalidColor' | 'invalidNumber'

/**
 * Validate editable text before passing it to the document compiler.
 */
export function getTokenError(
  category: EditableThemeCategory,
  value: number | string,
): TokenError | undefined {
  const text = String(value).trim()
  if (category === 'colors') {
    return /^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(text)
      ? undefined
      : 'invalidColor'
  }
  if (category === 'fonts') {
    return text ? undefined : 'fontRequired'
  }
  return /^(?:\d+(?:\.\d+)?|\.\d+)(?:pt|px|cm|mm|in|%)?$/.test(text)
    ? undefined
    : 'invalidNumber'
}

export function hasInvalidThemeTokens(theme: EditableTheme): boolean {
  return (['colors', 'fonts', 'fontSize', 'spacing'] as const).some(category =>
    Object.values(theme[category]).some(value =>
      getTokenError(category, value),
    ),
  )
}

export function normalizeColorInput(value: string): string {
  const color = normalizeThemeColor(value)
  return /^#[\da-f]{6}$/i.test(color) ? color : '#000000'
}
