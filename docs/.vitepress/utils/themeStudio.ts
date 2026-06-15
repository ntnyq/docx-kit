import type { DocxTheme } from 'docx-kit'

export type BuiltinTheme = DocxTheme & {
  id: string
  name: string
  description?: string
}

export interface EditableTheme {
  colors: Record<string, string>
  description: string
  fonts: Record<string, string>
  fontSize: Record<string, number | string>
  id: string
  name: string
  spacing: Record<string, number | string>
}

export type EditableThemeCategory = 'colors' | 'fonts' | 'fontSize' | 'spacing'

export function cloneThemeForEditing(theme: BuiltinTheme): EditableTheme {
  return {
    colors: { ...(theme.colors ?? {}) },
    description: theme.description ?? '',
    fonts: { ...(theme.fonts ?? {}) },
    fontSize: { ...(theme.fontSize ?? {}) },
    id: theme.id,
    name: theme.name,
    spacing: { ...(theme.spacing ?? {}) },
  }
}

export function serializeThemeToSnippet(theme: EditableTheme): string {
  return [
    "import type { DocxTheme } from 'docx-kit'",
    '',
    'export const customTheme: DocxTheme & {',
    '  id: string',
    '  name: string',
    '  description: string',
    '} = {',
    `  id: '${theme.id}',`,
    `  name: '${theme.name}',`,
    `  description: '${escapeSingleQuotes(theme.description)}',`,
    `  colors: ${formatTokenObject(theme.colors)},`,
    `  fonts: ${formatTokenObject(theme.fonts)},`,
    `  fontSize: ${formatTokenObject(theme.fontSize)},`,
    `  spacing: ${formatTokenObject(theme.spacing)},`,
    '}',
    '',
  ].join('\n')
}

export function toDocxTheme(theme: EditableTheme): DocxTheme {
  return {
    colors: { ...theme.colors },
    fonts: { ...theme.fonts },
    fontSize: normalizeNumericTokenMap(theme.fontSize) as NonNullable<
      DocxTheme['fontSize']
    >,
    spacing: normalizeNumericTokenMap(theme.spacing) as NonNullable<
      DocxTheme['spacing']
    >,
  }
}

export function updateThemeToken(
  theme: EditableTheme,
  category: EditableThemeCategory,
  key: string,
  value: string,
): EditableTheme {
  return {
    ...theme,
    [category]: {
      ...theme[category],
      [key]: value,
    },
  }
}

function escapeSingleQuotes(value: string): string {
  return value.replaceAll("'", "\\'")
}

function formatTokenObject(tokens: Record<string, number | string>): string {
  const entries = Object.entries(tokens)
  if (entries.length === 0) {
    return '{}'
  }

  const body = entries
    .map(([key, value]) => {
      const serializedValue =
        typeof value === 'number'
          ? String(value)
          : `'${escapeSingleQuotes(value)}'`

      return `    ${key}: ${serializedValue}`
    })
    .join(',\n')

  return `{\n${body}\n  }`
}

function normalizeNumericTokenMap(
  tokens: Record<string, number | string>,
): Record<string, number | string> {
  return Object.fromEntries(
    Object.entries(tokens).map(([key, value]) => {
      if (typeof value === 'number') {
        return [key, value]
      }

      const trimmed = value.trim()
      if (/^-?\d+(?:\.\d+)?$/.test(trimmed)) {
        return [key, Number(trimmed)]
      }

      return [key, trimmed]
    }),
  )
}
