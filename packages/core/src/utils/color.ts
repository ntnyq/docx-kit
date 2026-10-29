/**
 * Normalize CSS hexadecimal colors for native Word color, shading and border options.
 * Three-digit colors are expanded; other values retain their native validation.
 */
export function normalizeHexColor(color: string): string
export function normalizeHexColor(color?: string): string | undefined
export function normalizeHexColor(color?: string): string | undefined {
  const hex = color?.replace(/^#/, '')
  return hex && /^[\da-f]{3}$/i.test(hex)
    ? [...hex].map(char => char.repeat(2)).join('')
    : hex
}
