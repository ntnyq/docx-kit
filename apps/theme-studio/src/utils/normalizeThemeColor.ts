/**
 * Convert accepted HEX input to the six-digit color format required by DOCX.
 */
export function normalizeThemeColor(value: string): string {
  const color = value.trim()
  const short = /^#([\da-f])([\da-f])([\da-f])$/i.exec(color)
  return short
    ? `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`
    : color
}
