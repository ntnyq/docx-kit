/**
 * Warm theme — earthy tones for correspondence and editorial documents.
 *
 * @module themes/warm
 */

import type { DocxTheme } from '@docxkit/core'

/**
 * Warm theme — earthy tones for correspondence and editorial documents.
 *
 * Uses Georgia with Song/Hei East Asian fonts and a quiet warm-gray palette
 * suitable for correspondence and editorial documents on white paper.
 *
 * @remarks
 * - `colors` — 12 semantic color tokens (primary, accent, success, etc.)
 * - `fonts` — body/heading/code font stacks
 * - `fontSize` — 5-step type scale (xs/sm/base/lg/xl)
 * - `spacing` — 5-step spacing scale (xs/sm/md/lg/xl) in points
 */
export const warmTheme: DocxTheme & {
  description: string
  id: string
  name: string
} = {
  description: 'Earthy tones for correspondence and editorial documents.',
  id: 'warm',
  name: 'Warm',
  colors: {
    accent: '#8b5e34',
    background: '#ffffff',
    border: '#ddd6ca',
    danger: '#b91c1c',
    info: '#6b6578',
    muted: '#78716c',
    primary: '#44403c',
    secondary: '#57534e',
    success: '#15803d',
    surface: '#f6f3ee',
    text: '#292524',
    warning: '#a34c13',
  },
  fonts: {
    body: 'Georgia',
    bodyEastAsia: 'SimSun',
    code: 'Consolas',
    heading: 'Georgia',
    headingEastAsia: 'SimHei',
  },
  fontSize: {
    base: 12,
    lg: 15,
    sm: 10,
    xl: 24,
    xs: 8,
  },
  spacing: {
    lg: 26,
    md: 16,
    sm: 9,
    xl: 34,
    xs: 4,
  },
}
