/**
 * Ocean theme — deep blue and teal palette for structured business reports.
 *
 * @module themes/ocean
 */

import type { DocxTheme } from '@docxkit/core'

/**
 * Ocean theme — deep blue and teal palette for structured business reports.
 *
 * Uses Cambria headings and Calibri body text and a calming blue/teal color palette
 * suitable for reports, articles, and formal documents.
 *
 * @remarks
 * - `colors` — 12 semantic color tokens (primary, accent, success, etc.)
 * - `fonts` — body/heading/code font stacks
 * - `fontSize` — 5-step type scale (xs/sm/base/lg/xl)
 * - `spacing` — 5-step spacing scale (xs/sm/md/lg/xl) in points
 */
export const oceanTheme: DocxTheme & {
  description: string
  id: string
  name: string
} = {
  description: 'Deep blue and teal palette for structured business reports.',
  id: 'ocean',
  name: 'Ocean',
  colors: {
    accent: '#0d9488',
    background: '#ffffff',
    border: '#cbd9df',
    danger: '#e11d48',
    info: '#24758a',
    muted: '#64748b',
    primary: '#0f172a',
    secondary: '#1e293b',
    success: '#14765c',
    surface: '#f0f5f7',
    text: '#1e293b',
    warning: '#946200',
  },
  fonts: {
    body: 'Calibri',
    bodyEastAsia: 'Microsoft YaHei',
    code: 'Consolas',
    heading: 'Cambria',
    headingEastAsia: 'Microsoft YaHei',
  },
  fontSize: {
    base: 12,
    lg: 16,
    sm: 10,
    xl: 26,
    xs: 8,
  },
  spacing: {
    lg: 28,
    md: 18,
    sm: 10,
    xl: 36,
    xs: 5,
  },
}
