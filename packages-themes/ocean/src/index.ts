/**
 * Ocean theme — deep blue and teal palette with warm coral highlights.
 *
 * @module themes/ocean
 */

import type { DocxTheme } from '@docxkit/core'

/**
 * Ocean theme — deep blue and teal palette with warm coral highlights.
 *
 * Uses Georgia serif fonts and a calming blue/teal color palette
 * suitable for reports, articles, and formal documents.
 *
 * @remarks
 * - `colors` — 12 semantic color tokens (primary, accent, success, etc.)
 * - `fonts` — body/heading/code font stacks
 * - `fontSize` — 5-step type scale (xs/sm/base/lg/xl)
 * - `spacing` — 5-step spacing scale (xs/sm/md/lg/xl) in half-points
 */
export const oceanTheme: DocxTheme & {
  description: string
  id: string
  name: string
} = {
  description: 'Deep blue and teal palette with warm coral highlights.',
  id: 'ocean',
  name: 'Ocean',
  colors: {
    accent: '#0d9488',
    background: '#f0fdfa',
    border: '#99f6e4',
    danger: '#e11d48',
    info: '#38bdf8',
    muted: '#64748b',
    primary: '#0f172a',
    secondary: '#1e293b',
    success: '#059669',
    surface: '#ccfbf1',
    text: '#1e293b',
    warning: '#f59e0b',
  },
  fonts: {
    body: 'Georgia, serif',
    code: 'JetBrains Mono, monospace',
    heading: 'Georgia, serif',
  },
  fontSize: {
    base: 12,
    lg: 16,
    sm: 10,
    xl: 20,
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
