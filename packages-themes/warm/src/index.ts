/**
 * Warm theme — earthy tones with amber and terracotta accents.
 *
 * @module themes/warm
 */

import type { DocxTheme } from '@docxkit/core'

/**
 * Warm theme — earthy tones with amber and terracotta accents.
 *
 * Uses Garamond/Georgia serif fonts and warm color palette
 * (amber, terracotta, warm gray) suitable for invitations
 * and personal correspondence.
 *
 * @remarks
 * - `colors` — 12 semantic color tokens (primary, accent, success, etc.)
 * - `fonts` — body/heading/code font stacks
 * - `fontSize` — 5-step type scale (xs/sm/base/lg/xl)
 * - `spacing` — 5-step spacing scale (xs/sm/md/lg/xl) in half-points
 */
export const warmTheme: DocxTheme & {
  description: string
  id: string
  name: string
} = {
  description: 'Earthy tones with amber and terracotta accents.',
  id: 'warm',
  name: 'Warm',
  colors: {
    accent: '#d97706',
    background: '#fffbeb',
    border: '#fde68a',
    danger: '#b91c1c',
    info: '#7c3aed',
    muted: '#78716c',
    primary: '#44403c',
    secondary: '#57534e',
    success: '#15803d',
    surface: '#fef3c7',
    text: '#292524',
    warning: '#ea580c',
  },
  fonts: {
    body: 'Garamond, Georgia, serif',
    code: 'JetBrains Mono, monospace',
    heading: 'Garamond, Georgia, serif',
  },
  fontSize: {
    base: 12,
    lg: 15,
    sm: 10,
    xl: 19,
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
