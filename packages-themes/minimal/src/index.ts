/**
 * Minimal theme — clean, modern grayscale palette with blue accents.
 *
 * @module themes/minimal
 */

import type { DocxTheme } from '@docxkit/core'

/**
 * Minimal theme — clean, modern grayscale palette with subtle blue accents.
 *
 * Uses Inter/Arial sans-serif fonts and a restrained grayscale
 * color palette suitable for business documents and reports.
 *
 * @remarks
 * - `colors` — 12 semantic color tokens (primary, accent, success, etc.)
 * - `fonts` — body/heading/code font stacks
 * - `fontSize` — 5-step type scale (xs/sm/base/lg/xl)
 * - `spacing` — 5-step spacing scale (xs/sm/md/lg/xl) in half-points
 */
export const minimalTheme: DocxTheme & {
  description: string
  id: string
  name: string
} = {
  description: 'Clean grayscale palette with subtle blue accents.',
  id: 'minimal',
  name: 'Minimal',
  colors: {
    accent: '#2563eb',
    background: '#ffffff',
    border: '#e5e7eb',
    danger: '#dc2626',
    info: '#0ea5e9',
    muted: '#6b7280',
    primary: '#111827',
    secondary: '#374151',
    success: '#16a34a',
    surface: '#f9fafb',
    text: '#1f2937',
    warning: '#d97706',
  },
  fonts: {
    body: 'Inter, Arial, sans-serif',
    code: 'JetBrains Mono, monospace',
    heading: 'Inter, Arial, sans-serif',
  },
  fontSize: {
    base: 11,
    lg: 14,
    sm: 9,
    xl: 18,
    xs: 8,
  },
  spacing: {
    lg: 24,
    md: 16,
    sm: 8,
    xl: 32,
    xs: 4,
  },
}
