/**
 * Modern preset — clean, professional business style with blue accents.
 *
 * Characteristics:
 * - Dark navy headings with subtle restrained blue accents
 * - Calibri throughout for a contemporary feel
 * - Generous spacing and comfortable 1.35× line height
 * - Centered images with balanced margins
 *
 * @module presets/modern
 */

import type { DocxPreset } from '@docxkit/core'

/**
 * Modern preset — clean, professional business style with blue accents.
 *
 * Uses Calibri throughout, navy/blue headings, and generous 1.35× line height for comfortable reading.
 *
 * @remarks
 * - `config.defaults.paragraph` — base paragraph style (Calibri, 11pt, 1.35×)
 * - `config.defaults.image` — centered images with 10pt vertical margin
 * - `config.styles.h1`–`h6` — six heading levels with blue accent styling
 * - `config.styles.p` — default paragraph class (inherits defaults)
 */
export const modernPreset: DocxPreset = {
  id: 'modern',
  name: 'Modern',
  config: {
    page: { margin: '22mm 24mm', size: 'A4' },
    defaults: {
      image: {
        marginBottom: 10,
        marginTop: 10,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 11,
        lineHeight: 1.35,
        marginBottom: 8,
        marginTop: 0,
      },
      text: {
        color: '$colors.text',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 11,
      },
    },
    styles: {
      h1: {
        color: '$colors.primary',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 24,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.3,
        marginBottom: 12,
        marginTop: 0,
        textIndent: 0,
      },
      h2: {
        color: '$colors.accent',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 16,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.3,
        marginBottom: 10,
        marginTop: 20,
        textIndent: 0,
      },
      h3: {
        color: '$colors.accent',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 13,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.35,
        marginBottom: 8,
        marginTop: 16,
        textIndent: 0,
      },
      h4: {
        color: '#404040',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 12,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.35,
        marginBottom: 6,
        marginTop: 12,
        textIndent: 0,
      },
      h5: {
        color: '#404040',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 12,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.35,
        marginBottom: 5,
        marginTop: 10,
        textIndent: 0,
      },
      h6: {
        color: '#666666',
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 11,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.35,
        marginBottom: 4,
        marginTop: 8,
        textIndent: 0,
      },
      p: {
        fontFamily: 'Calibri',
        fontFamilyEastAsia: 'Microsoft YaHei',
        fontSize: 11,
        lineHeight: 1.35,
        marginBottom: 8,
        marginTop: 0,
      },
    },
    theme: {
      colors: {
        accent: '#2E75B6',
        border: '#d5d9df',
        muted: '#666666',
        primary: '#1B2A4A',
        surface: '#f4f6f8',
        text: '#252525',
      },
    },
  },
  description:
    'Clean business style with Calibri, blue accent headings, and generous spacing.',
}
