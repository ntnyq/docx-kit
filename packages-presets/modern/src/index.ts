/**
 * Modern preset — clean, professional business style with blue accents.
 *
 * Characteristics:
 * - Dark navy headings with subtle blue accent underlines
 * - Calibri throughout for a contemporary feel
 * - Generous spacing and comfortable 1.5× line height
 * - Centered images with balanced margins
 *
 * @module presets/modern
 */

import type { DocxPreset } from '@docxkit/core'

/**
 * Modern preset — clean, professional business style with blue accents.
 *
 * Uses Calibri throughout, navy/blue headings with subtle accent
 * underlines, and generous 1.5× line height for comfortable reading.
 *
 * @remarks
 * - `config.defaults.paragraph` — base paragraph style (Calibri, 11pt, 1.5×)
 * - `config.defaults.image` — centered images with 10pt vertical margin
 * - `config.styles.h1`–`h6` — six heading levels with blue accent styling
 * - `config.styles.p` — default paragraph class (inherits defaults)
 */
export const modernPreset: DocxPreset = {
  id: 'modern',
  name: 'Modern',
  config: {
    defaults: {
      image: {
        marginBottom: 10,
        marginTop: 10,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'Calibri',
        fontSize: 11,
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 0,
      },
    },
    styles: {
      h1: {
        borderBottom: { color: '#2E75B6', style: 'single', width: '1.5pt' },
        color: '#1B2A4A',
        fontFamily: 'Calibri',
        fontSize: 26,
        fontWeight: 'bold',
        lineHeight: 1.3,
        marginBottom: 12,
        marginTop: 24,
      },
      h2: {
        color: '#2E75B6',
        fontFamily: 'Calibri',
        fontSize: 20,
        fontWeight: 'bold',
        lineHeight: 1.3,
        marginBottom: 10,
        marginTop: 20,
      },
      h3: {
        color: '#2E75B6',
        fontFamily: 'Calibri',
        fontSize: 16,
        fontWeight: 'bold',
        lineHeight: 1.35,
        marginBottom: 8,
        marginTop: 16,
      },
      h4: {
        color: '#404040',
        fontFamily: 'Calibri',
        fontSize: 14,
        fontWeight: 'bold',
        lineHeight: 1.35,
        marginBottom: 6,
        marginTop: 12,
      },
      h5: {
        color: '#404040',
        fontFamily: 'Calibri',
        fontSize: 12,
        fontWeight: 'bold',
        lineHeight: 1.35,
        marginBottom: 5,
        marginTop: 10,
      },
      h6: {
        color: '#666666',
        fontFamily: 'Calibri',
        fontSize: 11,
        fontWeight: 'bold',
        lineHeight: 1.35,
        marginBottom: 4,
        marginTop: 8,
      },
      p: {
        fontFamily: 'Calibri',
        fontSize: 11,
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 0,
      },
    },
  },
  description:
    'Clean business style with Calibri, blue accent headings, and generous spacing.',
}
