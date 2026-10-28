/**
 * Academic preset — formal academic / thesis style.
 *
 * Characteristics:
 * - Times New Roman throughout
 * - Double-spaced body text with justified alignment
 * - Centered top-level headings
 * - Two-character first-line indent on paragraphs
 * - Centered images with generous vertical spacing
 *
 * @module presets/academic
 */

import type { DocxPreset } from '@docxkit/core'

/**
 * Academic preset — formal academic / thesis style.
 *
 * Uses Times New Roman throughout, double-spaced body text
 * with justified alignment and centered top-level headings.
 *
 * @remarks
 * - `config.defaults.paragraph` — base paragraph style (TNR, 12pt,
 *   double-spaced, justified, 24pt first-line indent)
 * - `config.defaults.image` — centered images with 12pt vertical margin
 * - `config.styles.h1`–`h6` — six heading levels (TNR, bold)
 * - `config.styles.p` — default paragraph class (inherits defaults)
 */
export const academicPreset: DocxPreset = {
  id: 'academic',
  name: 'Academic',
  config: {
    page: { margin: '25.4mm', size: 'A4' },
    defaults: {
      image: {
        marginBottom: 12,
        marginTop: 12,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
        lineHeight: 2,
        marginBottom: 0,
        marginTop: 0,
        textAlign: 'justify',
        textIndent: '24pt',
      },
      text: {
        color: '$colors.text',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
      },
    },
    styles: {
      h1: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 16,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 12,
        marginTop: 24,
        textAlign: 'center',
        textIndent: 0,
      },
      h2: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 14,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 18,
        textIndent: 0,
      },
      h3: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 14,
        textIndent: 0,
      },
      h4: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
        fontStyle: 'italic',
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 5,
        marginTop: 12,
        textIndent: 0,
      },
      h5: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 11,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 4,
        marginTop: 10,
        textIndent: 0,
      },
      h6: {
        color: '#333333',
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 10,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 4,
        marginTop: 8,
        textIndent: 0,
      },
      p: {
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
        lineHeight: 2,
        marginBottom: 0,
        marginTop: 0,
        textAlign: 'justify',
        textIndent: '24pt',
      },
    },
    theme: {
      colors: {
        accent: '#333333',
        border: '#d5d9df',
        muted: '#666666',
        primary: '#000000',
        surface: '#f4f6f8',
        text: '#252525',
      },
    },
  },
  description:
    'Formal academic style with Times New Roman, double spacing, and justified text.',
}
