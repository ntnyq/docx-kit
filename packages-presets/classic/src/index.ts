/**
 * Classic preset — formal, official-document style for Chinese reports and official correspondence.
 *
 * Characteristics:
 * - Black headings in SimHei / KaiTi
 * - Body text in SimSun with two-character first-line indent
 * - 1.5× line height throughout
 * - Centered images with generous spacing
 *
 * @module presets/classic
 */

import type { DocxPreset } from '@docxkit/core'

/**
 * Classic preset — formal, official-document style for Chinese reports and official correspondence.
 *
 * Uses black SimHei/KaiTi headings, SimSun body text with
 * two-character first-line indent, and 1.5× line height.
 *
 * @remarks
 * - `config.defaults.paragraph` — base paragraph style (SimSun, 14pt,
 *   1.5× line height, 28pt first-line indent)
 * - `config.defaults.image` — centered images with 8pt vertical margin
 * - `config.styles.h1`–`h6` — six heading levels (SimHei, KaiTi, SimSun)
 * - `config.styles.p` — default paragraph class (inherits defaults)
 */
export const classicPreset: DocxPreset = {
  id: 'classic',
  name: 'Classic',
  config: {
    page: { margin: '25mm 28mm', size: 'A4' },
    defaults: {
      image: {
        marginBottom: 8,
        marginTop: 8,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'SimSun',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 14,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 0,
        textIndent: '28pt',
      },
      text: {
        color: '$colors.text',
        fontFamily: 'SimSun',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 14,
      },
    },
    styles: {
      h1: {
        color: '#000000',
        fontFamily: 'SimHei',
        fontFamilyEastAsia: 'SimHei',
        fontSize: 22,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 12,
        marginTop: 20,
        textAlign: 'center',
        textIndent: 0,
      },
      h2: {
        color: '#000000',
        fontFamily: 'SimHei',
        fontFamilyEastAsia: 'SimHei',
        fontSize: 16,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 16,
        textIndent: 0,
      },
      h3: {
        color: '#000000',
        fontFamily: 'KaiTi',
        fontFamilyEastAsia: 'KaiTi',
        fontSize: 16,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 12,
        textIndent: 0,
      },
      h4: {
        color: '#000000',
        fontFamily: 'KaiTi',
        fontFamilyEastAsia: 'KaiTi',
        fontSize: 14,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.4,
        marginBottom: 5,
        marginTop: 10,
        textIndent: 0,
      },
      h5: {
        color: '#000000',
        fontFamily: 'SimSun',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 14,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.4,
        marginBottom: 4,
        marginTop: 8,
        textIndent: 0,
      },
      h6: {
        color: '#000000',
        fontFamily: 'SimSun',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 12,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.4,
        marginBottom: 4,
        marginTop: 8,
        textIndent: 0,
      },
      p: {
        fontFamily: 'SimSun',
        fontFamilyEastAsia: 'SimSun',
        fontSize: 14,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 0,
        textIndent: '28pt',
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
    'Formal official-document style with SimHei headings, SimSun body, and two-character indent.',
}
