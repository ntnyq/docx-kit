/**
 * Classic preset — formal, official-document style inspired by
 * Chinese government document standards (GB/T 9704).
 *
 * Characteristics:
 * - Black headings in SimHei / KaiTi
 * - Body text in SimSun with two-character first-line indent
 * - 1.5× line height throughout
 * - Centered images with generous spacing
 *
 * @module presets/classic
 */

import type { DocxPreset } from './types'

export const classicPreset: DocxPreset = {
  id: 'classic',
  name: 'Classic',
  config: {
    defaults: {
      image: {
        marginBottom: 8,
        marginTop: 8,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'SimSun',
        fontSize: 14,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 0,
        textIndent: '28pt',
      },
    },
    styles: {
      h1: {
        color: '#000000',
        fontFamily: 'SimHei',
        fontSize: 22,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 12,
        marginTop: 20,
        textAlign: 'center',
      },
      h2: {
        color: '#000000',
        fontFamily: 'SimHei',
        fontSize: 16,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 16,
      },
      h3: {
        color: '#000000',
        fontFamily: 'KaiTi',
        fontSize: 16,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 12,
      },
      h4: {
        color: '#000000',
        fontFamily: 'KaiTi',
        fontSize: 14,
        fontWeight: 'bold',
        lineHeight: 1.4,
        marginBottom: 5,
        marginTop: 10,
      },
      h5: {
        color: '#000000',
        fontFamily: 'SimSun',
        fontSize: 14,
        fontWeight: 'bold',
        lineHeight: 1.4,
        marginBottom: 4,
        marginTop: 8,
      },
      h6: {
        color: '#000000',
        fontFamily: 'SimSun',
        fontSize: 12,
        fontWeight: 'bold',
        lineHeight: 1.4,
        marginBottom: 4,
        marginTop: 8,
      },
      p: {
        fontFamily: 'SimSun',
        fontSize: 14,
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 0,
        textIndent: '28pt',
      },
    },
  },
  description:
    'Formal official-document style with SimHei headings, SimSun body, and two-character indent.',
}
