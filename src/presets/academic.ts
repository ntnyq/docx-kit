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

import type { DocxPreset } from './types'

export const academicPreset: DocxPreset = {
  id: 'academic',
  name: 'Academic',
  config: {
    defaults: {
      image: {
        marginBottom: 12,
        marginTop: 12,
        textAlign: 'center',
      },
      paragraph: {
        fontFamily: 'Times New Roman',
        fontSize: 12,
        lineHeight: 2,
        marginBottom: 0,
        marginTop: 0,
        textAlign: 'justify',
        textIndent: '24pt',
      },
    },
    styles: {
      h1: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontSize: 16,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 12,
        marginTop: 24,
        textAlign: 'center',
      },
      h2: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontSize: 14,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 8,
        marginTop: 18,
      },
      h3: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontSize: 12,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 6,
        marginTop: 14,
      },
      h4: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontSize: 12,
        fontStyle: 'italic',
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 5,
        marginTop: 12,
      },
      h5: {
        color: '#000000',
        fontFamily: 'Times New Roman',
        fontSize: 11,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 4,
        marginTop: 10,
      },
      h6: {
        color: '#333333',
        fontFamily: 'Times New Roman',
        fontSize: 10,
        fontWeight: 'bold',
        lineHeight: 1.5,
        marginBottom: 4,
        marginTop: 8,
      },
      p: {
        fontFamily: 'Times New Roman',
        fontSize: 12,
        lineHeight: 2,
        marginBottom: 0,
        marginTop: 0,
        textAlign: 'justify',
        textIndent: '24pt',
      },
    },
  },
  description:
    'Formal academic style with Times New Roman, double spacing, and justified text.',
}
