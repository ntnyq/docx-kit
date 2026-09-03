/**
 * Signature Block plugin — renders signature lines for contracts and approvals.
 *
 * Each party gets a cell in a borderless table with a bold label,
 * a signature line (underlined placeholder), and an optional date.
 *
 * @module plugins/signature-block
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(signatureBlockPlugin())
 *   .plugin('signatureBlock', {
 *     parties: [
 *       { label: '甲方（盖章）', date: '2026年  月  日' },
 *       { label: '乙方（盖章）', date: '2026年  月  日' },
 *     ],
 *     columns: 2,
 *   })
 *   .save('signature.docx')
 * ```
 */

import { definePlugin, DocxKitError } from '@docxkit/core'
import {
  AlignmentType,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'

/**
 * Options for the SignatureBlock plugin.
 */
export interface SignatureBlockOptions {
  /**
   * The signing parties.
   */
  parties: SignatureParty[]
  /**
   * Number of columns in the signature grid.
   * @default 2
   */
  columns?: number
}

/**
 * A single party in a signature block.
 */
export interface SignatureParty {
  /**
   * Party label (e.g. "甲方（盖章）").
   */
  label: string
  /**
   * Pre-filled date (e.g. "2026年  月  日").
   */
  date?: string
  /**
   * Pre-filled name (shown underlined when provided).
   */
  name?: string
}

/**
 * Create a SignatureBlock plugin instance.
 *
 * @returns A configured DocxPlugin for `'signatureBlock'`
 */
export function signatureBlockPlugin() {
  return definePlugin<'signatureBlock', SignatureBlockOptions>({
    name: 'signatureBlock',
    render(options) {
      const { columns = 2, parties } = options

      if (!Number.isSafeInteger(columns) || columns <= 0) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Signature block columns must be a positive integer',
        )
      }

      if (!parties.length) {
        return new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ color: '999999', text: '(no signatures)' })],
        })
      }

      // Build cell for a single party
      function buildPartyCell(party: SignatureParty): TableCell {
        const labelPara = new Paragraph({
          children: [new TextRun({ bold: true, text: party.label })],
          spacing: { after: 200 },
        })

        const sigPara = party.name
          ? new Paragraph({
              children: [new TextRun({ text: party.name, underline: {} })],
              spacing: { after: 80 },
            })
          : new Paragraph({
              spacing: { after: 80 },
              children: [
                new TextRun({
                  text: '_____________________________',
                  underline: {},
                }),
              ],
            })

        const datePara = new Paragraph({
          children: [
            new TextRun({ color: '666666', size: 18, text: party.date ?? '' }),
          ],
        })

        return new TableCell({ children: [labelPara, sigPara, datePara] })
      }

      // Arrange parties into rows of `columns` each
      const partyRows: SignatureParty[][] = []
      for (let i = 0; i < parties.length; i += columns) {
        partyRows[partyRows.length] = parties.slice(i, i + columns)
      }

      const rows = partyRows.map(rowParties => {
        let rowCells = rowParties.map(buildPartyCell)

        // Pad last row if needed so all rows have `columns` cells
        if (rowCells.length < columns) {
          const emptyCell = new TableCell({ children: [new Paragraph('')] })
          const padCount = columns - rowCells.length
          const padding = Array.from({ length: padCount }).map(() => emptyCell)
          rowCells = [...rowCells, ...padding]
        }

        return new TableRow({ children: rowCells })
      })

      return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: {
          bottom: 40,
          left: 100,
          right: 100,
          top: 40,
        },
      })
    },
  })
}
