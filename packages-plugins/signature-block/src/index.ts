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

import { createPluginStyle, definePlugin, DocxKitError } from '@docxkit/core'
import {
  AlignmentType,
  Paragraph,
  Table,
  TableBorders,
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
   * Maximum columns in the signature grid, capped at the number of parties.
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
    render(options, context) {
      const pluginStyles = createPluginStyle(context.style)
      const { columns = 2, parties } = options

      if (!Number.isSafeInteger(columns) || columns <= 0) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Signature block columns must be a positive integer',
        )
      }

      if (!parties.length) {
        return new Paragraph(
          pluginStyles.paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun(
                pluginStyles.text({ color: '999999', text: '(no signatures)' }),
              ),
            ],
          }),
        )
      }

      const columnCount = Math.min(columns, parties.length)

      // Build cell for a single party
      function buildPartyCell(party: SignatureParty): TableCell {
        const labelPara = new Paragraph(
          pluginStyles.paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun(pluginStyles.text({ bold: true, text: party.label })),
            ],
          }),
        )

        const sigPara = party.name
          ? new Paragraph(
              pluginStyles.paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    pluginStyles.text({ text: party.name, underline: {} }),
                  ),
                ],
              }),
            )
          : new Paragraph(
              pluginStyles.paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    pluginStyles.text({
                      text: '_____________________________',
                      underline: {},
                    }),
                  ),
                ],
              }),
            )

        const datePara = new Paragraph(
          pluginStyles.paragraph({
            children: [
              new TextRun(
                pluginStyles.text({
                  color: '666666',
                  size: 18,
                  text: party.date ?? '',
                }),
              ),
            ],
          }),
        )

        return new TableCell(
          pluginStyles.cell({ children: [labelPara, sigPara, datePara] }),
        )
      }

      // Bound padding by the actual input rather than a caller-supplied count.
      const partyRows: SignatureParty[][] = []
      for (let i = 0; i < parties.length; i += columnCount) {
        partyRows[partyRows.length] = parties.slice(i, i + columnCount)
      }

      const rows = partyRows.map(rowParties => {
        let rowCells = rowParties.map(buildPartyCell)

        // Pad the final row without sharing mutable docx cell objects.
        if (rowCells.length < columnCount) {
          const padCount = columnCount - rowCells.length
          const padding = Array.from(
            { length: padCount },
            () =>
              new TableCell(
                pluginStyles.cell({
                  children: [
                    new Paragraph(pluginStyles.paragraph({ text: '' })),
                  ],
                }),
              ),
          )
          rowCells = [...rowCells, ...padding]
        }

        return new TableRow({ children: rowCells })
      })

      return new Table({
        borders: TableBorders.NONE,
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
