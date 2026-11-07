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

import type {
  CellStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface SignatureBlockConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: SignatureBlockStyles
}

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
  /**
   * Per-part styles for this invocation.
   */
  styles?: SignatureBlockStyles
}

/**
 * Independently configurable SignatureBlock parts.
 */
export interface SignatureBlockStyles {
  /**
   * Cell shading, borders and padding.
   */
  cell?: Omit<CellStyleRule, 'height'>
  /**
   * Date run typography.
   */
  date?: TextStyleRule
  /**
   * Date paragraph alignment, spacing and flow.
   */
  dateParagraph?: ParagraphStyleRule
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Label run typography.
   */
  label?: TextStyleRule
  /**
   * Label paragraph alignment, spacing and flow.
   */
  labelParagraph?: ParagraphStyleRule
  /**
   * Padding paragraph alignment, spacing and flow.
   */
  paddingParagraph?: ParagraphStyleRule
  /**
   * Signature run typography.
   */
  signature?: TextStyleRule
  /**
   * Signature paragraph alignment, spacing and flow.
   */
  signatureParagraph?: ParagraphStyleRule
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
export function signatureBlockPlugin(pluginConfig: SignatureBlockConfig = {}) {
  return definePlugin<'signatureBlock', SignatureBlockOptions>({
    name: 'signatureBlock',
    render(options, context) {
      const partStyle = (...parts: (keyof SignatureBlockStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const cellStyle = partStyle('cell')
      const dateStyle = partStyle('date')
      const dateParagraphStyle = partStyle('dateParagraph')
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const labelStyle = partStyle('label')
      const labelParagraphStyle = partStyle('labelParagraph')
      const paddingParagraphStyle = partStyle('paddingParagraph')
      const signatureStyle = partStyle('signature')
      const signatureParagraphStyle = partStyle('signatureParagraph')
      const { columns = 2, parties } = options

      if (!Number.isSafeInteger(columns) || columns <= 0) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Signature block columns must be a positive integer',
        )
      }

      if (!parties.length) {
        return new Paragraph(
          emptyParagraphStyle.paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun(
                emptyTextStyle.text({
                  color: '999999',
                  text: '(no signatures)',
                }),
              ),
            ],
          }),
        )
      }

      const columnCount = Math.min(columns, parties.length)

      // Build cell for a single party
      function buildPartyCell(party: SignatureParty): TableCell {
        const labelPara = new Paragraph(
          labelParagraphStyle.paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun(labelStyle.text({ bold: true, text: party.label })),
            ],
          }),
        )

        const sigPara = party.name
          ? new Paragraph(
              signatureParagraphStyle.paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    signatureStyle.text({ text: party.name, underline: {} }),
                  ),
                ],
              }),
            )
          : new Paragraph(
              signatureParagraphStyle.paragraph({
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    signatureStyle.text({
                      text: '_____________________________',
                      underline: {},
                    }),
                  ),
                ],
              }),
            )

        const datePara = new Paragraph(
          dateParagraphStyle.paragraph({
            children: [
              new TextRun(
                dateStyle.text({
                  color: '666666',
                  size: 18,
                  text: party.date ?? '',
                }),
              ),
            ],
          }),
        )

        return new TableCell(
          cellStyle.cell({ children: [labelPara, sigPara, datePara] }),
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
                cellStyle.cell({
                  children: [
                    new Paragraph(
                      paddingParagraphStyle.paragraph({ text: '' }),
                    ),
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
