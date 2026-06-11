import { unindent } from '@ntnyq/utils'

export const PLUGIN_CODE = unindent(`
  import { DocxBuilder, definePlugin } from 'docx-kit'
  import { Paragraph, TextRun, BorderStyle, AlignmentType, HeadingLevel } from 'docx'

  // ─── Plugin: Section Divider ──────────────────────────────────────────
  const dividerPlugin = definePlugin<'divider', { label?: string }>({
    name: 'divider',
    async render(opts) {
      return [new Paragraph({
        text: opts.label ? \`── \${opts.label} ──\` : '──',
        alignment: AlignmentType.CENTER,
        spacing: { before: 300, after: 300 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999' } },
      })]
    },
  })

  // ─── Plugin: Info Box ─────────────────────────────────────────────────
  const infoBoxPlugin = definePlugin<'infoBox', { title: string; body: string }>({
    name: 'infoBox',
    async render(opts) {
      return [
        new Paragraph({ text: \`ℹ \${opts.title}\`, bold: true, spacing: { before: 200, after: 80 } }),
        new Paragraph({ text: opts.body, indent: { left: 360 }, italics: true, spacing: { before: 0, after: 200 } }),
      ]
    },
  })

  // ─── Plugin: Signature Block ──────────────────────────────────────────
  const signaturePlugin = definePlugin<'signature', { name: string; title?: string; date?: string }>({
    name: 'signature',
    async render(opts) {
      const today = opts.date || new Date().toISOString().slice(0, 10)
      return [
        new Paragraph({ text: '', spacing: { before: 400 } }),
        new Paragraph({ text: '_________________________', spacing: { after: 40 } }),
        new Paragraph({ text: opts.name, bold: true }),
        ...(opts.title ? [new Paragraph({ text: opts.title, italics: true })] : []),
        new Paragraph({ text: \`Date: \${today}\`, spacing: { before: 80 } }),
      ]
    },
  })

  // ─── Build Document ───────────────────────────────────────────────────
  const doc = new DocxBuilder({ page: { size: 'A4', margin: '2cm' } })
    .use(dividerPlugin)
    .use(infoBoxPlugin)
    .use(signaturePlugin)

    .h1('Plugin System Demo')
    .p('This document demonstrates three custom plugins registered at runtime.')

    .plugin('divider', { label: 'INFO BOXES' })

    .plugin('infoBox', {
      title: 'How Plugins Work',
      body: 'Each plugin defines a name, optional setup hook, and a render function. The render function receives user options and a PluginRenderContext with access to the docx library API for creating native OOXML elements.',
    })

    .plugin('infoBox', {
      title: 'Type Safety',
      body: 'Plugins use TypeScript const generics so that the builder\\'s .plugin() method is fully type-safe. Register a plugin with .use() and the option types propagate through the entire chain.',
    })

    .plugin('divider', { label: 'SIGNATURE' })

    .h2('Approval Section')
    .p('The signature block below is rendered by the signature plugin:')

    .plugin('signature', {
      name: 'Zhang Wei',
      title: 'Senior Engineer',
    })

    .p('')
    .p('Each plugin is defined once and can be reused throughout the document.')

  doc.toBlob()
`)
