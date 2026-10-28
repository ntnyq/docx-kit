import { unindent } from '@ntnyq/utils'

export const BASIC_CODE = unindent(`
  import { DocxBuilder } from 'docx-kit'

  const doc = new DocxBuilder()
    .h1('Hello from docx-kit!')
    .p('This is a basic document generated in the browser.')
    .p('Edit the code on the left and click Run to regenerate.')
    .bulletList([
      'Fast and fluent API',
      'CSS-like styling',
      'Browser + Node.js',
      'Plugin system',
    ])
    .hyperlink('https://github.com/ntnyq/docx-kit', 'Visit docx-kit on GitHub')
    .numberedList([
      'Write code in the editor',
      'Click Run to compile',
      'Download the .docx file',
    ])
    .table({
      columns: [
        { key: 'feature', title: 'Feature', width: '60%' },
        { key: 'status', title: 'Status', width: '40%' },
      ],
      data: [
        { feature: 'Builder API', status: 'Working' },
        { feature: 'CSS-like Styles', status: 'Working' },
        { feature: 'Plugins', status: 'Working' },
        { feature: 'Multi-Section', status: 'Working' },
      ],
    })
    .p('Happy documenting!')

  doc.toBlob()
`)
