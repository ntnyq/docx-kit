import { mkdir, rm, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { PNG } from 'pngjs'
import {
  calloutPlugin,
  createDocx,
  dataTablePlugin,
  defineStyles,
  letterheadPlugin,
  modernPreset,
  oceanTheme,
  signatureBlockPlugin,
  timelinePlugin,
} from '../../packages/docx-kit/src/node.ts'

const outputRoot = resolve(
  import.meta.dirname,
  '../..',
  'tmp/visual-regression',
)
const docxRoot = join(outputRoot, 'docx')
const fixtures = [
  {
    build: buildTypographyFixture,
    expectedPages: 1,
    name: '01-typography-and-styles',
  },
  {
    build: buildListsFixture,
    expectedPages: 1,
    name: '02-lists-and-links',
  },
  {
    build: buildTablesFixture,
    expectedPages: 1,
    name: '03-tables-and-spans',
  },
  {
    build: buildSectionsFixture,
    expectedPages: 2,
    name: '04-sections-and-headers',
  },
  {
    build: buildAnnotationsFixture,
    expectedPages: 1,
    name: '05-annotations-and-math',
  },
  {
    build: buildRevisionsFixture,
    expectedPages: 1,
    name: '06-revisions-and-textbox',
  },
  {
    build: buildThemeFixture,
    expectedPages: 1,
    name: '07-theme-and-preset',
  },
  {
    build: buildCalloutFixture,
    expectedPages: 1,
    name: '08-callouts-and-data',
  },
  {
    build: buildBusinessPluginsFixture,
    expectedPages: 2,
    name: '09-business-plugins',
  },
  {
    build: buildImageAndColumnsFixture,
    expectedPages: 1,
    name: '10-images-and-columns',
  },
]

await rm(outputRoot, { force: true, recursive: true })
await mkdir(docxRoot, { recursive: true })

for (const fixture of fixtures) {
  const path = join(docxRoot, `${fixture.name}.docx`)
  await fixture.build().save(path)
  console.log(`Generated ${path}`)
}

await writeFile(
  join(outputRoot, 'manifest.json'),
  `${JSON.stringify(
    fixtures.map(({ expectedPages, name }) => ({ expectedPages, name })),
    null,
    2,
  )}\n`,
)

function buildAnnotationsFixture() {
  return createBaseDocument('Annotations and Math')
    .p('The following controls exercise semantic OOXML content.')
    .checkbox({ checked: true, label: 'Compatibility review completed' })
    .checkbox({ checked: false, label: 'Release approved' })
    .comment({
      author: 'Ada Lovelace',
      comment: ['Confirm this wording before publication.'],
      date: '2026-07-26T00:00:00Z',
      initials: 'AL',
      children: [
        {
          style: { backgroundColor: '#FFF2CC' },
          text: 'This sentence has a review comment.',
          type: 'text',
        },
      ],
    })
    .footnote([
      'Footnotes should remain readable and anchored to the correct reference.',
    ])
    .math([
      {
        children: [{ text: 'a', type: 'text' }],
        type: 'radical',
      },
      { text: ' + ', type: 'text' },
      {
        denominator: [{ text: '2', type: 'text' }],
        numerator: [{ text: '1', type: 'text' }],
        type: 'fraction',
      },
    ])
    .p('The document continues after the semantic nodes without overlap.')
}

function buildBusinessPluginsFixture() {
  return createBaseDocument('Business Plugins', {}, false)
    .use(letterheadPlugin())
    .use(signatureBlockPlugin())
    .use(timelinePlugin())
    .plugin('letterhead', {
      address: '1 Document Way, Portland, OR',
      companyName: 'Docx Kit Labs',
      email: 'hello@example.com',
      phone: '+1 555 0100',
      tagline: 'Reliable document systems',
      website: 'example.com',
    })
    .h1('Business Plugins', { className: 'h1' })
    .h2('Delivery Timeline')
    .plugin('timeline', {
      accentColor: '2563EB',
      layout: 'left',
      events: [
        {
          date: 'Week 1',
          description: 'Confirm requirements and compatibility targets.',
          title: 'Planning',
        },
        {
          date: 'Week 2',
          description: 'Generate fixtures and validate document output.',
          title: 'Implementation',
        },
        {
          date: 'Week 3',
          description: 'Complete Word and LibreOffice review.',
          title: 'Verification',
        },
      ],
    })
    .pageBreak()
    .h2('Release Approval')
    .p('The undersigned approve the compatibility report.')
    .plugin('signatureBlock', {
      columns: 2,
      parties: [
        {
          date: 'Date: ____________',
          label: 'Engineering',
          name: 'Ada Lovelace',
        },
        {
          date: 'Date: ____________',
          label: 'Quality',
          name: 'Grace Hopper',
        },
      ],
    })
}

function buildCalloutFixture() {
  return createBaseDocument('Callouts and Data')
    .use(calloutPlugin())
    .use(dataTablePlugin())
    .plugin('callout', {
      content: 'All document-generation checks completed successfully.',
      title: 'Compatibility Status',
      type: 'success',
    })
    .plugin('callout', {
      content: 'Review visual differences before updating a baseline.',
      title: 'Baseline Policy',
      type: 'warning',
    })
    .h2('Quality Metrics')
    .plugin('dataTable', {
      align: { result: 'center', target: 'right' },
      striped: true,
      data: [
        { metric: 'Fixtures', result: '10', target: '10' },
        { metric: 'Platforms', result: '2', target: '2' },
        { metric: 'Critical diffs', result: '0', target: '0' },
      ],
      labels: {
        metric: 'Metric',
        result: 'Current',
        target: 'Target',
      },
    })
}

function buildImageAndColumnsFixture() {
  return createBaseDocument('Images and Columns')
    .image({
      data: createBanner(),
      height: 90,
      style: { marginBottom: 12, textAlign: 'center' },
      width: 420,
    })
    .p(
      'The generated image should be sharp, centered, and maintain its aspect ratio.',
    )
    .section({
      type: 'continuous',
      columns: {
        count: 2,
        separator: true,
        spacing: '12mm',
      },
    })
    .h2('Left Column')
    .p(
      'Column one contains enough text to exercise wrapping and consistent paragraph spacing.',
    )
    .columnBreak()
    .h2('Right Column')
    .p(
      'Column two begins after an explicit column break and should align with the first column.',
    )
}

function buildListsFixture() {
  return createBaseDocument('Lists and Links')
    .bookmark('navigation-target', ['Internal navigation target'])
    .p('The links below exercise both internal and external relationships.')
    .internalLink('navigation-target', 'Jump to the document bookmark')
    .hyperlink('https://example.com/docs', 'Open the external documentation')
    .h2('Nested Bullet List')
    .bulletList([
      'Top-level item',
      { level: 1, text: 'Nested item at level one' },
      { level: 2, text: 'Nested item at level two' },
      'Final top-level item',
    ])
    .h2('Numbered Procedure')
    .numberedList(
      [
        'Generate the DOCX fixture',
        'Convert the document with LibreOffice',
        'Compare rendered PNG pages',
      ],
      { numberingFormat: 'upperRoman', start: 3 },
    )
}

function buildRevisionsFixture() {
  return createBaseDocument('Revisions and Text Box', {
    features: { trackRevisions: true },
  })
    .p('The following lines exercise tracked revisions.')
    .insertedText({
      author: 'Grace Hopper',
      children: ['Inserted text should be visible as a tracked addition.'],
      date: '2026-07-26T00:00:00Z',
      revisionId: 101,
    })
    .deletedText({
      author: 'Grace Hopper',
      children: ['Deleted text remains available to revision-aware editors.'],
      date: '2026-07-26T00:00:00Z',
      revisionId: 102,
    })
    .thematicBreak({ style: { borderBottomColor: '#2563EB' } })
    .textBox({
      box: {
        height: '55pt',
        width: '260pt',
      },
      children: [
        {
          style: { color: '#1E3A8A', fontWeight: 'bold' },
          text: 'Positioned text box',
          type: 'text',
        },
      ],
    })
    .p('Body content should remain readable around the positioned text box.')
}

function buildSectionsFixture() {
  return createBaseDocument('Sections and Headers', {}, false)
    .section({
      page: { margin: '18mm', pageNumber: { start: 1 } },
      type: 'continuous',
      footer: {
        default: {
          children: [
            {
              text: 'Portrait section footer',
              type: 'paragraph',
              style: {
                color: '#64748B',
                fontSize: 9,
                textAlign: 'center',
              },
            },
          ],
        },
      },
      header: {
        default: {
          children: [
            {
              text: 'docx-kit compatibility suite',
              type: 'paragraph',
              style: {
                color: '#1E3A8A',
                fontSize: 10,
                borderBottom: {
                  color: '#2563EB',
                  style: 'single',
                  width: '1pt',
                },
              },
            },
          ],
        },
      },
    })
    .h1('Sections and Headers', { className: 'h1' })
    .p('This portrait section verifies rich headers and footers.')
    .section({
      type: 'nextPage',
      footer: {
        default: { children: ['Landscape section footer'] },
      },
      header: {
        default: { children: ['Landscape compatibility matrix'] },
      },
      page: {
        margin: '15mm',
        orientation: 'landscape',
        pageNumber: { format: 'upperRoman', start: 1 },
        size: 'A4',
      },
    })
    .h1('Landscape Section')
    .table({
      bordered: true,
      layout: 'fixed',
      columns: [
        { key: 'feature', title: 'Feature', width: '35%' },
        { key: 'word', title: 'Microsoft Word', width: '25%' },
        { key: 'libreoffice', title: 'LibreOffice', width: '25%' },
        { key: 'status', title: 'Status', width: '15%' },
      ],
      data: [
        {
          feature: 'Headers and footers',
          libreoffice: 'Reviewed',
          status: 'Pass',
          word: 'Reviewed',
        },
        {
          feature: 'Landscape orientation',
          libreoffice: 'Reviewed',
          status: 'Pass',
          word: 'Reviewed',
        },
      ],
      headerCellStyle: {
        backgroundColor: '#1E3A8A',
        color: '#FFFFFF',
        fontWeight: 'bold',
      },
    })
}

function buildTablesFixture() {
  return createBaseDocument('Tables and Spans')
    .p('Fixed-layout table with row spans, custom widths, and cell styles.')
    .table({
      bordered: true,
      layout: 'fixed',
      striped: true,
      width: '100%',
      columns: [
        { key: 'department', title: 'Department', width: '28%' },
        { key: 'owner', title: 'Owner', width: '24%' },
        { align: 'right', key: 'budget', title: 'Budget', width: '24%' },
        { align: 'center', key: 'status', title: 'Status', width: '24%' },
      ],
      data: [
        {
          _department_rowSpan: 2,
          budget: '$120,000',
          department: 'Engineering',
          owner: 'Ada',
          status: 'On track',
        },
        {
          budget: '$80,000',
          department: '',
          owner: 'Linus',
          status: 'At risk',
        },
        {
          budget: '$45,000',
          department: 'Design',
          owner: 'Dieter',
          status: 'On track',
        },
        {
          budget: '$60,000',
          department: 'Quality',
          owner: 'Grace',
          status: 'On track',
        },
      ],
      headerCellStyle: {
        backgroundColor: '#0F172A',
        color: '#FFFFFF',
        fontWeight: 'bold',
      },
      cellStyle: (_value, row) => ({
        backgroundColor: row.status === 'At risk' ? '#FFF2CC' : '#FFFFFF',
        verticalAlign: 'middle',
      }),
    })
    .h2('Border Variants')
    .table({
      borders: {
        bottom: { color: '#2563EB', style: 'double', width: '2pt' },
        top: { color: '#2563EB', style: 'double', width: '2pt' },
        insideHorizontal: {
          color: '#CBD5E1',
          style: 'dashed',
          width: '0.5pt',
        },
      },
      columns: [
        { key: 'label', title: 'Check' },
        { align: 'center', key: 'result', title: 'Result' },
      ],
      data: [
        { label: 'Outer border', result: 'Visible' },
        { label: 'Inner border', result: 'Visible' },
      ],
    })
}

function buildThemeFixture() {
  const styles = defineStyles({
    callout: {
      backgroundColor: '$colors.info',
      color: '#FFFFFF',
      padding: '$spacing.sm',
      borderLeft: {
        color: '$colors.primary',
        style: 'single',
        width: '3pt',
      },
    },
    subtitle: {
      color: '$colors.muted',
      fontSize: '$fontSize.lg',
      textAlign: 'center',
    },
    title: {
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      fontSize: 26,
      fontWeight: 'bold',
      textAlign: 'center',
    },
  })

  return createDocx({
    ...modernPreset.config,
    page: { margin: '20mm', size: 'A4' },
    theme: oceanTheme,
    metadata: {
      creator: 'docx-kit',
      title: 'Theme and Preset Visual Fixture',
    },
    styles: {
      ...modernPreset.config.styles,
      ...styles,
    },
  })
    .h1('Ocean Theme', { className: 'title' })
    .p('Modern preset combined with semantic design tokens.', {
      className: 'subtitle',
    })
    .p(
      'Theme tokens resolve into concrete colors, fonts, sizes, and spacing at compile time.',
      { className: 'callout' },
    )
    .h2('Consistent Hierarchy')
    .p(
      'Headings, body copy, and highlighted content should retain a coherent visual rhythm.',
    )
}

function buildTypographyFixture() {
  return createBaseDocument('Typography and Styles')
    .p([
      { style: { fontWeight: 'bold' }, text: 'Bold', type: 'text' },
      { text: ', ', type: 'text' },
      { style: { fontStyle: 'italic' }, text: 'italic', type: 'text' },
      { text: ', ', type: 'text' },
      {
        style: { color: '#DC2626', textDecoration: 'underline' },
        text: 'underlined red',
        type: 'text',
      },
      { text: ', and ', type: 'text' },
      {
        style: { backgroundColor: '#FEF08A', fontWeight: 'bold' },
        text: 'highlighted text',
        type: 'text',
      },
      { text: '.', type: 'text' },
    ])
    .h2('Paragraph Rhythm')
    .p(
      'This justified paragraph verifies line height, first-line indentation, margins, and predictable wrapping across editors.',
      {
        style: {
          lineHeight: 1.6,
          marginBottom: 12,
          textAlign: 'justify',
          textIndent: '24pt',
        },
      },
    )
    .p('Centered note with background and border.', {
      style: {
        backgroundColor: '#EFF6FF',
        border: { color: '#2563EB', style: 'single', width: '1pt' },
        padding: 8,
        textAlign: 'center',
      },
    })
    .h3('Keep-next Heading')
    .p('The heading and this paragraph should remain together.', {
      style: { keepLines: true, keepNext: true },
    })
}

function createBanner() {
  const png = new PNG({ height: 120, width: 560 })

  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      const offset = (png.width * y + x) * 4
      const progress = x / (png.width - 1)
      png.data[offset] = Math.round(37 + 212 * progress)
      png.data[offset + 1] = Math.round(99 + 56 * progress)
      png.data[offset + 2] = Math.round(235 - 190 * progress)
      png.data[offset + 3] = 255
    }
  }

  return PNG.sync.write(png)
}

function createBaseDocument(title, overrides = {}, includeTitle = true) {
  const styles = defineStyles({
    h1: {
      borderBottom: { color: '#2563EB', style: 'single', width: '1.5pt' },
      color: '#0F172A',
      fontSize: 24,
      fontWeight: 'bold',
      marginBottom: 12,
    },
    h2: {
      color: '#1D4ED8',
      fontSize: 17,
      fontWeight: 'bold',
      marginBottom: 8,
      marginTop: 16,
    },
    h3: {
      color: '#334155',
      fontSize: 13,
      fontWeight: 'bold',
      marginBottom: 6,
      marginTop: 12,
    },
  })

  const document = createDocx({
    page: { margin: '20mm', size: 'A4' },
    styles,
    defaults: {
      text: { fontFamily: 'Arial', fontSize: 11 },
      paragraph: {
        fontFamily: 'Arial',
        fontSize: 11,
        lineHeight: 1.35,
        marginBottom: 7,
      },
    },
    metadata: {
      creator: 'docx-kit visual regression',
      title,
    },
    ...overrides,
  })

  return includeTitle ? document.h1(title, { className: 'h1' }) : document
}
