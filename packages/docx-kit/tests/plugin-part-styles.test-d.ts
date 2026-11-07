import { expectTypeOf } from 'vitest'
import { calloutPlugin, createDocx, dataTablePlugin } from '../src/node'
import type {
  CalloutStyles as BrowserCalloutStyles,
  DataTableStyles as BrowserDataTableStyles,
} from '../src/browser'
import type {
  CalloutConfig,
  CalloutOptions,
  CalloutStyles,
  DataTableConfig,
  DataTableOptions,
  DataTableStyles,
} from '../src/node'

expectTypeOf<CalloutConfig['styles']>().toEqualTypeOf<
  CalloutStyles | undefined
>()
expectTypeOf<CalloutOptions['styles']>().toEqualTypeOf<
  CalloutStyles | undefined
>()
expectTypeOf<DataTableConfig['styles']>().toEqualTypeOf<
  DataTableStyles | undefined
>()
expectTypeOf<DataTableOptions['styles']>().toEqualTypeOf<
  DataTableStyles | undefined
>()
expectTypeOf<BrowserCalloutStyles>().toEqualTypeOf<CalloutStyles>()
expectTypeOf<BrowserDataTableStyles>().toEqualTypeOf<DataTableStyles>()

const doc = createDocx()
  .use(
    calloutPlugin({
      styles: { cell: { margin: '$spacing.md' }, title: { fontSize: 18 } },
    }),
  )
  .use(dataTablePlugin({ styles: { headerText: { bold: false } } }))

doc.plugin('callout', {
  content: 'Body',
  styles: { title: { color: '#123456' } },
  type: 'info',
})
doc.plugin('dataTable', {
  data: [],
  styles: { bodyParagraph: { textAlign: 'right' } },
})
// @ts-expect-error Part names are checked on factory defaults.
calloutPlugin({ styles: { titel: {} } })
// @ts-expect-error Typography does not belong on a cell part.
dataTablePlugin({ styles: { headerCell: { fontSize: 12 } } })
doc.plugin('callout', {
  content: 'Body',
  // @ts-expect-error Callout title parts only accept text style properties.
  styles: { title: { textAlign: 'center' } },
  type: 'info',
})
// @ts-expect-error Invocation parts retain plugin-specific inference.
doc.plugin('dataTable', { data: [], styles: { title: {} } })
doc.plugin('dataTable', {
  data: [],
  // @ts-expect-error Font sizes must use supported units.
  styles: { bodyText: { fontSize: 'huge' } },
})

// @ts-expect-error Row height is not a supported cell style property.
dataTablePlugin({ styles: { bodyCell: { height: 20 } } })
