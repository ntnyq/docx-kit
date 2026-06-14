/**
 * Built-in plugin type map — augments {@link BuiltinPluginMap} from @docxkit/core
 * with all built-in plugins so they are type-safe without explicit `.use()` calls.
 *
 * @module docx-kit/plugin-map
 */

import type { CalloutOptions } from '@docxkit/plugin-callout'
import type { CodeBlockOptions } from '@docxkit/plugin-code-block'
import type { CoverPageOptions } from '@docxkit/plugin-cover-page'
import type { DataTableOptions } from '@docxkit/plugin-data-table'
import type { EChartsPluginOptions } from '@docxkit/plugin-echarts'
import type { MeetingMinutesOptions } from '@docxkit/plugin-meeting-minutes'
import type { PageNumberOptions } from '@docxkit/plugin-page-number'
import type { PropertyTableOptions } from '@docxkit/plugin-property-table'
import type { QRCodePluginOptions } from '@docxkit/plugin-qrcode'
import type { SignatureBlockOptions } from '@docxkit/plugin-signature-block'
import type { TimelineOptions } from '@docxkit/plugin-timeline'
import type { WatermarkOptions } from '@docxkit/plugin-watermark'
import type { BadgeOptions } from '../../../../packages-plugins/badge/src'
import type { ChangelogOptions } from '../../../../packages-plugins/changelog/src'
import type { DividerOptions } from '../../../../packages-plugins/divider/src'
import type { InvoiceOptions } from '../../../../packages-plugins/invoice/src'
import type { LetterheadOptions } from '../../../../packages-plugins/letterhead/src'
import type { TocOptions } from '../../../../packages-plugins/toc/src'

declare module '@docxkit/core' {
  export interface BuiltinPluginMap {
    badge: BadgeOptions
    callout: CalloutOptions
    changelog: ChangelogOptions
    codeBlock: CodeBlockOptions
    coverPage: CoverPageOptions
    dataTable: DataTableOptions
    divider: DividerOptions
    echarts: EChartsPluginOptions
    invoice: InvoiceOptions
    letterhead: LetterheadOptions
    meetingMinutes: MeetingMinutesOptions
    pageNumber: PageNumberOptions
    propertyTable: PropertyTableOptions
    qrcode: QRCodePluginOptions
    signatureBlock: SignatureBlockOptions
    timeline: TimelineOptions
    toc: TocOptions
    watermark: WatermarkOptions
  }
}
