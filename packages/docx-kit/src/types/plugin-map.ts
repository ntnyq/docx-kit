/**
 * Built-in plugin type map — augments {@link BuiltinPluginMap} from `@docxkit/core`
 * with all built-in plugins so they are type-safe without explicit `.use()` calls.
 *
 * @module docx-kit/plugin-map
 */

import type { BadgeOptions } from '@docxkit/plugin-badge'
import type { BarcodeOptions } from '@docxkit/plugin-barcode'
import type { CalloutOptions } from '@docxkit/plugin-callout'
import type { ChangelogOptions } from '@docxkit/plugin-changelog'
import type { CodeBlockOptions } from '@docxkit/plugin-code-block'
import type { CoverPageOptions } from '@docxkit/plugin-cover-page'
import type { DataTableOptions } from '@docxkit/plugin-data-table'
import type { DividerOptions } from '@docxkit/plugin-divider'
import type { EChartsPluginOptions } from '@docxkit/plugin-echarts'
import type { InvoiceOptions } from '@docxkit/plugin-invoice'
import type { LetterheadOptions } from '@docxkit/plugin-letterhead'
import type { MeetingMinutesOptions } from '@docxkit/plugin-meeting-minutes'
import type { PageNumberOptions } from '@docxkit/plugin-page-number'
import type { PropertyTableOptions } from '@docxkit/plugin-property-table'
import type { QRCodePluginOptions } from '@docxkit/plugin-qrcode'
import type { SignatureBlockOptions } from '@docxkit/plugin-signature-block'
import type { TimelineOptions } from '@docxkit/plugin-timeline'
import type { TocOptions } from '@docxkit/plugin-toc'
import type { WatermarkOptions } from '@docxkit/plugin-watermark'

declare module '@docxkit/core' {
  export interface BuiltinPluginMap {
    badge: BadgeOptions
    barcode: BarcodeOptions
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
