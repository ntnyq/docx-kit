/**
 * Built-in plugin type map — augments {@link BuiltinPluginMap} from @docxkit/core
 * with all 12 built-in plugins so they are type-safe without explicit `.use()` calls.
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

declare module '@docxkit/core' {
  export interface BuiltinPluginMap {
    callout: CalloutOptions
    codeBlock: CodeBlockOptions
    coverPage: CoverPageOptions
    dataTable: DataTableOptions
    echarts: EChartsPluginOptions
    meetingMinutes: MeetingMinutesOptions
    pageNumber: PageNumberOptions
    propertyTable: PropertyTableOptions
    qrcode: QRCodePluginOptions
    signatureBlock: SignatureBlockOptions
    timeline: TimelineOptions
    watermark: WatermarkOptions
  }
}
