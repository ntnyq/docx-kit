import { BASIC_CODE } from './basic'
import { COMPREHENSIVE_CODE } from './comprehensive'
import { MULTISECTION_CODE } from './multi-section'
import { PLUGIN_CODE } from './plugins'
import { RICH_CONTENT_CODE } from './rich-content'
import { THEME_CODE } from './theme'

export interface Preset {
  code: string
  label: string
}

export {
  BASIC_CODE,
  COMPREHENSIVE_CODE,
  MULTISECTION_CODE,
  PLUGIN_CODE,
  RICH_CONTENT_CODE,
  THEME_CODE,
}

export const PRESETS: Preset[] = [
  { code: COMPREHENSIVE_CODE, label: 'Comprehensive' },
  { code: BASIC_CODE, label: 'Basic' },
  { code: PLUGIN_CODE, label: 'Plugins' },
  { code: MULTISECTION_CODE, label: 'Multi-Section' },
  { code: THEME_CODE, label: 'Theme' },
  { code: RICH_CONTENT_CODE, label: 'Rich Content' },
]

export const DEFAULT_CODE = COMPREHENSIVE_CODE
