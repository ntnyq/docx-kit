import { BASIC_CODE } from './basic'
import { COMPREHENSIVE_CODE } from './comprehensive'
import { MULTISECTION_CODE } from './multi-section'
import { PLUGIN_CODE } from './plugins'

export interface Preset {
  code: string
  label: string
}

export { BASIC_CODE, COMPREHENSIVE_CODE, MULTISECTION_CODE, PLUGIN_CODE }

export const PRESETS: Preset[] = [
  { code: COMPREHENSIVE_CODE, label: 'Comprehensive' },
  { code: BASIC_CODE, label: 'Basic' },
  { code: PLUGIN_CODE, label: 'Plugins' },
  { code: MULTISECTION_CODE, label: 'Multi-Section' },
]

export const DEFAULT_CODE = COMPREHENSIVE_CODE
