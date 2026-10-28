import { BASIC_CODE } from './basic'
import { CHECKLIST_CODE } from './checklist'
import { COMPREHENSIVE_CODE } from './comprehensive'
import { MULTISECTION_CODE } from './multi-section'
import { PLUGIN_CODE } from './plugins'
import { RICH_CONTENT_CODE } from './rich-content'
import { THEME_CODE } from './theme'

export interface Preset {
  code: string
  id: string
  label: string
}

export {
  BASIC_CODE,
  CHECKLIST_CODE,
  COMPREHENSIVE_CODE,
  MULTISECTION_CODE,
  PLUGIN_CODE,
  RICH_CONTENT_CODE,
  THEME_CODE,
}

export const PRESETS: Preset[] = [
  { code: COMPREHENSIVE_CODE, id: 'comprehensive', label: 'Comprehensive' },
  { code: BASIC_CODE, id: 'basic', label: 'Basic' },
  { code: PLUGIN_CODE, id: 'plugins', label: 'Plugins' },
  { code: CHECKLIST_CODE, id: 'checklist', label: 'Checklist' },
  { code: MULTISECTION_CODE, id: 'multi-section', label: 'Multi-Section' },
  { code: THEME_CODE, id: 'theme', label: 'Theme' },
  { code: RICH_CONTENT_CODE, id: 'rich-content', label: 'Rich Content' },
]

export const DEFAULT_CODE = COMPREHENSIVE_CODE
