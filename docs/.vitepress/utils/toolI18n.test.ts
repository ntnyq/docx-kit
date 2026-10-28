import { describe, expect, it } from 'vitest'
import { shallowRef } from 'vue'
import { createToolI18n, resolveToolLocale } from './toolI18n'

describe('documentation tool language', () => {
  it('updates controls and accessible labels when the document language changes', () => {
    const language = shallowRef('en-US')
    const { formatMessage, locale, messages } = createToolI18n(language)

    expect(messages.value.editorTitle).toBe('Code Editor')
    expect(messages.value.studioTitle).toBe('Theme Studio')

    language.value = 'zh-CN'
    expect(locale.value).toBe('zh-CN')
    expect(messages.value.editorTitle).toBe('代码编辑器')
    expect(messages.value.studioTitle).toBe('主题工作室')
    expect(messages.value.runError).toBe('运行错误：')
    expect(formatMessage('chooseColor', { key: 'primary' })).toBe(
      '选择 primary 颜色',
    )
    expect(formatMessage('updatedAt', { time: '12:34' })).toBe('更新于 12:34')

    language.value = 'zh-TW'
    expect(messages.value.editorTitle).toBe('程式碼編輯器')
    expect(messages.value.studioTitle).toBe('主題工作室')
    expect(messages.value.copy).toBe('複製')
    expect(messages.value.documentPreview).toBe('文件預覽')
    expect(formatMessage('updatedAt', { time: '12:34' })).toBe('更新於 12:34')
  })

  it('localizes catalog labels while preserving their IDs and custom values', () => {
    const language = shallowRef('zh-CN')
    const i18n = createToolI18n(language)
    const theme = { id: 'ocean', name: 'My ocean theme' }

    expect(i18n.exampleLabel('Multi-Section')).toBe('多节文档')
    expect(i18n.exampleLabel('Checklist')).toBe('任务清单')
    expect(i18n.styleTitle('classic', 'Classic')).toBe('经典公文')
    expect(i18n.layoutTitle('modern', 'Business report')).toBe('商务报告')
    expect(i18n.themeTitle(theme.id, theme.name)).toBe('海洋')
    expect(theme).toEqual({ id: 'ocean', name: 'My ocean theme' })
    expect(i18n.themeTitle('custom', 'My theme')).toBe('My theme')

    language.value = 'zh-TW'
    expect(i18n.exampleLabel('Multi-Section')).toBe('多節文件')
    expect(i18n.exampleLabel('Checklist')).toBe('工作清單')
    expect(i18n.styleTitle('academic', 'Academic')).toBe('學術論文')
    expect(i18n.layoutDescription('classic', '')).toContain('首行縮排')
    expect(i18n.exampleLabel('Custom example')).toBe('Custom example')
  })

  it('falls back to English for unsupported document languages', () => {
    const language = shallowRef('fr-FR')
    const { locale, messages, themeTitle } = createToolI18n(language)

    expect(resolveToolLocale('')).toBe('en-US')
    expect(locale.value).toBe('en-US')
    expect(messages.value.downloadWord).toBe('Download Word')
    expect(themeTitle('warm', 'Warm')).toBe('Warm')
  })

  it('keeps message coverage and interpolation fields aligned across languages', () => {
    const language = shallowRef('en-US')
    const { messages } = createToolI18n(language)
    const english = messages.value
    const keys = Object.keys(english).sort()

    for (const locale of ['zh-CN', 'zh-TW']) {
      language.value = locale
      expect(Object.keys(messages.value).sort()).toEqual(keys)
      for (const [key, value] of Object.entries(messages.value)) {
        expect(value.trim()).not.toBe('')
        const original =
          Object.entries(english).find(([name]) => name === key)?.[1] ?? ''
        expect(value.match(/\{\w+\}/g) ?? []).toEqual(
          original.match(/\{\w+\}/g) ?? [],
        )
      }
    }
  })
})
