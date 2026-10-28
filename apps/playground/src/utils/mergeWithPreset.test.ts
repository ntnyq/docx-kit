import { describe, expect, it } from 'vitest'
import { mergeWithPreset } from './mergeWithPreset'

describe('playground style preset application', () => {
  it('keeps user choices while filling missing defaults from the selected preset', () => {
    const preset = {
      styles: { h1: { fontSize: 28 }, h2: { fontSize: 18 } },
      defaults: {
        cell: { fontSize: 10 },
        paragraph: { fontSize: 12, lineHeight: 1.5 },
      },
    }
    const user = {
      defaults: { paragraph: { fontSize: 16 } },
      styles: { h1: { fontSize: 32 } },
    }
    expect(mergeWithPreset(preset, user)).toMatchObject({
      styles: { h1: { fontSize: 32 }, h2: { fontSize: 18 } },
      defaults: {
        cell: { fontSize: 10 },
        paragraph: { fontSize: 16, lineHeight: 1.5 },
      },
    })
    expect(preset.defaults.paragraph.fontSize).toBe(12)
    expect(user.defaults.paragraph).toEqual({ fontSize: 16 })
  })
})
