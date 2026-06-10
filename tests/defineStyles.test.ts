import { describe, expect, it } from 'vitest'
import { defineStyles } from '../src/types/style'

describe('defineStyles', () => {
  it('returns the same object', () => {
    const styles = defineStyles({
      h1: { fontSize: '24pt', fontWeight: 'bold' },
    })
    expect(styles.h1).toBeDefined()
    expect(styles.h1!.fontSize).toBe('24pt')
    expect(styles.h1!.fontWeight).toBe('bold')
  })

  it('supports empty stylesheet', () => {
    const styles = defineStyles({})
    expect(styles).toEqual({})
  })

  it('supports multiple style entries', () => {
    const styles = defineStyles({
      body: { fontSize: '12pt', lineHeight: 1.5 },
      red: { color: '#f00' },
      title: { fontSize: '28pt', fontWeight: 700 },
    })
    expect(Object.keys(styles)).toHaveLength(3)
    expect(styles.body!.fontSize).toBe('12pt')
    expect(styles.red!.color).toBe('#f00')
    expect(styles.title!.fontSize).toBe('28pt')
  })

  it('supports numeric font weights', () => {
    const styles = defineStyles({
      bold: { fontWeight: 700 },
      light: { fontWeight: 300 },
    })
    expect(styles.bold!.fontWeight).toBe(700)
    expect(styles.light!.fontWeight).toBe(300)
  })

  it('supports border rules', () => {
    const styles = defineStyles({
      bordered: {
        border: { color: '#333', style: 'single', width: '1pt' },
      },
    })
    expect(styles.bordered!.border!.style).toBe('single')
  })
})
