/**
 * Compile-time type tests for theme token support in style rules.
 *
 * The runtime side (resolveThemeTokens) is covered by tests in
 * tests/theme/theme.test.ts. This file focuses on whether the
 * TypeScript types accept `$category.key` references in fields
 * that should support theme tokens.
 *
 * If a test below fails to typecheck (e.g. an assignment that
 * should be allowed), it means the underlying type is too narrow
 * and the user-facing type error reported in the playground has
 * regressed.
 *
 * The runtime `it()` calls below are minimal smoke checks; their
 * real job is to make the file valid vitest output.
 */

import { describe, expect, it } from 'vitest'
import type {
  DocxStyleRule,
  StyleToken,
  ThemeToken,
  UnitValue,
} from '@docxkit/types'

describe('theme token types', () => {
  it('UnitValue accepts theme tokens', () => {
    const values: UnitValue[] = [
      '$spacing.lg',
      '$fontSize.base',
      '$spacing.md',
      12,
      '12pt',
      `$spacing.${'lg'}`,
    ]
    expect(values).toHaveLength(6)
  })

  it('StyleToken generic works for string and number bases', () => {
    const color: StyleToken<string> = '$colors.primary'
    const literal: StyleToken<string> = '#1a56db'
    const size: StyleToken<number> = '$fontSize.lg'
    const num: StyleToken<number> = 16
    expect([color, literal, size, num]).toHaveLength(4)
  })

  it('ThemeToken template literal type accepts all categories', () => {
    const tokens: ThemeToken[] = [
      '$colors.primary',
      '$fonts.heading',
      '$spacing.xl',
      '$fontSize.base',
    ]
    expect(tokens).toHaveLength(4)
  })

  it('DocxStyleRule accepts theme tokens in token-aware fields', () => {
    // The exact pattern from the user's playground report:
    // `subtitle: { fontSize: '$fontSize.lg', color: '$colors.muted', marginBottom: '$spacing.xl' }`
    const styleRule: DocxStyleRule = {
      backgroundColor: '$colors.info',
      border: { color: '$colors.primary', width: '$spacing.xs' },
      borderBottom: { color: '$colors.muted' },
      borderTop: { color: '$colors.primary' },
      color: '$colors.muted',
      fontFamily: '$fonts.heading',
      fontSize: '$fontSize.lg',
      height: '$spacing.half',
      letterSpacing: '$spacing.xs',
      lineHeight: 1.5,
      margin: '$spacing.lg',
      marginBottom: '$spacing.xl',
      marginLeft: '$spacing.sm',
      marginRight: '$spacing.lg',
      marginTop: '$spacing.md',
      textAlign: 'center',
      textIndent: '$spacing.xs',
      width: '$spacing.full',
    }
    expect(styleRule).toBeDefined()
  })

  it('DocxStyleRule accepts plain values mixed with tokens', () => {
    const rule: DocxStyleRule = {
      backgroundColor: '$colors.info',
      color: '#000000',
      fontFamily: 'Arial',
      fontSize: 12,
      marginBottom: '$spacing.lg',
      marginTop: '24pt',
    }
    expect(rule).toBeDefined()
  })

  it('DocxStyleRule accepts HexColor literals', () => {
    const rule: DocxStyleRule = {
      backgroundColor: '#00ff00',
      color: '#ff0000',
    }
    expect(rule).toBeDefined()
  })
})
