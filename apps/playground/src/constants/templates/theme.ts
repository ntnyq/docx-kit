import { unindent } from '@ntnyq/utils'

/**
 * Theme system playground preset.
 *
 * Demonstrates:
 * - useTheme() to load built-in themes
 * - $category.key token syntax in defineStyles()
 * - theme config property in createDocx()
 */
export const THEME_CODE = unindent(`
  import { createDocx, defineStyles, useTheme } from 'docx-kit'

  // 1. Pick a built-in theme (try: 'minimal', 'ocean', 'warm')
  const theme = useTheme('ocean')

  // 2. Define styles with theme token references
  //    Tokens use $category.key syntax and resolve at compile time.
  const styles = defineStyles({
    title: {
      fontSize: 28,
      fontWeight: 'bold',
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      textAlign: 'center',
      marginBottom: '$spacing.md',
    },
    subtitle: {
      fontSize: '$fontSize.lg',
      color: '$colors.muted',
      textAlign: 'center',
      marginBottom: '$spacing.xl',
    },
    body: {
      fontSize: '$fontSize.base',
      color: '$colors.text',
      fontFamily: '$fonts.body',
      lineHeight: 1.6,
    },
    callout: {
      backgroundColor: '$colors.info',
      color: '#ffffff',
      padding: '$spacing.sm',
    },
    footer: {
      fontSize: '$fontSize.sm',
      color: '$colors.muted',
      textAlign: 'center',
      marginTop: '$spacing.xl',
    },
  })

  // 3. Build with theme
  const doc = createDocx({
    styles,
    theme,
    page: { size: 'A4', margin: '20mm 25mm' },
  })

  doc
    .h1('Theme System Demo', { className: 'title' })
    .p('This document uses the ocean theme with semantic design tokens.', {
      className: 'subtitle',
    })
    .p(
      'Theme tokens are referenced in styles using the $category.key syntax. ' +
      'At compile time, $colors.primary is replaced with the actual color ' +
      'from the theme (e.g. #0ea5e9 for the ocean theme).',
      { className: 'body' },
    )
    .p(
      'Callout: switch themes by changing useTheme("warm") or useTheme("minimal").',
      { className: 'callout' },
    )
    .p('Footer text uses $fontSize.sm and $colors.muted from the theme.', {
      className: 'footer',
    })

  doc.toBlob()
`)
