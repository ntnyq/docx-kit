import { resolveStyle } from '../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from './compileStyle'
import type { DocxKitConfig, DocxStyleRule } from '@docxkit/types'
import type { IStylesOptions } from 'docx'

/**
 * Write native Word defaults so plugin output and newly typed text share the
 * document typography. Body indentation stays on body paragraphs, never on
 * table cells, captions, or other plugin paragraphs.
 */
export function compileDocumentStyles(config: DocxKitConfig): IStylesOptions {
  const theme = config.theme
  const body = resolveStyle({
    theme,
    base: {
      color: theme?.colors?.text ?? '#252525',
      fontFamily: theme?.fonts?.body ?? 'Calibri',
      fontFamilyEastAsia: theme?.fonts?.bodyEastAsia ?? 'Microsoft YaHei',
      fontSize: theme?.fontSize?.base ?? 11,
      ...config.defaults?.text,
      ...config.defaults?.paragraph,
    },
  })
  const heading = (level: number) => {
    const key = `h${level}`
    const style = resolveStyle({
      className: config.styles?.[key] ? key : undefined,
      styles: config.styles,
      theme,
      base: {
        color: theme?.colors?.primary ?? '#202020',
        fontFamily: theme?.fonts?.heading ?? body.fontFamily,
        fontWeight: 'bold',
        keepLines: true,
        keepNext: true,
        lineHeight: 1.2,
        marginBottom: theme?.spacing?.sm ?? 8,
        marginTop: level === 1 ? 0 : (theme?.spacing?.md ?? 16),
        textIndent: 0,
        fontFamilyEastAsia:
          theme?.fonts?.headingEastAsia ?? body.fontFamilyEastAsia,
        fontSize:
          level === 1
            ? (theme?.fontSize?.xl ?? 24)
            : level === 2
              ? (theme?.fontSize?.lg ?? 16)
              : level === 3
                ? 13
                : 11,
        ...config.defaults?.text,
      },
    })
    return {
      paragraph: compileParagraphStyle(style),
      run: compileTextStyle(style),
    }
  }
  const title: DocxStyleRule = {
    ...body,
    color: '#000000',
    fontFamily: theme?.fonts?.heading ?? body.fontFamily,
    fontSize: 28,
    fontWeight: 'bold',
    keepNext: true,
    lineHeight: 1.15,
    marginBottom: 12,
    textIndent: 0,
    fontFamilyEastAsia:
      theme?.fonts?.headingEastAsia ?? body.fontFamilyEastAsia,
  }
  return {
    default: {
      heading1: heading(1),
      heading2: heading(2),
      heading3: heading(3),
      heading4: heading(4),
      heading5: heading(5),
      heading6: heading(6),
      document: {
        run: compileTextStyle(body),
        paragraph: compileParagraphStyle({
          lineHeight: 1.35,
          marginBottom: 0,
          widowControl: true,
        }),
      },
      title: {
        paragraph: compileParagraphStyle(title),
        run: compileTextStyle(title),
      },
    },
  }
}
