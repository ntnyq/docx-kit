import { createHighlighterCore } from '@shikijs/core'
import { createJavaScriptRegexEngine } from '@shikijs/engine-javascript'
import typescript from '@shikijs/langs/typescript'
import githubDark from '@shikijs/themes/github-dark'
import githubLight from '@shikijs/themes/github-light'
import type { HighlighterCore } from '@shikijs/core'

export interface HighlightedToken {
  content: string
  htmlStyle?: Record<string, string>
}

let highlighterPromise: Promise<HighlighterCore> | undefined

/**
 * Tokenize theme source once for both interface color modes, without HTML output.
 */
export async function highlightThemeCode(
  code: string,
): Promise<HighlightedToken[]> {
  highlighterPromise ??= createHighlighterCore({
    engine: createJavaScriptRegexEngine(),
    langs: [typescript],
    themes: [githubLight, githubDark],
  }).catch(error => {
    highlighterPromise = undefined
    throw error
  })
  const highlighter = await highlighterPromise
  const { tokens } = highlighter.codeToTokens(code, {
    defaultColor: false,
    lang: 'typescript',
    themes: { dark: 'github-dark', light: 'github-light' },
  })
  return tokens.flatMap((line, index): HighlightedToken[] =>
    index === 0 ? line : [{ content: '\n' }, ...line],
  )
}
