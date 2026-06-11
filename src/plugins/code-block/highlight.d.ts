declare module 'highlight.js' {
  interface HighlightAPI {
    getLanguage(name: string): unknown
    highlight(code: string, options: { language: string }): HighlightResult
  }

  interface HighlightResult {
    value: string
    language?: string
  }

  const hljs: HighlightAPI
  export default hljs
}
