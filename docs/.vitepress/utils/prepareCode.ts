import { transform } from 'sucrase'

/**
 * Transform user-written TypeScript code into executable JavaScript.
 *
 * Pipeline:
 * 1. `import { X } from 'docx'` → `const { X } = docx` (injected namespace)
 * 2. `import { X } from 'docx-kit'` → stripped (injected as individual args)
 * 3. Remaining `import` lines → stripped
 * 4. TypeScript→JavaScript via sucrase (strips generics, type annotations, etc.)
 * 5. Last expression auto-wrapped with `return` if needed
 * 6. Wrapped in `"use strict"; return (async () => { … })()`
 */
export function prepareCode(raw: string): string {
  // Step 1: Transform docx imports → destructuring from injected 'docx' namespace.
  //    import { Paragraph, BorderStyle } from 'docx'
  //    →  const { Paragraph, BorderStyle } = docx
  let body = raw.replace(
    /^import\s+\{([^}]+)\}\s+from\s+['"]docx['"]\s*;?/gm,
    (_match, names: string) =>
      `const { ${names
        .split(',')
        .map(s => s.trim())
        .join(', ')} } = docx`,
  )

  // Step 2: Strip docx-kit imports (they're injected as individual args).
  body = body.replace(
    /^import\s+\{[^}]+\}\s+from\s+['"]docx-kit['"]\s*;?/gm,
    '',
  )

  // Step 3: Strip any remaining import / import type statements.
  body = body.replace(/^import\s+(?:\S.*)?$/gm, '')

  // Step 4: Transpile TypeScript → JavaScript via sucrase.
  //    Strips generics, type annotations, enum declarations, etc.
  body = transform(body, {
    transforms: ['typescript'],
  }).code.trim()

  // Step 5: Find the last non-empty line and prepend 'return ' if it's an expression.
  const lines = body.split('\n')
  let lastIdx = lines.length - 1
  while (lastIdx >= 0 && lines[lastIdx].trim() === '') {
    lastIdx--
  }

  if (lastIdx >= 0) {
    const trimmed = lines[lastIdx].trim()
    const isDeclaration =
      /^(?:const|let|var|if|for|while|function|class|import|export|return|throw)\b/
    const isBlockEnd = /^[})]/
    const isComment = /^\/\//

    if (
      !isDeclaration.test(trimmed)
      && !isBlockEnd.test(trimmed)
      && !isComment.test(trimmed)
      && trimmed !== ''
    ) {
      const indent = lines[lastIdx].match(/^(\s*)/)?.[1] ?? ''
      lines[lastIdx] = `${indent}return ${trimmed}`
    }
  }

  const source = lines.join('\n')

  // Step 6: Wrap in async IIFE so user `await` works.
  return `"use strict";
return (async () => {
${source}
})()`
}
