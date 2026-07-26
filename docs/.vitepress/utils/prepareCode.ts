import { transform } from 'sucrase'

/**
 * Transform user-written TypeScript code into executable JavaScript.
 *
 * Pipeline:
 * 1. Named `docx` imports → destructuring from the injected `docx` namespace
 * 2. Named `docx-kit` imports → destructuring from the injected `docxKit` namespace
 * 3. Remaining `import` lines → stripped
 * 4. TypeScript→JavaScript via sucrase (strips generics, type annotations, etc.)
 * 5. Last expression auto-wrapped with `return` if needed
 * 6. Wrapped in `"use strict"; return (async () => { … })()`
 */
export function prepareCode(raw: string): string {
  let body = replaceNamedImports(raw, 'docx', 'docx')

  body = replaceNamedImports(body, 'docx-kit', 'docxKit')

  // Step 3: Strip unsupported imports after the two runtime namespaces have
  // been mapped. The first pattern handles multiline named/type imports.
  body = body.replace(
    /^[\t ]*import\s+(?:type\s+)?\{[^}]*\}\s+from\s+['"][^'"]+['"][\t ]*;?/gm,
    '',
  )
  body = body.replace(/^[\t ]*import[^\n]*$/gm, '')

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

/**
 * Replace named ESM imports with destructuring from an injected namespace.
 *
 * Type-only specifiers are omitted, and aliases are translated from
 * `source as local` to object-destructuring syntax (`source: local`).
 */
function replaceNamedImports(
  source: string,
  moduleName: string,
  namespaceName: string,
): string {
  const escapedModuleName = moduleName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const importPattern = new RegExp(
    `^[\\t ]*import\\s+\\{([^}]+)\\}\\s+from\\s+['"]${escapedModuleName}['"][\\t ]*;?`,
    'gm',
  )

  return source.replace(importPattern, (_match, names: string) => {
    const runtimeNames = names
      .split(',')
      .map(name => name.trim())
      .filter(name => name.length > 0 && !name.startsWith('type '))
      .map(name => name.replace(/\s+as\s+/, ': '))

    if (runtimeNames.length === 0) {
      return ''
    }

    return `const { ${runtimeNames.join(', ')} } = ${namespaceName};`
  })
}
