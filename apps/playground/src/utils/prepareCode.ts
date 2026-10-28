import { parse } from 'acorn'
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

  // Return the complete final expression, regardless of line breaks or comments.
  const program = parse(body, {
    allowAwaitOutsideFunction: true,
    allowReturnOutsideFunction: true,
    ecmaVersion: 'latest',
  })
  const last = program.body.at(-1)
  if (last?.type === 'ExpressionStatement') {
    body = `${body.slice(0, last.start)}return (${body.slice(last.expression.start, last.expression.end)});${body.slice(last.end)}`
  }

  // Step 6: Wrap in async IIFE so user `await` works.
  return `"use strict";
return (async () => {
${body}
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
