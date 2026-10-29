import { parse } from 'acorn'
import { transform } from 'sucrase'

/**
 * Transform user-written TypeScript code into executable JavaScript.
 *
 * Pipeline:
 * 1. Strip TypeScript syntax while retaining runtime imports
 * 2. Rewrite import declarations using their parsed source ranges
 * 3. Last expression auto-wrapped with `return` if needed
 * 4. Wrapped in `"use strict"; return (async () => { … })()`
 */
export function prepareCode(raw: string): string {
  let body = transform(raw, {
    keepUnusedImports: true,
    transforms: ['typescript'],
  }).code.trim()
  body = replaceImports(body)

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

  // Wrap in async IIFE so user `await` works.
  return `"use strict";
return (async () => {
${body}
})()`
}

/**
 * Replace named ESM imports with destructuring from an injected namespace.
 *
 * TypeScript has already been removed. Only actual import declarations are
 * rewritten; strings, comments and dynamic imports retain their content.
 */
function replaceImports(source: string): string {
  const program = parse(source, {
    allowAwaitOutsideFunction: true,
    allowReturnOutsideFunction: true,
    ecmaVersion: 'latest',
    sourceType: 'module',
  })
  for (const statement of program.body.toReversed()) {
    if (statement.type !== 'ImportDeclaration') {
      continue
    }
    const namespace =
      statement.source.value === 'docx-kit'
        ? 'docxKit'
        : statement.source.value === 'docx'
          ? 'docx'
          : undefined
    const names = statement.specifiers.flatMap(specifier => {
      if (specifier.type !== 'ImportSpecifier') {
        return []
      }
      const imported =
        specifier.imported.type === 'Identifier'
          ? specifier.imported.name
          : JSON.stringify(specifier.imported.value)
      return [
        imported === specifier.local.name
          ? imported
          : `${imported}: ${specifier.local.name}`,
      ]
    })
    const replacement =
      namespace && names.length
        ? `const { ${names.join(', ')} } = ${namespace};`
        : ''
    source = `${source.slice(0, statement.start)}${replacement}${source.slice(statement.end)}`
  }
  return source
}
