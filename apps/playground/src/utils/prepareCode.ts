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
  // Parse emitted JavaScript once; applying edits from right to left keeps all
  // source ranges valid while imports and the final expression are rewritten.
  const program = parse(body, {
    allowAwaitOutsideFunction: true,
    allowReturnOutsideFunction: true,
    ecmaVersion: 'latest',
    sourceType: 'module',
  })
  let hasFollowingStatement = false
  for (const statement of program.body.toReversed()) {
    if (statement.type !== 'ImportDeclaration') {
      if (!hasFollowingStatement && statement.type === 'ExpressionStatement') {
        body = `${body.slice(0, statement.start)}return (${body.slice(statement.expression.start, statement.expression.end)});${body.slice(statement.end)}`
      }
      hasFollowingStatement = true
      continue
    }
    const namespace =
      statement.source.value === 'docx-kit'
        ? 'arguments[0]'
        : statement.source.value === 'docx'
          ? 'arguments[1]'
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
    // The async arrow inherits the executor's arguments. Unlike a named
    // namespace variable, these references cannot be shadowed by import aliases.
    const bindings = namespace
      ? [
          ...(names.length
            ? [`const { ${names.join(', ')} } = ${namespace};`]
            : []),
          ...statement.specifiers.flatMap(specifier =>
            specifier.type === 'ImportNamespaceSpecifier'
              ? [`const ${specifier.local.name} = ${namespace};`]
              : [],
          ),
        ]
      : []
    const replacement = bindings.join('\n')
    // A rewritten import with bindings is a declaration in the emitted code;
    // erased imports do not hide an earlier final expression.
    hasFollowingStatement ||= replacement.length > 0
    body = `${body.slice(0, statement.start)}${replacement}${body.slice(statement.end)}`
  }
  // The async arrow inherits the executor's arguments for runtime bindings.
  return `"use strict";
return (async () => {
${body}
})()`
}
