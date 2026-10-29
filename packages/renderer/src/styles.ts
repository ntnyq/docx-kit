import { generate, ident, parse, walk } from 'css-tree'
import type { CssNode, Declaration, List } from 'css-tree'

let scopeId = 0

/**
 * Keep parsed document CSS, embedded font names and positioned content inside
 * one render. The wrapper preserves light-DOM access for preview consumers.
 */
export function isolatePreviewStyles(
  nodes: Node[],
  document: Document,
): HTMLElement {
  const scope = `docx-kit-render-${++scopeId}`
  const selector = `[data-docx-kit-render="${scope}"]`
  const wrapper = document.createElement('div')
  wrapper.dataset.docxKitRender = scope
  // Containment also bounds fixed/absolute positioning supplied by document CSS.
  wrapper.style.setProperty('contain', 'layout paint style', 'important')
  wrapper.style.setProperty('isolation', 'isolate', 'important')
  wrapper.append(...nodes)

  const styles = [...wrapper.querySelectorAll('style')].map(element => {
    try {
      return {
        element,
        ast: parse(element.textContent, {
          parseCustomProperty: true,
          onParseError(error) {
            throw error
          },
        }),
      }
    } catch {
      // Never fall back to mounting CSS that could not be fully parsed.
      element.textContent = ''
      return {
        ast: undefined,
        element,
      }
    }
  })
  const fonts = new Map<string, string>()
  for (const { ast } of styles) {
    if (!ast) {
      continue
    }
    walk(ast, node => {
      if (node.type !== 'Atrule' || node.name.toLowerCase() !== 'font-face') {
        return
      }
      node.block?.children.forEach(child => {
        if (
          child.type !== 'Declaration'
          || ident.decode(child.property).toLowerCase() !== 'font-family'
        ) {
          return
        }
        const name = fontFamilyKey(child.value)
        if (name && !fonts.has(name)) {
          fonts.set(name, `${scope}-font-${fonts.size}`)
        }
      })
    })
  }

  for (const { ast, element } of styles) {
    if (ast?.type !== 'StyleSheet') {
      continue
    }
    scopeRules(ast.children, selector, fonts)
    element.textContent = generate(ast)
  }
  for (const element of wrapper.querySelectorAll<HTMLElement>('[style]')) {
    try {
      const ast = parse(element.getAttribute('style') ?? '', {
        context: 'declarationList',
        parseCustomProperty: true,
        onParseError(error) {
          throw error
        },
      })
      if (ast.type === 'DeclarationList') {
        sanitizeDeclarations(ast.children, fonts)
        element.setAttribute('style', generate(ast))
      }
    } catch {
      element.removeAttribute('style')
    }
  }
  return wrapper
}

function fontFamilyKey(value: CssNode): string | undefined {
  return value.type === 'Value'
    ? value.children.toArray().map(fontNamePart).join(' ').toLowerCase()
    : undefined
}

function fontNamePart(node: CssNode): string {
  return node.type === 'String'
    ? node.value
    : node.type === 'Identifier'
      ? ident.decode(node.name)
      : generate(node)
}

function hasUnsafeValue(declaration: Declaration): boolean {
  let unsafe = false
  walk(declaration.value, node => {
    if (node.type === 'Raw') {
      unsafe = true
    }
    if (node.type === 'Url' && !/^(?:blob:|data:|#)/i.test(node.value)) {
      unsafe = true
    }
  })
  return unsafe
}

function sanitizeDeclarations(
  children: List<CssNode>,
  fonts: Map<string, string>,
) {
  children.forEach((node, item, list) => {
    if (node.type !== 'Declaration' || hasUnsafeValue(node)) {
      list.remove(item)
      return
    }
    if (
      ident.decode(node.property).toLowerCase() === 'font-family'
      || /^--docx-(?:major|minor)HAnsi-font$/.test(ident.decode(node.property))
    ) {
      if (node.value.type !== 'Value') {
        return
      }
      // A family is a quoted string or a sequence of identifiers; commas split
      // families, while spaces in a quoted name stay part of that name.
      const families: CssNode[][] = [[]]
      node.value.children.forEach(child => {
        if (child.type === 'Operator' && child.value === ',') {
          families.push([])
        } else {
          families.at(-1)?.push(child)
        }
      })
      const value = parse(
        families
          .map(family => {
            const key = family.map(fontNamePart).join(' ').toLowerCase()
            const scopedName = fonts.get(key)
            return scopedName
              ? JSON.stringify(scopedName)
              : family.map(child => generate(child)).join(' ')
          })
          .join(','),
        { context: 'value' },
      )
      if (value.type === 'Value') {
        node.value = value
      } else {
        list.remove(item)
      }
    }
  })
}

function scopeRules(
  children: List<CssNode>,
  scope: string,
  fonts: Map<string, string>,
) {
  children.forEach((node, item, list) => {
    if (node.type === 'Rule' && node.prelude.type === 'SelectorList') {
      const selectors = node.prelude.children
        .map(selector => {
          const text = generate(selector)
          return text === ':root' ? scope : `${scope} ${text}`
        })
        .toArray()
      const prelude = parse(selectors.join(','), { context: 'selectorList' })
      if (prelude.type !== 'SelectorList') {
        list.remove(item)
        return
      }
      node.prelude = prelude
      sanitizeDeclarations(node.block.children, fonts)
      return
    }
    if (node.type === 'Atrule' && node.block) {
      const name = node.name.toLowerCase()
      if (name === 'media' && node.prelude?.type === 'AtrulePrelude') {
        scopeRules(node.block.children, scope, fonts)
        return
      }
      if (name === 'font-face') {
        sanitizeDeclarations(node.block.children, fonts)
        return
      }
    }
    // Imports, global registries, nested CSS and unparsed fragments are not
    // emitted by the adapter and must not gain host-document privileges.
    list.remove(item)
  })
}
