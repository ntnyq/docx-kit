/**
 * Enforce navigation and iframe isolation before preview nodes are mounted.
 */
export function sanitizePreviewNode(node: Node): Node {
  if (node.nodeType !== 1) {
    return node
  }
  const element = node as Element
  if (element.localName === 'iframe') {
    // No allow-scripts or allow-same-origin: opt-in HTML stays inert and isolated.
    element.setAttribute('sandbox', '')
    element.setAttribute('referrerpolicy', 'no-referrer')
  }
  if (element.localName === 'a') {
    const href = element.getAttribute('href')
    if (href && !isSafeHyperlink(href)) {
      element.removeAttribute('href')
    }
    element.setAttribute('rel', 'noopener noreferrer')
  }
  return node
}

function isSafeHyperlink(href: string): boolean {
  try {
    const { protocol } = new URL(href, 'https://docx-kit.invalid/')
    return ['http:', 'https:', 'mailto:', 'tel:'].includes(protocol)
  } catch {
    return false
  }
}
