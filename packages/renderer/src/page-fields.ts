type PageField = 'NUMPAGES' | 'PAGE'

const FIELD_CLASS_PREFIX = 'docx-kit-page-field-'

/**
 * Adapt complex PAGE/NUMPAGES fields that docx-preview otherwise discards.
 * Only the parsed preview tree changes; exported DOCX fields stay dynamic.
 */
export function preparePageFields(document: unknown) {
  if (!isRecord(document) || !Array.isArray(document.parts)) {
    return
  }
  for (const part of document.parts) {
    if (!isRecord(part)) {
      continue
    }
    const root = part.body ?? part.rootElement
    if (isRecord(root)) {
      prepareNode(root)
    }
  }
}

/**
 * Resolve fields after pagination, including repeated header/footer instances.
 * Values describe preview pages, whose pagination can differ from Word.
 */
export function resolvePageFields(container: HTMLElement, className: string) {
  const pages = [...container.querySelectorAll(':scope section')].filter(page =>
    page.classList.contains(className),
  )
  for (const [index, page] of pages.entries()) {
    const values: Record<PageField, number> = {
      NUMPAGES: pages.length,
      PAGE: index + 1,
    }
    for (const field of ['PAGE', 'NUMPAGES'] as const) {
      for (const element of page.querySelectorAll(
        `.${FIELD_CLASS_PREFIX}${field}`,
      )) {
        element.textContent = String(values[field])
      }
    }
  }
}

function getChildren(node: Record<string, unknown>): Record<string, unknown>[] {
  return Array.isArray(node.children) ? node.children.filter(isRecord) : []
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function prepareNode(node: Record<string, unknown>) {
  const children = getChildren(node)
  for (let index = 0; index < children.length; index++) {
    const run = children[index]
    if (
      !getChildren(run).some(
        child => child.type === 'complexField' && child.charType === 'begin',
      )
    ) {
      prepareNode(run)
      continue
    }

    let depth = 0
    let instruction = ''
    let end = index
    for (; end < children.length; end++) {
      for (const child of getChildren(children[end])) {
        if (child.type === 'complexField') {
          if (child.charType === 'begin') {
            depth++
          }
          if (child.charType === 'end') {
            depth--
          }
        } else if (child.type === 'instruction' && depth === 1) {
          instruction += child.text
        }
      }
      if (depth === 0) {
        break
      }
    }
    const field = /^\s*(PAGE|NUMPAGES)\s*(?:\\\*\s+MERGEFORMAT\s*)?$/i
      .exec(instruction)?.[1]
      ?.toUpperCase()
    if (!field || end === children.length) {
      continue
    }

    // Retain the instruction run's font/style and discard any stale cached value.
    const fieldRun =
      children
        .slice(index, end + 1)
        .find(child =>
          getChildren(child).some(item => item.type === 'instruction'),
        ) ?? run
    for (let cursor = index; cursor <= end; cursor++) {
      children[cursor].children = []
    }
    fieldRun.fieldRun = false
    fieldRun.className =
      `${typeof fieldRun.className === 'string' ? fieldRun.className : ''} ${FIELD_CLASS_PREFIX}${field}`.trim()
    fieldRun.children = [{ text: '', type: 'text' }]
    index = end
  }
}
