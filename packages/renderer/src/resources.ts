/**
 * Per-render media ownership for the pinned docx-preview adapter.
 */
interface PreviewDocument {
  blobToURL: (
    blob: Blob | null,
    path?: string,
  ) => string | Promise<string | null> | null
}

/**
 * Track URLs created by this parsed document, including late or failed renders.
 * Never patch the browser's global URL methods or revoke another preview's media.
 */
export function trackDocumentResources(document: unknown): () => void {
  if (!isPreviewDocument(document)) {
    throw new Error('Unsupported docx-preview media adapter')
  }
  const urls = new Set<string>()
  const createURL = document.blobToURL.bind(document)
  let isDisposed = false

  document.blobToURL = async (blob, path) => {
    const url = await createURL(blob, path)
    if (url?.startsWith('blob:')) {
      if (isDisposed) {
        URL.revokeObjectURL(url)
      } else {
        urls.add(url)
      }
    }
    return url
  }

  return () => {
    isDisposed = true
    for (const url of urls) {
      URL.revokeObjectURL(url)
    }
    urls.clear()
  }
}

function isPreviewDocument(document: unknown): document is PreviewDocument {
  return (
    typeof document === 'object'
    && document !== null
    && 'blobToURL' in document
    && typeof document.blobToURL === 'function'
  )
}
