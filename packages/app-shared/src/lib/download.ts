/**
 * Save generated content locally without uploading the document.
 */
export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  // Keep the URL alive until the browser has started consuming the download.
  setTimeout(() => URL.revokeObjectURL(url), 30_000)
}
