/**
 * Compile an image node into a `docx` Paragraph containing an `ImageRun`.
 *
 * @module compiler/nodes/compileImage
 */

import { DocxKitError } from '@docxkit/types'
import { Paragraph } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { dataUrlToUint8Array } from '../../utils/dataUrl'
import { createImageRun } from '../../utils/image'
import { compileParagraphStyle } from '../compileStyle'
import { toPx } from '../units'
import type { Buffer } from 'node:buffer'
import type { DocxKitConfig, ImageNode, StyleSheet } from '@docxkit/types'

type FloatingOptions = Record<string, unknown>

export async function compileImage<TStyles extends StyleSheet>(
  node: ImageNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  if (
    node.data == null
    || (typeof node.data === 'string' && node.data.length === 0)
  ) {
    throw new DocxKitError('IMAGE_INVALID_DATA', 'Image data is empty or null')
  }

  const data = await normalizeImageData(node.data)
  const imageType = node.imageType ?? 'png'

  const style = resolveStyle({
    base: config.defaults?.image,
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return new Paragraph({
    ...compileParagraphStyle(style),
    children: [
      createImageRun({
        data,
        floating: compileFloating(node.floating),
        height: toPx(node.height),
        type: imageType,
        width: toPx(node.width),
      }),
    ],
  })
}

export async function normalizeImageData(
  data: unknown,
): Promise<string | ArrayBuffer | Buffer | Uint8Array> {
  if (typeof Blob !== 'undefined' && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer())
  }
  return data as string | ArrayBuffer | Buffer | Uint8Array
}

function compileFloating(
  floating: ImageNode['floating'],
): FloatingOptions | undefined {
  if (!floating) {
    return undefined
  }
  if (floating === true) {
    return {}
  }
  return {
    horizontalPosition:
      floating.x === undefined ? undefined : { offset: floating.x },
    verticalPosition:
      floating.y === undefined ? undefined : { offset: floating.y },
  }
}

export { dataUrlToUint8Array }
