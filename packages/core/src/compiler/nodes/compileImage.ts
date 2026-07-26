/**
 * Compile an image node into a `docx` Paragraph containing an `ImageRun`.
 *
 * @module compiler/nodes/compileImage
 */

import { DocxKitError } from '@docxkit/types'
import { Paragraph, TextWrappingSide, TextWrappingType } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { dataUrlToUint8Array } from '../../utils/dataUrl'
import {
  createImageRun,
  readImageMetadata,
  resolveImageDimensions,
} from '../../utils/image'
import { compileParagraphStyle } from '../compileStyle'
import { toPx } from '../units'
import type { Buffer } from 'node:buffer'
import type { DocxKitConfig, ImageNode, StyleSheet } from '@docxkit/types'
import type { IFloating } from 'docx'

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
  const metadata = readImageMetadata(data, node.imageType)
  if (!metadata) {
    throw new DocxKitError(
      'IMAGE_INVALID_DATA',
      'Image format could not be detected; provide a supported imageType',
    )
  }
  const dimensions = resolveImageDimensions(
    toPx(node.width),
    toPx(node.height),
    metadata,
    { height: 180, width: 300 },
  )

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
        alt: node.alt,
        data,
        floating: compileImageFloating(node.floating),
        height: dimensions.height,
        type: metadata.type,
        width: dimensions.width,
      }),
    ],
  })
}

export function compileImageFloating(
  floating: ImageNode['floating'],
): IFloating | undefined {
  if (!floating) {
    return undefined
  }
  if (floating === true) {
    return {
      horizontalPosition: { offset: 0 },
      verticalPosition: { offset: 0 },
    }
  }
  return {
    horizontalPosition: { offset: toEmu(floating.x) ?? 0 },
    verticalPosition: { offset: toEmu(floating.y) ?? 0 },
    wrap: floating.wrap
      ? {
          side: TextWrappingSide.BOTH_SIDES,
          type: {
            square: TextWrappingType.SQUARE,
            tight: TextWrappingType.TIGHT,
            topAndBottom: TextWrappingType.TOP_AND_BOTTOM,
          }[floating.wrap],
        }
      : undefined,
  }
}

export async function normalizeImageData(
  data: unknown,
): Promise<string | ArrayBuffer | Buffer | Uint8Array> {
  if (typeof Blob !== 'undefined' && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer())
  }
  if (typeof data === 'string' && data.startsWith('data:')) {
    return dataUrlToUint8Array(data)
  }
  return data as string | ArrayBuffer | Buffer | Uint8Array
}

function toEmu(value: ImageNode['width']): number | undefined {
  const pixels = toPx(value)
  return pixels === undefined ? undefined : Math.round(pixels * 9525)
}

export { dataUrlToUint8Array }
