import { describe, expect, it } from 'vitest'
import {
  createImageRun,
  readImageMetadata,
  resolveImageDimensions,
} from '../../src/utils/image'

function createPngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24)
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const view = new DataView(bytes.buffer)
  view.setUint32(16, width)
  view.setUint32(20, height)
  return bytes
}

describe('image metadata', () => {
  it('detects PNG type and intrinsic dimensions', () => {
    expect(readImageMetadata(createPngHeader(400, 200))).toEqual({
      height: 200,
      type: 'png',
      width: 400,
    })
  })

  it('normalizes JPEG type hints for docx', () => {
    expect(
      readImageMetadata(new Uint8Array([0xff, 0xd8, 0xff]), 'jpeg'),
    ).toMatchObject({ type: 'jpg' })
  })

  it('preserves the intrinsic aspect ratio when one size is omitted', () => {
    const dimensions = resolveImageDimensions(
      100,
      undefined,
      { height: 200, type: 'png', width: 400 },
      { height: 180, width: 300 },
    )

    expect(dimensions).toEqual({ height: 50, width: 100 })
  })
})

describe('createImageRun', () => {
  it('requires a raster fallback for SVG images', () => {
    expect(() =>
      createImageRun({
        data: new Uint8Array([60, 115, 118, 103, 47, 62]),
        type: 'svg',
      }),
    ).toThrow('raster fallback')
  })

  it('accepts SVG images with a raster fallback', () => {
    expect(() =>
      createImageRun({
        data: new Uint8Array([60, 115, 118, 103, 47, 62]),
        type: 'svg',
        fallback: {
          data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
          type: 'png',
        },
      }),
    ).not.toThrow()
  })
})
