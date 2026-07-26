import { Packer } from 'docx'
import { describe, expect, it, vi } from 'vitest'
import { saveDocument, streamDocument } from '../src/node/fs'
import type { Document } from 'docx'

const { createWriteStream } = vi.hoisted(() => ({
  createWriteStream: vi.fn(),
}))

vi.mock('node:fs', async () => {
  const { Writable } = await import('node:stream')
  createWriteStream.mockImplementation(
    () =>
      new Writable({
        write(_chunk, _encoding, callback) {
          callback()
        },
      }),
  )
  return { createWriteStream }
})

// Mock docx Packer
vi.mock('docx', async () => {
  const actual = await vi.importActual('docx')
  const { Readable } = await import('node:stream')
  return {
    ...(actual as object),
    Packer: {
      toStream: vi.fn(() => Readable.from([new Uint8Array([80, 75, 3, 4])])),
    },
  }
})

// eslint-disable-next-line @typescript-eslint/consistent-type-assertions
const mockDoc = {} as Document

describe('saveDocument (node)', () => {
  it('exposes the packer stream', () => {
    const stream = streamDocument(mockDoc)

    expect(Packer.toStream).toHaveBeenCalledWith(mockDoc)
    expect(stream).toBeDefined()
  })

  it('streams the package to a file', async () => {
    await saveDocument(mockDoc, 'test.docx')
    expect(Packer.toStream).toHaveBeenCalledWith(mockDoc)
    expect(createWriteStream).toHaveBeenCalledWith('test.docx')
  })
})
