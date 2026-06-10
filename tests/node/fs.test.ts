import { writeFile } from 'node:fs/promises'

import { Packer } from 'docx'
import { describe, expect, it, vi } from 'vitest'
import { saveDocument } from '../../src/node/fs'
import type { Document } from 'docx'

// Mock node:fs/promises
vi.mock('node:fs/promises', () => ({
  writeFile: vi.fn().mockResolvedValue(undefined),
}))

// Mock docx Packer
vi.mock('docx', async () => {
  const actual = await vi.importActual('docx')
  return {
    ...(actual as object),
    Packer: {
      toBuffer: vi.fn().mockResolvedValue(new Uint8Array([80, 75, 3, 4])),
    },
  }
})

// eslint-disable-next-line @typescript-eslint/consistent-type-assertions
const mockDoc = {} as Document

describe('saveDocument (node)', () => {
  it('packs and writes to file', async () => {
    await saveDocument(mockDoc, 'test.docx')
    expect(Packer.toBuffer).toHaveBeenCalledWith(mockDoc)
    expect(writeFile).toHaveBeenCalledWith('test.docx', expect.any(Uint8Array))
  })
})
