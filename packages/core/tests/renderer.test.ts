import { Packer } from 'docx'
import { describe, expect, it, vi } from 'vitest'
import {
  packToBase64String,
  packToBlob,
  packToBuffer,
} from '../src/renderer/pack'
import type { Document } from 'docx'

// Mock the docx Packer module
vi.mock('docx', async () => {
  const actual = await vi.importActual('docx')
  return {
    ...(actual as object),
    Packer: {
      toBase64String: vi.fn().mockResolvedValue('base64data'),
      toBlob: vi.fn().mockResolvedValue(new Blob(['test'])),
      toArrayBuffer: vi
        .fn()
        .mockResolvedValue(new Uint8Array([80, 75, 3, 4]).buffer),
    },
  }
})

// eslint-disable-next-line @typescript-eslint/consistent-type-assertions
const mockDoc = {} as Document

describe('packToBase64String', () => {
  it('calls Packer.toBase64String', async () => {
    const result = await packToBase64String(mockDoc)
    expect(Packer.toBase64String).toHaveBeenCalledWith(mockDoc)
    expect(result).toBe('base64data')
  })
})

describe('packToBlob', () => {
  it('calls Packer.toBlob', async () => {
    const result = await packToBlob(mockDoc)
    expect(Packer.toBlob).toHaveBeenCalledWith(mockDoc)
    expect(result).toBeInstanceOf(Blob)
  })
})

describe('packToBuffer', () => {
  it('calls Packer.toArrayBuffer and returns Uint8Array', async () => {
    const result = await packToBuffer(mockDoc)
    expect(Packer.toArrayBuffer).toHaveBeenCalledWith(mockDoc)
    expect(result).toBeInstanceOf(Uint8Array)
  })
})
