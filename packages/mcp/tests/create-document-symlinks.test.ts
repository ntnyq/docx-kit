import {
  lstat,
  mkdtemp,
  readFile,
  rm,
  symlink,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { createDocument } from '../src/tools/createDocx'

vi.mock(import('node:fs'), async importOriginal => {
  const original = await importOriginal()
  return {
    ...original,
    // Zero simulates the absent Windows flag. Keep filesystem operations real.
    constants: { ...original.constants, O_NOFOLLOW: 0 },
  }
})

describe('createDocument without O_NOFOLLOW', () => {
  it.each([
    { location: 'inside', targetExists: true },
    { location: 'outside', targetExists: true },
    { location: 'inside', targetExists: false },
    { location: 'outside', targetExists: false },
  ])(
    'rejects a leaf symlink $location the output directory (target exists: $targetExists)',
    async ({ location, targetExists }) => {
      const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))
      const outside = await mkdtemp(
        path.join(tmpdir(), 'docx-kit-mcp-outside-'),
      )
      try {
        const target = path.join(
          location === 'inside' ? directory : outside,
          'original.docx',
        )
        const link = path.join(directory, 'linked.docx')
        if (targetExists) {
          await writeFile(target, 'original')
        }
        await symlink(target, link, 'file')
        expect((await lstat(link)).isSymbolicLink()).toBe(true)

        await expect(
          createDocument(
            { outputPath: 'linked.docx', schema: { content: [] } },
            { outputDirectory: directory },
          ),
        ).rejects.toThrow('Output path must not be a symbolic link')

        const contentsOrError = await readFile(target, 'utf8').catch(
          (error: unknown) => error,
        )
        const missingFileError = expect.objectContaining({ code: 'ENOENT' })
        expect(contentsOrError).toEqual(
          targetExists ? 'original' : missingFileError,
        )
        expect((await lstat(link)).isSymbolicLink()).toBe(true)
      } finally {
        await Promise.all([
          rm(directory, { force: true, recursive: true }),
          rm(outside, { force: true, recursive: true }),
        ])
      }
    },
  )

  it.each([false, true])(
    'writes a regular file (already exists: %s)',
    async fileExists => {
      const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))
      try {
        const filePath = path.join(directory, 'result.docx')
        if (fileExists) {
          await writeFile(filePath, 'x'.repeat(100_000))
        }

        const result = await createDocument(
          { outputPath: 'result.docx', schema: { content: [] } },
          { outputDirectory: directory },
        )
        const bytes = await readFile(filePath)

        expect(result.filePath).toBe(filePath)
        expect(result.size).toBe(bytes.byteLength)
        expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b])
        expect(bytes.byteLength).toBeLessThan(100_000)
      } finally {
        await rm(directory, { force: true, recursive: true })
      }
    },
  )
})
