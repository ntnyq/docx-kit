import { mkdtemp, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import process from 'node:process'
import prompts from 'prompts'
import { describe, expect, it, vi } from 'vitest'
import { createPlugin } from '../src/commands/create-plugin'

describe('scaffold cancellation', () => {
  it.each([1, 2, 3, 4])(
    'does not write files when cancelled after %i answers',
    async count => {
      const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-cancel-'))
      const cwd = vi.spyOn(process, 'cwd').mockReturnValue(directory)
      const log = vi.spyOn(console, 'log').mockImplementation(() => {})
      try {
        prompts.inject([
          ...['cancelled-plugin', 'Description', 'Author', 'MIT'].slice(
            0,
            count,
          ),
          new Error('Simulated cancellation'),
        ])
        await createPlugin()
        expect(await readdir(directory)).toEqual([])
        expect(log).toHaveBeenCalledWith('Scaffold cancelled.')
      } finally {
        cwd.mockRestore()
        log.mockRestore()
        await rm(directory, { force: true, recursive: true })
      }
    },
  )
})
