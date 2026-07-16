import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createPlugin } from '../src/commands/create-plugin'

// Mock prompts to avoid interactive input during tests
vi.mock('prompts', () => ({
  default: vi.fn(async questions => {
    const answers: Record<string, string> = {}
    for (const q of questions) {
      answers[q.name] = q.initial ?? 'test-answer'
    }
    return answers
  }),
}))

describe('createPlugin', () => {
  const tmpDir = path.join(process.cwd(), '__test_scaffold__')

  afterEach(async () => {
    vi.restoreAllMocks()
    // Clean up test directories
    await fs.rm(tmpDir, { force: true, recursive: true }).catch(() => {})
    await fs
      .rm(path.join(process.cwd(), 'test-answer'), {
        force: true,
        recursive: true,
      })
      .catch(() => {})
  })

  it('creates plugin project directory with all files', async () => {
    // Override cwd for testing
    vi.spyOn(process, 'cwd').mockReturnValue(tmpDir)

    await createPlugin('docx-kit-plugin-test')

    // Check all expected files exist
    const projectDir = path.join(tmpDir, 'docx-kit-plugin-test')

    const expectedFiles = [
      'docx-kit.plugin.json',
      'eslint.config.mjs',
      path.join('src', 'index.ts'),
      path.join('tests', 'index.test.ts'),
      'package.json',
      'tsconfig.json',
      'tsdown.config.ts',
      'README.md',
    ]

    for (const file of expectedFiles) {
      await expect(
        fs.access(path.join(projectDir, file)),
      ).resolves.toBeUndefined()
    }

    // Verify manifest content
    const manifest = JSON.parse(
      await fs.readFile(path.join(projectDir, 'docx-kit.plugin.json'), 'utf-8'),
    )
    expect(manifest.name).toBe('docx-kit-plugin-test')
    expect(manifest.plugin.name).toBe('test')
    expect(manifest.docxKit).toBe('^0.3.0')

    // Verify package.json content
    const pkg = JSON.parse(
      await fs.readFile(path.join(projectDir, 'package.json'), 'utf-8'),
    )
    expect(pkg.name).toBe('docx-kit-plugin-test')
    expect(pkg.keywords).toContain('docx-kit-plugin')
    expect(pkg.devDependencies.tsdown).toBe('^0.22.9')
  })

  it('does not overwrite existing directory', async () => {
    // Create an existing directory
    await fs.mkdir(tmpDir, { recursive: true })
    await fs.mkdir(path.join(tmpDir, 'existing-project'), { recursive: true })

    vi.spyOn(process, 'cwd').mockReturnValue(tmpDir)

    const consoleErrorSpy = vi
      .spyOn(console, 'error')
      .mockImplementation(() => {})

    // This should fail because directory already exists
    await createPlugin('existing-project')

    expect(consoleErrorSpy).toHaveBeenCalled()

    consoleErrorSpy.mockRestore()
  })

  it('handles user cancellation gracefully', async () => {
    // Mock prompts to return null (user cancelled)
    const promptsModule = await import('prompts')
    vi.mocked(promptsModule.default).mockResolvedValueOnce(null as any)

    const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => {})

    await createPlugin()

    expect(consoleLogSpy).toHaveBeenCalledWith('Scaffold cancelled.')

    consoleLogSpy.mockRestore()
  })
})
