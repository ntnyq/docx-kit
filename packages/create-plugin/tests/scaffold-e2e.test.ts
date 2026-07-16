import { execFileSync } from 'node:child_process'
import fs from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import process from 'node:process'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createPlugin } from '../src/commands/create-plugin'

vi.mock('prompts', () => ({
  default: vi.fn(async questions => {
    const answers: Record<string, string> = {}
    for (const question of questions) {
      answers[question.name] = question.initial ?? 'test-answer'
    }
    return answers
  }),
}))

const repositoryRoot = path.resolve(import.meta.dirname, '../../..')

function runPnpm(cwd: string, args: string[]): void {
  try {
    execFileSync('pnpm', args, {
      cwd,
      encoding: 'utf8',
      env: { ...process.env, CI: '1' },
      timeout: 120_000,
    })
  } catch (err) {
    const failure = err as Error & { stderr?: string; stdout?: string }
    throw new Error(
      [failure.message, failure.stdout, failure.stderr]
        .filter(Boolean)
        .join('\n'),
      { cause: err },
    )
  }
}

describe('generated plugin project', () => {
  let temporaryRoot: string | undefined

  afterEach(async () => {
    vi.restoreAllMocks()
    if (temporaryRoot) {
      await fs.rm(temporaryRoot, { force: true, recursive: true })
    }
  })

  it('installs, lints, typechecks, tests, and builds', async () => {
    temporaryRoot = await fs.mkdtemp(
      path.join(tmpdir(), 'docx-kit-plugin-scaffold-'),
    )
    vi.spyOn(process, 'cwd').mockReturnValue(temporaryRoot)
    vi.spyOn(console, 'log').mockImplementation(() => {})

    await createPlugin('docx-kit-plugin-smoke')

    const projectDirectory = path.join(temporaryRoot, 'docx-kit-plugin-smoke')
    const packagePath = path.join(projectDirectory, 'package.json')
    const packageJson = JSON.parse(await fs.readFile(packagePath, 'utf8')) as {
      devDependencies: Record<string, string>
    }
    const localDependencies: Record<string, string> = {
      docx: path.join(repositoryRoot, 'packages/core/node_modules/docx'),
      'docx-kit': path.join(repositoryRoot, 'packages/docx-kit'),
      eslint: path.join(repositoryRoot, 'node_modules/eslint'),
      tsdown: path.join(repositoryRoot, 'node_modules/tsdown'),
      typescript: path.join(repositoryRoot, 'node_modules/typescript'),
      vitest: path.join(repositoryRoot, 'node_modules/vitest'),
      '@ntnyq/eslint-config': path.join(
        repositoryRoot,
        'node_modules/@ntnyq/eslint-config',
      ),
    }

    for (const [name, dependencyPath] of Object.entries(localDependencies)) {
      packageJson.devDependencies[name] = `link:${dependencyPath}`
    }
    await fs.writeFile(packagePath, `${JSON.stringify(packageJson, null, 2)}\n`)

    runPnpm(projectDirectory, [
      'install',
      '--ignore-scripts',
      '--offline',
      '--store-dir',
      path.join(temporaryRoot, '.pnpm-store'),
    ])
    runPnpm(projectDirectory, ['run', 'lint'])
    runPnpm(projectDirectory, ['run', 'typecheck'])
    runPnpm(projectDirectory, ['run', 'test'])
    runPnpm(projectDirectory, ['run', 'build'])

    await expect(
      fs.access(path.join(projectDirectory, 'dist/index.js')),
    ).resolves.toBeUndefined()
    await expect(
      fs.access(path.join(projectDirectory, 'dist/index.d.ts')),
    ).resolves.toBeUndefined()
  }, 120_000)
})
