import { spawnSync } from 'node:child_process'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
  utimesSync,
  writeFileSync,
} from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import process from 'node:process'
import { pathToFileURL } from 'node:url'
import { afterEach, beforeEach, expect, it } from 'vitest'
import {
  checkMonacoTypes,
  generateMonacoTypes,
} from '../../scripts/generate-monaco-types'

const repositoryRoot = resolve(import.meta.dirname, '../..')
const require = createRequire(import.meta.url)
const tsxLoader = pathToFileURL(require.resolve('tsx')).href
let root: string

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'docx-kit-monaco-types-'))
})

afterEach(() => rmSync(root, { force: true, recursive: true }))

function createCliFixture() {
  for (const directory of [
    'packages',
    'packages-plugins',
    'packages-presets',
    'packages-themes',
    'node_modules',
  ]) {
    symlinkSync(
      join(repositoryRoot, directory),
      join(root, directory),
      'junction',
    )
  }
  writeFileSync(join(root, 'package.json'), '{"type":"module"}\n')
  const script = join(root, 'scripts/generate-monaco-types.ts')
  mkdirSync(dirname(script))
  copyFileSync(join(repositoryRoot, 'scripts/generate-monaco-types.ts'), script)
  const output = join(
    root,
    'apps/playground/src/utils/monacoTypes.generated.ts',
  )
  mkdirSync(dirname(output), { recursive: true })
  return {
    output,
    run(check = true) {
      return spawnSync(
        process.execPath,
        ['--import', tsxLoader, script, ...(check ? ['--check'] : [])],
        {
          cwd: root,
          encoding: 'utf8',
          env: { ...process.env, TSX_DISABLE_CACHE: '1' },
          timeout: 30_000,
        },
      )
    },
  }
}

function createSource() {
  const sourcePath = join(root, 'index.d.ts')
  writeFileSync(sourcePath, 'export declare const version: "1"\n')
  return [
    { sourcePath, virtualPath: 'file:///node_modules/example/index.d.ts' },
  ]
}

it('detects changed declaration inputs without rewriting the generated file', () => {
  const sources = createSource()
  const output = join(root, 'generated.ts')
  const content = generateMonacoTypes(sources)
  writeFileSync(output, content)
  expect(checkMonacoTypes(output, sources)).toEqual([])
  writeFileSync(sources[0].sourcePath, 'export declare const version: "2"\n')
  expect(checkMonacoTypes(output, sources)).toEqual([`Outdated: ${output}`])
  expect(readFileSync(output, 'utf8')).toBe(content)
})

it('reports a missing output without creating it', () => {
  const sources = createSource()
  const output = join(root, 'missing.ts')
  expect(checkMonacoTypes(output, sources)).toEqual([`Missing: ${output}`])
  expect(existsSync(output)).toBe(false)
})

it('rejects missing declaration inputs without touching an existing output', () => {
  const sources = createSource()
  const output = join(root, 'generated.ts')
  writeFileSync(output, 'keep this content')
  rmSync(sources[0].sourcePath)
  expect(() => checkMonacoTypes(output, sources)).toThrow('missing declaration')
  expect(readFileSync(output, 'utf8')).toBe('keep this content')
})

it('accepts synchronized uncommitted output without a Git repository', () => {
  const fixture = createCliFixture()
  const content = generateMonacoTypes()
  writeFileSync(fixture.output, content)
  utimesSync(fixture.output, 1, 1)
  const modifiedAt = statSync(fixture.output).mtimeMs
  const result = fixture.run()
  expect(result.error).toBeUndefined()
  expect(result.stderr).toBe('')
  expect(result.status).toBe(0)
  expect(readFileSync(fixture.output, 'utf8')).toBe(content)
  expect(statSync(fixture.output).mtimeMs).toBe(modifiedAt)
})

it('fails CLI checks for stale output and leaves the bytes unchanged', () => {
  const fixture = createCliFixture()
  writeFileSync(fixture.output, 'outdated content\n')
  const result = fixture.run()
  expect(result.error).toBeUndefined()
  expect(result.status).toBe(1)
  expect(result.stderr).toContain('Outdated:')
  expect(result.stderr).toContain('generate:monaco-types')
  expect(readFileSync(fixture.output, 'utf8')).toBe('outdated content\n')
})

it('reports missing CLI output and preserves the normal generation command', () => {
  const fixture = createCliFixture()
  const result = fixture.run()
  expect(result.status).toBe(1)
  expect(result.stderr).toContain('Missing:')
  expect(existsSync(fixture.output)).toBe(false)
  const generated = fixture.run(false)
  expect(generated.stderr).toBe('')
  expect(generated.status).toBe(0)
  expect(readFileSync(fixture.output, 'utf8')).toBe(generateMonacoTypes())
  expect(fixture.run().status).toBe(0)
})

it('rejects stale declarations before apps:verify can run the application build', () => {
  const fixture = createCliFixture()
  const manifest = JSON.parse(
    readFileSync(join(repositoryRoot, 'package.json'), 'utf8'),
  ) as {
    scripts: Record<string, string>
  }
  writeFileSync(
    join(root, 'package.json'),
    JSON.stringify({
      type: 'module',
      scripts: {
        'apps:build': `node -e "require('node:fs').writeFileSync('build-ran', 'ran')"`,
        'apps:bundle': 'node --version',
        'apps:generated': `node --import ${tsxLoader} scripts/generate-monaco-types.ts --check`,
        'apps:typecheck': 'node --version',
        'apps:verify': manifest.scripts['apps:verify'],
      },
    }),
  )
  writeFileSync(fixture.output, 'outdated content\n')
  const result = spawnSync('pnpm', ['run', 'apps:verify'], {
    cwd: root,
    encoding: 'utf8',
    env: { ...process.env, TSX_DISABLE_CACHE: '1' },
    shell: process.platform === 'win32',
    timeout: 30_000,
  })
  expect(result.error).toBeUndefined()
  expect(result.status).not.toBe(0)
  expect(result.stderr).toContain('Outdated:')
  expect(existsSync(join(root, 'build-ran'))).toBe(false)
  expect(readFileSync(fixture.output, 'utf8')).toBe('outdated content\n')
})
