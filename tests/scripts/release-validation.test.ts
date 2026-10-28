import { spawnSync } from 'node:child_process'
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import process from 'node:process'
import { pathToFileURL } from 'node:url'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

const repositoryRoot = resolve(import.meta.dirname, '../..')
const require = createRequire(import.meta.url)
const tsxLoader = pathToFileURL(require.resolve('tsx')).href
let temporaryRoot: string

beforeEach(() => {
  temporaryRoot = mkdtempSync(join(tmpdir(), 'docx-kit-release-validation-'))
  for (const directory of [
    'docs',
    'packages',
    'packages-plugins',
    'packages-presets',
    'packages-themes',
    'scripts',
  ]) {
    mkdirSync(join(temporaryRoot, directory))
  }
  writeFileSync(
    join(temporaryRoot, 'package.json'),
    JSON.stringify({ private: true, type: 'module' }),
  )
  writeFileSync(join(temporaryRoot, 'README.md'), '# Fixture\n')
})

afterEach(() => {
  rmSync(temporaryRoot, { force: true, recursive: true })
})

function createPackage(files: string[]) {
  const directory = join(temporaryRoot, 'packages', 'fixture')
  mkdirSync(join(directory, 'dist'), { recursive: true })
  writeFileSync(join(directory, 'dist/index.js'), 'export {}\n')
  writeFileSync(join(directory, 'dist/index.d.ts'), 'export {}\n')
  writeFileSync(join(directory, 'README.md'), '# Fixture package\n')
  writeFileSync(
    join(directory, 'package.json'),
    JSON.stringify({
      files,
      name: 'docx-kit-release-validation-fixture',
      type: 'module',
      version: '1.0.0',
      exports: {
        '.': { default: './dist/index.js', types: './dist/index.d.ts' },
      },
    }),
  )
  return directory
}

function runScript(name: string) {
  const script = join(temporaryRoot, 'scripts', name)
  copyFileSync(join(repositoryRoot, 'scripts', name), script)
  return spawnSync(process.execPath, ['--import', tsxLoader, script], {
    cwd: temporaryRoot,
    encoding: 'utf8',
    env: process.env,
    timeout: 30_000,
  })
}

describe('published artifact validation', () => {
  it('accepts entrypoints included in the packed package', () => {
    createPackage(['dist'])
    const result = runScript('check-package-artifacts.ts')
    expect(result.error).toBeUndefined()
    expect(result.status).toBe(0)
  })

  it('rejects existing entrypoints excluded by files', () => {
    createPackage(['README.md'])
    const result = runScript('check-package-artifacts.ts')
    expect(result.status).toBe(1)
    expect(result.stderr).toContain('not included in packed files')
    expect(result.stderr).toContain('dist/index.js')
  })

  it('rejects missing entrypoints without a dot-slash prefix', () => {
    const directory = createPackage(['dist'])
    writeFileSync(
      join(directory, 'package.json'),
      JSON.stringify({
        files: ['dist'],
        main: 'dist/missing.js',
        name: 'docx-kit-release-validation-fixture',
        version: '1.0.0',
      }),
    )
    const result = runScript('check-package-artifacts.ts')
    expect(result.status).toBe(1)
    expect(result.stderr).toContain('dist/missing.js')
  })
})

describe('barcode documentation validation', () => {
  it.each(['doc.use(barcodePlugin)', '{ plugin: barcodePlugin }'])(
    'rejects an uncalled factory in %s',
    source => {
      writeFileSync(
        join(temporaryRoot, 'README.md'),
        `\`\`\`ts\n${source}\n\`\`\`\n`,
      )
      const result = runScript('check-doc-examples.ts')
      expect(result.status).toBe(1)
      expect(result.stderr).toContain('barcodePlugin')
    },
  )

  it('accepts a called barcode factory', () => {
    writeFileSync(
      join(temporaryRoot, 'README.md'),
      '```ts\ndoc.use(barcodePlugin())\n```\n',
    )
    const result = runScript('check-doc-examples.ts')
    expect(result.error).toBeUndefined()
    expect(result.status).toBe(0)
  })
})
