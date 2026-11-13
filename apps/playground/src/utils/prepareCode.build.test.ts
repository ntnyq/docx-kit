import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { build } from 'vite'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { sucraseTypeScriptProfile } from '../../vite.config'
import { prepareCode } from './prepareCode'
import { COMPILER_FIXTURES } from './prepareCode.fixtures'

let prepareProductionCode: (source: string) => unknown
let compilerDirectory: string | undefined

afterAll(async () => {
  if (compilerDirectory) {
    await rm(compilerDirectory, { force: true, recursive: true })
  }
})

beforeAll(async () => {
  const root = resolve(import.meta.dirname, '../..')
  const result = await build({
    configFile: false,
    logLevel: 'silent',
    plugins: [sucraseTypeScriptProfile()],
    root,
    build: {
      minify: true,
      write: false,
      lib: {
        entry: resolve(root, 'src/utils/prepareCode.ts'),
        formats: ['es'],
      },
    },
  })
  const output = Array.isArray(result) ? result[0] : result
  if (!output || !('output' in output)) {
    throw new Error('Expected one production compiler bundle')
  }
  const chunk = output.output.find(
    entry => entry.type === 'chunk' && entry.isEntry,
  )
  if (!chunk || chunk.type !== 'chunk') {
    throw new Error('Production compiler JavaScript was not generated')
  }
  compilerDirectory = await mkdtemp(join(tmpdir(), 'docx-kit-compiler-'))
  const compilerPath = join(compilerDirectory, 'compiler.mjs')
  await writeFile(compilerPath, chunk.code)
  const runtime: unknown = await import(
    /* @vite-ignore */ pathToFileURL(compilerPath).href
  )
  if (
    !runtime
    || typeof runtime !== 'object'
    || !('prepareCode' in runtime)
    || typeof runtime.prepareCode !== 'function'
  ) {
    throw new Error('Production compiler export is missing')
  }
  const prepare = runtime.prepareCode
  prepareProductionCode = source => {
    const code: unknown = prepare(source)
    return code
  }
})

describe('production TypeScript compiler profile', () => {
  it.each(COMPILER_FIXTURES)(
    'matches public Sucrase behavior for $name',
    ({ source }) => {
      expect(prepareProductionCode(source)).toBe(prepareCode(source))
    },
  )

  it.each(['const value: number =', 'return 42; const value ='])(
    'rejects invalid source: %s',
    source => {
      expect(() => prepareProductionCode(source)).toThrow(SyntaxError)
    },
  )
})
