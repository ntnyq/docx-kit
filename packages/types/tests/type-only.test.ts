import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import ts from 'typescript'
import { describe, expect, it } from 'vitest'

const packageRoot = resolve(import.meta.dirname, '..')
const sourceRoot = resolve(packageRoot, 'src')
const sourceFiles = readdirSync(sourceRoot, {
  encoding: 'utf8',
  recursive: true,
}).filter(file => file.endsWith('.ts'))

describe('@docxkit/types package boundary', () => {
  it('exposes only a declaration entry point', () => {
    const manifest = JSON.parse(
      readFileSync(resolve(packageRoot, 'package.json'), 'utf8'),
    )

    expect(manifest.exports).toEqual({
      '.': { types: './dist/index.d.ts' },
    })
    expect(manifest.main).toBeUndefined()
    expect(manifest.module).toBeUndefined()
  })

  it.each(sourceFiles)('%s contains no runtime code', file => {
    const source = readFileSync(resolve(sourceRoot, file), 'utf8')
    const { outputText } = ts.transpileModule(source, {
      fileName: file,
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        removeComments: true,
        target: ts.ScriptTarget.ESNext,
        verbatimModuleSyntax: true,
      },
    })

    expect(outputText.trim()).toBe('export {};')
  })
})
