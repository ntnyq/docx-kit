import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import path, { dirname, join, relative } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { DIAGNOSTIC_CODES, ERROR_CODES } from '../../packages/errors/src/index'
import {
  checkErrorDocs,
  generateErrorDocs,
} from '../../scripts/generate-error-docs'

vi.mock('node:path', async importOriginal => {
  const original = await importOriginal<typeof import('node:path')>()
  return {
    ...original,
    relative: vi.fn(original.relative),
  }
})

afterEach(() => vi.mocked(relative).mockImplementation(path.relative))

describe('error documentation synchronization', () => {
  it('generates three complete registries with localized repair pages', () => {
    const pages = generateErrorDocs()
    expect(pages.size).toBe(57)
    for (const prefix of ['', '/zh-CN', '/zh-TW']) {
      const index = pages.get(`docs${prefix}/errors/index.md`)
      expect(index).toContain(`${prefix}/errors/docxkit_plugin_not_registered/`)
      expect(index).toContain('DOCXKIT_HEADER_FOOTER_INVALID_CONTENT')
      const page = pages.get(
        `docs${prefix}/errors/docxkit_plugin_not_registered/index.md`,
      )
      expect(page).toContain('.use(qrcodePlugin())')
      expect(page).toContain(`${prefix}/guide/errors`)
      expect(page).toContain('`code`: `PLUGIN_NOT_REGISTERED`')
    }
  })

  it('detects missing, outdated and surplus pages without changing files', () => {
    const root = mkdtempSync(join(tmpdir(), 'docx-kit-error-docs-'))
    try {
      for (const [path, text] of generateErrorDocs()) {
        mkdirSync(dirname(join(root, path)), { recursive: true })
        writeFileSync(join(root, path), text)
      }
      expect(checkErrorDocs(root)).toEqual([])
      const changed = 'docs/errors/index.md'
      writeFileSync(join(root, changed), 'Manual change')
      const missing = 'docs/zh-CN/errors/docxkit_plugin_not_registered/index.md'
      rmSync(join(root, missing))
      const surplus = 'docs/zh-TW/errors/retired/index.md'
      mkdirSync(dirname(join(root, surplus)), { recursive: true })
      writeFileSync(join(root, surplus), 'Unexpected')
      expect(checkErrorDocs(root)).toEqual(
        expect.arrayContaining([
          `Outdated: ${changed}`,
          `Missing: ${missing}`,
          `Unexpected: ${surplus}`,
        ]),
      )
      expect(readFileSync(join(root, changed), 'utf8')).toBe('Manual change')
    } finally {
      rmSync(root, { force: true, recursive: true })
    }
  })
  it('accepts synchronized pages when the host returns Windows path separators', () => {
    const root = mkdtempSync(join(tmpdir(), 'docx-kit-windows-error-docs-'))
    try {
      for (const [page, content] of generateErrorDocs()) {
        mkdirSync(dirname(join(root, page)), { recursive: true })
        writeFileSync(join(root, page), content)
      }
      vi.mocked(relative).mockImplementation((from, to) =>
        path.relative(from, to).replaceAll('/', '\\'),
      )
      expect(checkErrorDocs(root)).toEqual([])
      writeFileSync(join(root, 'docs/zh-CN/errors/index.md'), 'Changed')
      expect(checkErrorDocs(root)).toEqual([
        'Outdated: docs/zh-CN/errors/index.md',
      ])
    } finally {
      rmSync(root, { force: true, recursive: true })
    }
  })

  it('includes a real localized page for every runtime documentation URL', () => {
    const pages = generateErrorDocs()
    for (const code of Object.values(ERROR_CODES)) {
      const slug = DIAGNOSTIC_CODES[code].toLowerCase()
      for (const prefix of ['', '/zh-CN', '/zh-TW']) {
        const page = pages.get(`docs${prefix}/errors/${slug}/index.md`)
        expect(page).toBeDefined()
        expect(page).toContain(`https://docx-kit.ntnyq.dev/errors/${slug}/`)
        expect(page).toContain(`](${prefix}/errors/)`)
      }
    }
  })
})
