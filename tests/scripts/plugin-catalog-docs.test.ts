import { readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const repositoryRoot = resolve(import.meta.dirname, '../..')
const pluginRoot = join(repositoryRoot, 'packages-plugins')
const pluginSlugs = readdirSync(pluginRoot, { withFileTypes: true })
  .filter(entry => entry.isDirectory())
  .map(entry => entry.name)

describe('localized plugin catalogs', () => {
  it.each(['', 'zh-CN', 'zh-TW'])(
    'keeps complete catalog rows aligned with published packages in %s',
    locale => {
      const source = readFileSync(
        join(repositoryRoot, 'docs', locale, 'plugins/index.md'),
        'utf8',
      )
      const tables = source
        .split(/\n\s*\n/u)
        .filter(block => block.startsWith('|'))
      expect(tables.length).toBeGreaterThanOrEqual(2)
      const catalog = tables[0].split('\n').slice(2)
      const documentedSlugs: string[] = []
      for (const row of catalog) {
        const cells = row
          .split(/(?<!\\)\|/u)
          .slice(1, -1)
          .map(cell => cell.trim())
        expect(cells).toHaveLength(4)
        const slug = cells[0].match(/\]\(\.\/([a-z-]+)\)/u)?.[1]
        if (!slug || !pluginSlugs.includes(slug)) {
          throw new Error(`Unknown plugin documentation link: ${cells[0]}`)
        }
        const metadata: { name: string } = JSON.parse(
          readFileSync(join(pluginRoot, slug, 'package.json'), 'utf8'),
        )
        const runtimeName = slug.replace(/-([a-z])/gu, (_, letter: string) =>
          letter.toUpperCase(),
        )
        expect(cells[1]).toBe(`\`${metadata.name}\``)
        expect(cells[2]).toBe(`\`${runtimeName}\``)
        expect(cells[3]).not.toBe('')
        documentedSlugs.push(slug)
      }
      expect(documentedSlugs).toHaveLength(pluginSlugs.length)
      expect(new Set(documentedSlugs)).toEqual(new Set(pluginSlugs))
      const compatibilityRows = tables[1].split('\n').slice(2)
      expect(
        compatibilityRows.some(row => row.startsWith('| `crossReference` |')),
      ).toBe(true)
    },
  )
})
