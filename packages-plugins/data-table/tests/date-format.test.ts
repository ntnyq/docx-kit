import { execFileSync } from 'node:child_process'
import { createRequire } from 'node:module'
import path from 'node:path'
import process from 'node:process'
import { pathToFileURL } from 'node:url'
import { expect, it } from 'vitest'

const require = createRequire(import.meta.url)
const tsxLoader = pathToFileURL(require.resolve('tsx')).href
const source = pathToFileURL(
  path.resolve(import.meta.dirname, '../src/index.ts'),
).href
const script = `
  import { dataTablePlugin } from ${JSON.stringify(source)}
  import { Document, Packer } from 'docx'
  import JSZip from 'jszip'
  const table = await dataTablePlugin().render({
    data: [
      { date: '2026-11-13' },
      { date: '2024-02-29' },
      { date: '2026-02-30' },
      { date: 'not-a-date' },
      { date: new Date(2026, 10, 13) },
      { date: '2026-11-13T00:00:00Z' },
      { date: new Date(NaN) },
    ],
    format: { date: 'date' },
  }, { config: {} })
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [table] }] })))
  const xml = await zip.file('word/document.xml').async('string')
  console.log(JSON.stringify([...xml.matchAll(/<w:t[^>]*>([^<]*)<\\/w:t>/g)].map(match => match[1])))
`

it.each([
  ['UTC', '2026-11-13'],
  ['America/Los_Angeles', '2026-11-12'],
  ['Asia/Shanghai', '2026-11-13'],
])(
  'preserves civil dates while formatting instants locally in %s',
  (zone, instantDate) => {
    const output = execFileSync(
      process.execPath,
      ['--import', tsxLoader, '--input-type=module', '-e', script],
      {
        cwd: path.resolve(import.meta.dirname, '..'),
        encoding: 'utf8',
        env: { ...process.env, TSX_DISABLE_CACHE: '1', TZ: zone },
      },
    )
    expect(JSON.parse(output)).toEqual([
      'date',
      '2026-11-13',
      '2024-02-29',
      '2026-02-30',
      'not-a-date',
      '2026-11-13',
      instantDate,
      'Invalid Date',
    ])
  },
)
