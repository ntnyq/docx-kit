import { existsSync, readdirSync, statSync } from 'node:fs'
import { basename, join, relative, resolve } from 'node:path'
import process from 'node:process'

const KIBIBYTE = 1024
const MEBIBYTE = 1024 * KIBIBYTE
const budgets = {
  editorWorker: 350 * KIBIBYTE,
  lazyChunk: 3 * MEBIBYTE,
  playgroundRoute: 512 * KIBIBYTE,
  typescriptWorker: 7 * MEBIBYTE,
}
const forbiddenWorkers = new Set(['css.worker', 'html.worker', 'json.worker'])
const repositoryRoot = resolve(import.meta.dirname, '..')
const assetsRoot = join(repositoryRoot, 'apps/playground/dist/assets')
interface Asset {
  bytes: number
  path: string
}

const errors: string[] = []

if (!existsSync(assetsRoot)) {
  console.error('Playground assets are missing. Run pnpm apps:build first.')
  process.exit(1)
}

const assets = collectFiles(assetsRoot)
  .filter(path => path.endsWith('.js'))
  .map(path => ({
    bytes: statSync(path).size,
    name: basename(path),
    path,
  }))
const playgroundChunks = assets.filter(asset =>
  /^index-.+\.js$/u.test(asset.name),
)
const editorWorkers = assets.filter(asset =>
  /^editor\.worker-.+\.js$/u.test(asset.name),
)
const typescriptWorkers = assets.filter(asset =>
  /^ts\.worker-.+\.js$/u.test(asset.name),
)

requireAssets('playground route chunk', playgroundChunks)
requireAssets('Monaco editor worker', editorWorkers, 1)
requireAssets('Monaco TypeScript worker', typescriptWorkers, 1)

for (const asset of assets) {
  const workerName = asset.name.split('-')[0]
  if (forbiddenWorkers.has(workerName)) {
    errors.push(
      `${displayPath(asset.path)} should not be emitted by the TypeScript-only playground`,
    )
  }
}

checkBudgets(playgroundChunks, budgets.playgroundRoute, 'playground route')
checkBudgets(editorWorkers, budgets.editorWorker, 'editor worker')
checkBudgets(typescriptWorkers, budgets.typescriptWorker, 'TypeScript worker')
checkBudgets(
  assets.filter(asset => !asset.name.includes('.worker-')),
  budgets.lazyChunk,
  'lazy JavaScript chunk',
)

if (errors.length > 0) {
  console.error('Playground bundle audit failed:')
  errors.forEach(error => console.error(`  - ${error}`))
  process.exitCode = 1
} else {
  const largestLazyChunk = assets
    .filter(asset => !asset.name.includes('.worker-'))
    .sort((left, right) => right.bytes - left.bytes)[0]
  const largestPlaygroundChunk = [...playgroundChunks].sort(
    (left, right) => right.bytes - left.bytes,
  )[0]

  console.log(
    [
      'Playground bundle audit passed.',
      `Playground route: ${formatBytes(largestPlaygroundChunk.bytes)} / ${formatBytes(budgets.playgroundRoute)}.`,
      `Largest lazy chunk: ${formatBytes(largestLazyChunk.bytes)} / ${formatBytes(budgets.lazyChunk)}.`,
      `TypeScript worker: ${formatBytes(typescriptWorkers[0].bytes)} / ${formatBytes(budgets.typescriptWorker)}.`,
    ].join(' '),
  )
}

function checkBudgets(entries: Asset[], maximumBytes: number, label: string) {
  for (const entry of entries) {
    if (entry.bytes > maximumBytes) {
      errors.push(
        `${label} ${displayPath(entry.path)} is ${formatBytes(entry.bytes)}; budget is ${formatBytes(maximumBytes)}`,
      )
    }
  }
}

function collectFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    const path = join(directory, entry.name)
    return entry.isDirectory() ? collectFiles(path) : [path]
  })
}

function displayPath(path: string) {
  return relative(repositoryRoot, path)
}

function formatBytes(bytes: number) {
  if (bytes >= MEBIBYTE) {
    return `${(bytes / MEBIBYTE).toFixed(2)} MiB`
  }
  return `${(bytes / KIBIBYTE).toFixed(1)} KiB`
}

function requireAssets(label: string, entries: Asset[], expectedMinimum = 1) {
  if (entries.length < expectedMinimum) {
    errors.push(`missing ${label}`)
  }
}
