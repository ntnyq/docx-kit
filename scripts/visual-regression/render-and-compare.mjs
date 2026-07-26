import { spawnSync } from 'node:child_process'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { basename, join, relative, resolve } from 'node:path'
import process from 'node:process'
import { pathToFileURL } from 'node:url'
import pixelmatch from 'pixelmatch'
import { PNG } from 'pngjs'

const updateBaselines = process.argv.includes('--update')
const repositoryRoot = resolve(import.meta.dirname, '../..')
const temporaryRoot = join(repositoryRoot, 'tmp/visual-regression')
const docxRoot = join(temporaryRoot, 'docx')
const pdfRoot = join(temporaryRoot, 'pdf')
const currentRoot = join(temporaryRoot, 'current')
const profileRoot = join(temporaryRoot, 'libreoffice-profile')
const baselineRoot = join(repositoryRoot, 'tests/visual/baselines')
const artifactRoot = join(repositoryRoot, 'output/visual-regression')
const diffRoot = join(artifactRoot, 'diffs')
const manifestPath = join(temporaryRoot, 'manifest.json')
const maximumMismatchRatio = 0.01
const minimumInkRatio = 0.002
const errors = []

if (!existsSync(manifestPath)) {
  console.error(
    'Visual fixture manifest is missing. Run visual:generate first.',
  )
  process.exit(1)
}

const libreOffice = resolveBinary(
  process.env.LIBREOFFICE_BIN,
  'soffice',
  'libreoffice',
)
const pdfToPpm = resolveBinary(process.env.PDFTOPPM_BIN, 'pdftoppm')
const pdfInfo = resolveBinary(process.env.PDFINFO_BIN, 'pdfinfo')
const fixtures = JSON.parse(readFileSync(manifestPath, 'utf8'))

for (const directory of [pdfRoot, currentRoot, profileRoot, artifactRoot]) {
  rmSync(directory, { force: true, recursive: true })
  mkdirSync(directory, { recursive: true })
}
mkdirSync(diffRoot, { recursive: true })
mkdirSync(baselineRoot, { recursive: true })

for (const fixture of fixtures) {
  renderFixture(fixture)
}

const currentPages = listPngFiles(currentRoot)
const expectedBaselines = new Set(currentPages.map(path => basename(path)))

if (updateBaselines) {
  for (const path of listPngFiles(baselineRoot)) {
    if (!expectedBaselines.has(basename(path))) {
      rmSync(path)
    }
  }
  for (const path of currentPages) {
    copyFileSync(path, join(baselineRoot, basename(path)))
  }
  console.log(`Updated ${currentPages.length} visual regression baselines.`)
} else {
  comparePages(currentPages)
  detectStaleBaselines(expectedBaselines)
}

writeFileSync(
  join(artifactRoot, 'report.json'),
  `${JSON.stringify(
    {
      errors,
      fixtureCount: fixtures.length,
      libreOffice: getVersion(libreOffice),
      pageCount: currentPages.length,
      updateBaselines,
    },
    null,
    2,
  )}\n`,
)

if (errors.length > 0) {
  console.error('Visual regression failed:')
  errors.forEach(error => console.error(`  - ${error}`))
  process.exitCode = 1
} else if (!updateBaselines) {
  console.log(
    `Visual regression passed: ${fixtures.length} fixtures, ${currentPages.length} rendered pages.`,
  )
}

function comparePages(currentPages) {
  for (const currentPath of currentPages) {
    const name = basename(currentPath)
    const baselinePath = join(baselineRoot, name)

    if (!existsSync(baselinePath)) {
      errors.push(`missing baseline ${displayPath(baselinePath)}`)
      continue
    }

    const baseline = PNG.sync.read(readFileSync(baselinePath))
    const current = PNG.sync.read(readFileSync(currentPath))

    if (
      baseline.width !== current.width
      || baseline.height !== current.height
    ) {
      errors.push(
        `${name} dimensions changed from ${baseline.width}x${baseline.height} to ${current.width}x${current.height}`,
      )
      continue
    }

    const diff = new PNG({ height: current.height, width: current.width })
    const mismatchedPixels = pixelmatch(
      baseline.data,
      current.data,
      diff.data,
      current.width,
      current.height,
      { includeAA: false, threshold: 0.2 },
    )
    const mismatchRatio = mismatchedPixels / (current.width * current.height)

    if (mismatchRatio > maximumMismatchRatio) {
      const diffPath = join(diffRoot, name)
      writeFileSync(diffPath, PNG.sync.write(diff))
      errors.push(
        `${name} changed by ${(mismatchRatio * 100).toFixed(2)}% (allowed ${(maximumMismatchRatio * 100).toFixed(2)}%); diff: ${displayPath(diffPath)}`,
      )
    }
  }
}

function detectStaleBaselines(expectedBaselines) {
  for (const path of listPngFiles(baselineRoot)) {
    if (!expectedBaselines.has(basename(path))) {
      errors.push(`stale baseline ${displayPath(path)}`)
    }
  }
}

function displayPath(path) {
  return relative(repositoryRoot, path)
}

function getVersion(binary) {
  return run(binary, ['--version']).stdout.trim().split('\n')[0]
}

function listPngFiles(directory) {
  if (!existsSync(directory)) {
    return []
  }
  return readdirSync(directory)
    .filter(name => name.endsWith('.png'))
    .map(name => join(directory, name))
    .sort()
}

function renderFixture({ expectedPages, name }) {
  const docxPath = join(docxRoot, `${name}.docx`)
  const profileUrl = pathToFileURL(profileRoot).href

  run(libreOffice, [
    '--headless',
    `-env:UserInstallation=${profileUrl}`,
    '--convert-to',
    'pdf',
    '--outdir',
    pdfRoot,
    docxPath,
  ])

  const pdfPath = join(pdfRoot, `${name}.pdf`)
  if (!existsSync(pdfPath) || statSync(pdfPath).size === 0) {
    errors.push(`${name} did not produce a PDF`)
    return
  }

  const info = run(pdfInfo, [pdfPath]).stdout
  const pageCount = Number(info.match(/^Pages:\s+(\d+)$/mu)?.[1] ?? '')
  if (pageCount !== expectedPages) {
    errors.push(
      `${name} rendered ${pageCount || 'an unknown number of'} pages; expected ${expectedPages}`,
    )
  }

  const prefix = join(currentRoot, name)
  run(pdfToPpm, ['-png', '-r', '96', pdfPath, prefix])

  const renderedPages = listPngFiles(currentRoot).filter(path =>
    basename(path).startsWith(`${name}-`),
  )
  if (renderedPages.length !== pageCount) {
    errors.push(
      `${name} produced ${renderedPages.length} PNG pages for a ${pageCount}-page PDF`,
    )
  }

  for (const path of renderedPages) {
    const image = PNG.sync.read(readFileSync(path))
    let inkPixels = 0
    for (let index = 0; index < image.data.length; index += 4) {
      if (
        image.data[index] < 248
        || image.data[index + 1] < 248
        || image.data[index + 2] < 248
      ) {
        inkPixels += 1
      }
    }
    const inkRatio = inkPixels / (image.width * image.height)
    if (inkRatio < minimumInkRatio) {
      errors.push(
        `${displayPath(path)} appears blank (${(inkRatio * 100).toFixed(2)}% ink)`,
      )
    }
  }
}

function resolveBinary(...candidates) {
  const availableCandidates = candidates.filter(Boolean)

  for (const candidate of availableCandidates) {
    if (candidate.includes('/') && existsSync(candidate)) {
      return candidate
    }
    const result = spawnSync('which', [candidate], { encoding: 'utf8' })
    if (result.status === 0) {
      return result.stdout.trim()
    }
  }

  console.error(
    `Required binary not found: ${availableCandidates.join(' or ')}`,
  )
  process.exit(1)
}

function run(command, arguments_) {
  const result = spawnSync(command, arguments_, {
    encoding: 'utf8',
    env: process.env,
    maxBuffer: 10 * 1024 * 1024,
  })

  if (result.status !== 0) {
    console.error(result.stdout)
    console.error(result.stderr)
    throw new Error(
      `${basename(command)} exited with status ${result.status ?? 'unknown'}`,
    )
  }

  return result
}
