import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import process from 'node:process'

const repositoryRoot = resolve(import.meta.dirname, '..')
const packageRoots = [
  'packages',
  'packages-plugins',
  'packages-presets',
  'packages-themes',
]

function collectLocalPaths(value, field, paths) {
  if (typeof value === 'string') {
    if (value.startsWith('./')) {
      paths.push({ field, path: value })
    }
    return
  }

  if (Array.isArray(value)) {
    value.forEach((item, index) =>
      collectLocalPaths(item, `${field}[${index}]`, paths),
    )
    return
  }

  if (value && typeof value === 'object') {
    Object.entries(value).forEach(([key, item]) =>
      collectLocalPaths(item, `${field}.${key}`, paths),
    )
  }
}

function collectPackageDirectories() {
  return packageRoots.flatMap(packageRoot => {
    const absoluteRoot = join(repositoryRoot, packageRoot)
    return readdirSync(absoluteRoot, { withFileTypes: true })
      .filter(entry => entry.isDirectory())
      .map(entry => join(absoluteRoot, entry.name))
      .filter(packageDirectory =>
        existsSync(join(packageDirectory, 'package.json')),
      )
  })
}

const missingArtifacts = []

for (const packageDirectory of collectPackageDirectories()) {
  const manifestPath = join(packageDirectory, 'package.json')
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))

  if (manifest.private) {
    continue
  }

  const artifactPaths = []
  for (const field of [
    'bin',
    'browser',
    'exports',
    'main',
    'module',
    'types',
  ]) {
    collectLocalPaths(manifest[field], field, artifactPaths)
  }

  for (const artifact of artifactPaths) {
    if (!existsSync(resolve(packageDirectory, artifact.path))) {
      missingArtifacts.push(
        `${manifest.name}: ${artifact.field} -> ${artifact.path}`,
      )
    }
  }
}

if (missingArtifacts.length > 0) {
  console.error('Missing package artifacts:')
  missingArtifacts.forEach(artifact => console.error(`  - ${artifact}`))
  process.exitCode = 1
} else {
  console.log('All published package entrypoints resolve to built artifacts.')
}
