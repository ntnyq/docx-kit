import { execFileSync } from 'node:child_process'
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import process from 'node:process'

const repositoryRoot = resolve(import.meta.dirname, '..')
const packageRoots = [
  'packages',
  'packages-plugins',
  'packages-presets',
  'packages-themes',
]
const temporaryRoot = mkdtempSync(join(tmpdir(), 'docx-kit-package-smoke-'))
const tarballDirectory = join(temporaryRoot, 'tarballs')

function collectPackages() {
  return packageRoots.flatMap(packageRoot => {
    const absoluteRoot = join(repositoryRoot, packageRoot)
    return readdirSync(absoluteRoot, { withFileTypes: true })
      .filter(entry => entry.isDirectory())
      .map(entry => join(absoluteRoot, entry.name))
      .filter(packageDirectory =>
        existsSync(join(packageDirectory, 'package.json')),
      )
      .map(packageDirectory => ({
        directory: packageDirectory,
        manifest: JSON.parse(
          readFileSync(join(packageDirectory, 'package.json'), 'utf8'),
        ),
      }))
      .filter(pkg => !pkg.manifest.private)
  })
}

function run(command, args, cwd) {
  execFileSync(command, args, {
    cwd,
    env: process.env,
    stdio: 'inherit',
  })
}

function runSilently(command, args, cwd) {
  execFileSync(command, args, {
    cwd,
    env: process.env,
    stdio: 'pipe',
  })
}

let succeeded = false

try {
  const packages = collectPackages()
  const overrides = {}

  for (const pkg of packages) {
    const existingTarballs = new Set(
      existsSync(tarballDirectory) ? readdirSync(tarballDirectory) : [],
    )
    runSilently(
      'pnpm',
      ['pack', '--pack-destination', tarballDirectory],
      pkg.directory,
    )
    const tarball = readdirSync(tarballDirectory).find(
      entry => !existingTarballs.has(entry),
    )

    if (!tarball) {
      throw new Error(
        `pnpm pack did not create a tarball for ${pkg.manifest.name}`,
      )
    }

    overrides[pkg.manifest.name] = `file:${join(tarballDirectory, tarball)}`
  }

  const externalPeers = Object.assign(
    {},
    ...packages.map(pkg => pkg.manifest.peerDependencies ?? {}),
  )
  for (const workspacePackageName of Object.keys(overrides)) {
    delete externalPeers[workspacePackageName]
  }

  writeFileSync(
    join(temporaryRoot, 'package.json'),
    `${JSON.stringify(
      {
        name: 'docx-kit-package-smoke',
        private: true,
        type: 'module',
        dependencies: {
          ...externalPeers,
          'docx-kit': overrides['docx-kit'],
        },
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(
    join(temporaryRoot, 'pnpm-workspace.yaml'),
    `${JSON.stringify({ overrides, packages: [] }, null, 2)}\n`,
  )
  writeFileSync(
    join(temporaryRoot, 'smoke.mjs'),
    `import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

const expectedExports = [
  ['docx-kit', 'createDocx'],
  ['docx-kit/ai', 'buildPrompt'],
  ['docx-kit/browser', 'createDocx'],
  ['docx-kit/loader', 'createPluginLoader'],
  ['docx-kit/loader/browser', 'loadUrlPlugin'],
  ['docx-kit/loader/node', 'loadNpmPlugin'],
  ['docx-kit/mcp', 'TOOL_DEFINITIONS'],
  ['docx-kit/node', 'saveDocument'],
  ['docx-kit/pdk', 'createPluginTestContext'],
  ['docx-kit/registry', 'createPluginRegistry'],
]

for (const [specifier, exportName] of expectedExports) {
  const module = await import(specifier)
  if (!(exportName in module)) {
    throw new Error(\`Missing export \${exportName} from \${specifier}\`)
  }
}

const nodeModule = await import('docx-kit/node')
const outputPath = join(import.meta.dirname, 'consumer-smoke.docx')
await nodeModule
  .createDocx()
  .h1('Packed consumer smoke test')
  .p('Node.js save adapter is available.')
  .save(outputPath)

const output = await readFile(outputPath)
if (output.length < 4 || output[0] !== 0x50 || output[1] !== 0x4b) {
  throw new Error('Node.js save adapter did not produce a DOCX ZIP archive')
}
`,
  )
  writeFileSync(
    join(temporaryRoot, 'smoke.ts'),
    `import { buildPrompt } from 'docx-kit/ai'
import { createDocx } from 'docx-kit'
import { loadUrlPlugin } from 'docx-kit/loader/browser'
import { loadNpmPlugin } from 'docx-kit/loader/node'
import { TOOL_DEFINITIONS } from 'docx-kit/mcp'
import {
  createDocx as createNodeDocx,
  saveDocument,
} from 'docx-kit/node'
import { createPluginTestContext } from 'docx-kit/pdk'
import { createPluginRegistry } from 'docx-kit/registry'

void [
  buildPrompt,
  createDocx,
  createPluginRegistry,
  createPluginTestContext,
  loadNpmPlugin,
  loadUrlPlugin,
  saveDocument,
  TOOL_DEFINITIONS,
]

void createNodeDocx().save('consumer-smoke.docx')
`,
  )

  run(
    'pnpm',
    ['install', '--ignore-scripts', '--no-frozen-lockfile'],
    temporaryRoot,
  )
  run('node', ['smoke.mjs'], temporaryRoot)
  run(
    join(repositoryRoot, 'node_modules/.bin/tsc'),
    [
      '--module',
      'NodeNext',
      '--moduleResolution',
      'NodeNext',
      '--noEmit',
      '--skipLibCheck',
      '--strict',
      '--target',
      'ES2022',
      'smoke.ts',
    ],
    temporaryRoot,
  )

  succeeded = true
  console.log(
    'Packed docx-kit passed runtime import and TypeScript consumer checks.',
  )
} finally {
  if (succeeded || process.env.DOCX_KIT_KEEP_SMOKE_TEMP !== '1') {
    rmSync(temporaryRoot, { force: true, recursive: true })
  } else {
    console.error(`Smoke-test files retained at ${temporaryRoot}`)
  }
}
