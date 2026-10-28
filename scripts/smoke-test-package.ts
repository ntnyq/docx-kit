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
import { pathToFileURL } from 'node:url'

const repositoryRoot = resolve(import.meta.dirname, '..')
const packageRoots = [
  'packages',
  'packages-plugins',
  'packages-presets',
  'packages-themes',
]
const temporaryRoot = mkdtempSync(join(tmpdir(), 'docx-kit-package-smoke-'))
const tarballDirectory = join(temporaryRoot, 'tarballs')

interface PackageManifest {
  name: string
  bin?: Record<string, string>
  peerDependencies?: Record<string, string>
  private?: boolean
}

function checkInstalledCli() {
  const packageDirectory = join(
    temporaryRoot,
    'node_modules/@docxkit/create-plugin',
  )
  const command = 'create-docx-kit-plugin'
  const entry = readPackageManifest(packageDirectory).bin?.[command]
  if (!entry) {
    throw new Error(`Installed CLI does not expose ${command}`)
  }
  if (
    !readFileSync(join(packageDirectory, entry), 'utf8').startsWith(
      '#!/usr/bin/env node',
    )
  ) {
    throw new Error('Installed CLI entry is missing its Node.js shebang')
  }

  const projectName = 'docx-kit-plugin-packed-consumer'
  const preloadPath = join(temporaryRoot, 'cli-prompts.mjs')
  writeFileSync(
    preloadPath,
    `import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const packageRequire = createRequire(require.resolve('@docxkit/create-plugin/package.json'))
packageRequire('prompts').inject(${JSON.stringify([
      projectName,
      'Packed CLI consumer',
      'docx-kit',
      'MIT',
      '0.1.0',
    ])})
`,
  )

  execFileSync('pnpm', ['exec', command, projectName], {
    cwd: temporaryRoot,
    shell: process.platform === 'win32',
    stdio: 'inherit',
    timeout: 30_000,
    env: {
      ...process.env,
      NODE_OPTIONS: [
        process.env.NODE_OPTIONS,
        `--import=${pathToFileURL(preloadPath).href}`,
      ]
        .filter(Boolean)
        .join(' '),
    },
  })

  const projectDirectory = join(temporaryRoot, projectName)
  for (const file of [
    'package.json',
    'docx-kit.plugin.json',
    'src/index.ts',
    'tests/index.test.ts',
    'tsconfig.json',
    'tsdown.config.ts',
    'eslint.config.mjs',
    'README.md',
  ]) {
    if (!existsSync(join(projectDirectory, file))) {
      throw new Error(`Installed CLI did not generate ${file}`)
    }
  }
  if (readPackageManifest(projectDirectory).name !== projectName) {
    throw new Error('Installed CLI generated an incorrect package name')
  }
}

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
        manifest: readPackageManifest(packageDirectory),
      }))
      .filter(pkg => !pkg.manifest.private)
  })
}

function readPackageManifest(directory: string): PackageManifest {
  return JSON.parse(readFileSync(join(directory, 'package.json'), 'utf8'))
}

function run(command: string, args: string[], cwd: string) {
  execFileSync(command, args, {
    cwd,
    env: process.env,
    stdio: 'inherit',
  })
}

function runSilently(command: string, args: string[], cwd: string) {
  execFileSync(command, args, {
    cwd,
    env: process.env,
    stdio: 'pipe',
  })
}

let succeeded = false

try {
  const packages = collectPackages()
  const overrides: Record<string, string> = {}

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
          '@docxkit/core': overrides['@docxkit/core'],
          '@docxkit/create-plugin': overrides['@docxkit/create-plugin'],
          '@docxkit/errors': overrides['@docxkit/errors'],
          '@docxkit/types': overrides['@docxkit/types'],
          'docx-kit': overrides['docx-kit'],
        },
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(
    join(temporaryRoot, 'pnpm-workspace.yaml'),
    [
      'packages: []',
      'overrides:',
      ...Object.entries(overrides).map(
        ([name, tarball]) =>
          `  ${JSON.stringify(name)}: ${JSON.stringify(tarball)}`,
      ),
      '',
    ].join('\n'),
  )
  writeFileSync(
    join(temporaryRoot, 'smoke.mjs'),
    `import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

const expectedExports = [
  ['@docxkit/core', 'definePlugin'],
  ['@docxkit/core', 'defineStyles'],
  ['@docxkit/core', 'inlineImg'],
  ['@docxkit/core', 'span'],
  ['@docxkit/errors', 'DocxKitError'],
  ['@docxkit/errors', 'ERROR_CODES'],
  ['docx-kit', 'defineStyles'],
  ['docx-kit', 'createDocx'],
  ['docx-kit/ai', 'buildPrompt'],
  ['docx-kit/browser', 'createDocx'],
  ['docx-kit/loader', 'createPluginLoader'],
  ['docx-kit/loader/browser', 'loadUrlPlugin'],
  ['docx-kit/loader/node', 'loadNpmPlugin'],
  ['docx-kit/markdown', 'markdownToNodes'],
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

const { DocxKitError } = await import('@docxkit/errors')
for (const specifier of ['@docxkit/core', 'docx-kit', 'docx-kit/browser']) {
  const module = await import(specifier)
  if (module.DocxKitError !== DocxKitError) {
    throw new Error(\`Duplicated error class in \${specifier}\`)
  }
}

try {
  await import('@docxkit/types')
  throw new Error('@docxkit/types must not have a runtime entry point')
} catch (error) {
  if (error.code !== 'ERR_PACKAGE_PATH_NOT_EXPORTED') throw error
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
import { createDocx, definePlugin, defineStyles, inlineImg, span } from 'docx-kit'
import { DocxKitError, ERROR_CODES, type ErrorCode } from '@docxkit/errors'
import { loadUrlPlugin } from 'docx-kit/loader/browser'
import { loadNpmPlugin } from 'docx-kit/loader/node'
import { markdownToNodes, type MarkdownOptions } from 'docx-kit/markdown'
import { createDocxKitServer, TOOL_DEFINITIONS } from 'docx-kit/mcp'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
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

const styles = defineStyles({ body: { fontSize: '12pt' } })
const text = span('Runtime helpers', { bold: true })
const plugin = definePlugin({ name: 'consumer', render: () => [text] })
const code: ErrorCode = ERROR_CODES.EXPORT_FAILED
void [styles, plugin, inlineImg, new DocxKitError(code, 'Consumer error')]

void createNodeDocx().save('consumer-smoke.docx')
const markdownOptions: MarkdownOptions = { gfm: true }
for (const node of markdownToNodes('# Markdown consumer', markdownOptions)) {
  createDocx().add(node)
}

const server = await createDocxKitServer()
const [transport] = InMemoryTransport.createLinkedPair()
await server.connect(transport)
await server.close()
`,
  )
  writeFileSync(
    join(temporaryRoot, 'types-smoke.ts'),
    `import type {
  DocxKitConfig,
  DocxStyleRule,
  ParagraphNode,
  StyleSheet,
  TextNode,
} from '@docxkit/types'

const styles = { body: { fontSize: '12pt' } } satisfies StyleSheet
const config: DocxKitConfig<typeof styles> = { styles }
const text: TextNode = { type: 'text', text: 'Strict NodeNext consumer' }
const paragraph: ParagraphNode = { type: 'paragraph', children: [text] }

// @ts-expect-error bold only accepts a boolean
const invalidStyle: DocxStyleRule = { bold: 'yes' }

// @ts-expect-error runtime helpers belong to @docxkit/core
import type { defineStyles } from '@docxkit/types'
// @ts-expect-error errors belong to @docxkit/errors
import type { DocxKitError } from '@docxkit/types'

void [config, invalidStyle, paragraph]
`,
  )

  run(
    'pnpm',
    ['install', '--ignore-scripts', '--no-frozen-lockfile'],
    temporaryRoot,
  )
  checkInstalledCli()
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
      'false',
      '--strict',
      '--target',
      'ES2022',
      'types-smoke.ts',
    ],
    temporaryRoot,
  )
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
    'Packed packages passed installed CLI, runtime imports, strict NodeNext types, and TypeScript consumer checks.',
  )
} finally {
  if (succeeded || process.env.DOCX_KIT_KEEP_SMOKE_TEMP !== '1') {
    rmSync(temporaryRoot, { force: true, recursive: true })
  } else {
    console.error(`Smoke-test files retained at ${temporaryRoot}`)
  }
}
