/**
 * Scaffold command — creates a new docx-kit plugin project directory.
 *
 * @module commands/create-plugin
 */

import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import prompts from 'prompts'
import {
  renderEslintConfig,
  renderManifest,
  renderPackageJson,
  renderPluginSource,
  renderPluginTest,
  renderReadme,
  renderTsconfigJson,
  renderTsdownConfig,
} from '../templates/plugin/index'

/**
 * Resolved scaffold options.
 */
interface ScaffoldOptions {
  author: string
  description: string
  directory: string
  license: string
  pluginName: string
  shortName: string
  version: string
}

/**
 * Run the scaffold command.
 *
 * Creates a new plugin project directory with all necessary
 * files: manifest, source, test, package.json, tsconfig, and README.
 *
 * @param nameArg - — Optional plugin name from CLI argument
 */
export async function createPlugin(nameArg?: string): Promise<void> {
  const options = await promptOptions(nameArg)

  if (!options) {
    console.log('Scaffold cancelled.')
    return
  }

  const targetDir = path.resolve(process.cwd(), options.directory)

  // Check if directory already exists
  if (await pathExists(targetDir)) {
    console.error(
      `Directory "${options.directory}" already exists. Choose a different name.`,
    )
    return
  }

  console.log(`Creating plugin project: ${options.pluginName}`)
  console.log(`  Directory: ${targetDir}`)

  await writeScaffold(targetDir, options)

  console.log('')
  console.log('  Plugin project created successfully!')
  console.log('')
  console.log('  Next steps:')
  console.log(`    cd ${options.directory}`)
  console.log('    pnpm install')
  console.log('    pnpm run test')
  console.log('')
  console.log('  Happy plugin development!')
}

/**
 * Extract the short plugin name from a package name.
 *
 * Removes `docx-kit-plugin-` prefix if present.
 */
function deriveShortName(packageName: string): string {
  const unscopedName = packageName.split('/').at(-1) ?? packageName
  const prefix = 'docx-kit-plugin-'
  if (unscopedName.startsWith(prefix)) {
    return unscopedName.slice(prefix.length)
  }
  return unscopedName
}

/**
 * Check if a path exists on the filesystem.
 */
async function pathExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath)
    return true
  } catch {
    return false
  }
}

/**
 * Interactive prompts for scaffold options.
 *
 * @param initialName - — Optional name from CLI argument
 */
async function promptOptions(
  initialName?: string,
): Promise<ScaffoldOptions | null> {
  const namePrompt = initialName ?? 'my-plugin'

  const result = await prompts([
    {
      initial: namePrompt,
      message: 'Plugin package name:',
      name: 'pluginName',
      type: 'text',
    },
    {
      initial: 'A docx-kit plugin',
      message: 'Description:',
      name: 'description',
      type: 'text',
    },
    {
      message: 'Author:',
      name: 'author',
      type: 'text',
    },
    {
      initial: 'MIT',
      message: 'License:',
      name: 'license',
      type: 'text',
    },
    {
      initial: '0.1.0',
      message: 'Initial version:',
      name: 'version',
      type: 'text',
    },
  ])

  // User cancelled (Ctrl+C) — prompts returns null or undefined
  if (!result || !result.pluginName) {
    return null
  }

  return {
    author: result.author ?? '',
    description: result.description ?? 'A docx-kit plugin',
    directory: result.pluginName,
    license: result.license ?? 'MIT',
    pluginName: result.pluginName,
    shortName: deriveShortName(result.pluginName),
    version: result.version ?? '0.1.0',
  }
}

/**
 * Write all scaffold files to the target directory.
 */
async function writeScaffold(
  targetDir: string,
  options: ScaffoldOptions,
): Promise<void> {
  const srcDir = path.join(targetDir, 'src')
  const testDir = path.join(targetDir, 'tests')

  await fs.mkdir(srcDir, { recursive: true })
  await fs.mkdir(testDir, { recursive: true })

  await fs.writeFile(
    path.join(targetDir, 'docx-kit.plugin.json'),
    renderManifest(
      options.pluginName,
      options.version,
      '^0.3.0',
      options.description,
    ),
  )

  await fs.writeFile(
    path.join(srcDir, 'index.ts'),
    renderPluginSource(options.shortName),
  )

  await fs.writeFile(
    path.join(testDir, 'index.test.ts'),
    renderPluginTest(options.shortName),
  )

  await fs.writeFile(
    path.join(targetDir, 'package.json'),
    renderPackageJson(
      options.pluginName,
      options.shortName,
      options.description,
      options.author,
      options.license,
      options.version,
    ),
  )

  await fs.writeFile(
    path.join(targetDir, 'tsconfig.json'),
    renderTsconfigJson(),
  )

  await fs.writeFile(
    path.join(targetDir, 'tsdown.config.ts'),
    renderTsdownConfig(),
  )

  await fs.writeFile(
    path.join(targetDir, 'eslint.config.mjs'),
    renderEslintConfig(),
  )

  await fs.writeFile(
    path.join(targetDir, 'README.md'),
    renderReadme(
      options.pluginName,
      options.shortName,
      options.description,
      options.author,
    ),
  )
}
