import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import process from 'node:process'

const documentationRoots = [
  'README.md',
  'docs',
  'packages',
  'packages-plugins',
  'packages-presets',
  'packages-themes',
]

const pluginFactories = [
  'badgePlugin',
  'calloutPlugin',
  'changelogPlugin',
  'codeBlockPlugin',
  'coverPagePlugin',
  'dataTablePlugin',
  'dividerPlugin',
  'echartsPlugin',
  'invoicePlugin',
  'letterheadPlugin',
  'meetingMinutesPlugin',
  'pageNumberPlugin',
  'propertyTablePlugin',
  'qrcodePlugin',
  'signatureBlockPlugin',
  'timelinePlugin',
  'tocPlugin',
  'watermarkPlugin',
]

function collectMarkdownFiles(path) {
  const stats = statSync(path)
  if (stats.isFile()) {
    return path.endsWith('.md') ? [path] : []
  }

  return readdirSync(path, { withFileTypes: true }).flatMap(entry => {
    if (entry.name === 'dist' || entry.name === 'node_modules') {
      return []
    }
    return collectMarkdownFiles(join(path, entry.name))
  })
}

function getLineNumber(source, index) {
  return source.slice(0, index).split('\n').length
}

const errors = []
const markdownFiles = documentationRoots.flatMap(collectMarkdownFiles)

for (const file of markdownFiles) {
  const source = readFileSync(file, 'utf8')

  if (source.includes('.save(') && !source.includes("from 'docx-kit/node'")) {
    errors.push(`${file}: Node.js .save() example has no docx-kit/node import`)
  }

  for (const factory of pluginFactories) {
    const missingFactoryCall = `.use(${factory})`
    if (source.includes(missingFactoryCall)) {
      errors.push(
        `${file}: call the ${factory} factory before passing it to .use()`,
      )
    }
  }

  for (const match of source.matchAll(/```(?:ts|typescript)\n([\s\S]*?)```/g)) {
    const block = match[1]
    const line = getLineNumber(source, match.index)
    const importsBrowserEntry = block.split('\n').some(sourceLine => {
      const trimmedLine = sourceLine.trim()
      const normalizedLine = trimmedLine.endsWith(';')
        ? trimmedLine.slice(0, -1)
        : trimmedLine
      const isImportLine =
        normalizedLine.startsWith('import ')
        || normalizedLine.startsWith('} from ')
      return isImportLine && normalizedLine.endsWith("from 'docx-kit'")
    })

    if (block.includes('.save(') && importsBrowserEntry) {
      errors.push(
        `${file}:${line}: .save() is unavailable from the browser entry`,
      )
    }

    if (block.includes('.save(') && block.includes('echartsPlugin')) {
      errors.push(
        `${file}:${line}: ECharts requires browser Blob export, not Node.js .save()`,
      )
    }

    if (/const\s+\w+\s*=\s*renderDocx\(/u.test(block)) {
      errors.push(
        `${file}:${line}: renderDocx() must be awaited before using its builder`,
      )
    }

    for (const factory of pluginFactories) {
      const uncalledInlinePlugin = new RegExp(
        `plugin:\\s*${factory}(?!\\s*\\()`,
        'u',
      )
      if (uncalledInlinePlugin.test(block)) {
        errors.push(
          `${file}:${line}: call the ${factory} factory in inline plugin sources`,
        )
      }
    }
  }
}

if (errors.length > 0) {
  console.error(errors.join('\n'))
  process.exitCode = 1
} else {
  console.log(`Validated ${markdownFiles.length} documentation files.`)
}
