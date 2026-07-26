/**
 * CLI entry point for `@docxkit/create-plugin`.
 *
 * Run via `npx @docxkit/create-plugin <name>` or
 * `create-docx-kit-plugin <name>` (if installed globally).
 *
 * @module create-plugin
 */

import process from 'node:process'
import { createPlugin } from './commands/create-plugin'

const nameArg = process.argv[2]

createPlugin(nameArg).catch((error: unknown) => {
  console.error('Failed to scaffold plugin project:', error)
  process.exit(1)
})
