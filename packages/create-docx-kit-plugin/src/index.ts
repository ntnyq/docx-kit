/**
 * CLI entry point for `create-docx-kit-plugin`.
 *
 * Run via `npx create-docx-kit-plugin <name>` or
 * `create-docx-kit-plugin <name>` (if installed globally).
 *
 * @module create-docx-kit-plugin
 */

import process from 'node:process'
import { createPlugin } from './commands/create-plugin'

const nameArg = process.argv[2]

createPlugin(nameArg).catch((err: unknown) => {
  console.error('Failed to scaffold plugin project:', err)
  process.exit(1)
})
