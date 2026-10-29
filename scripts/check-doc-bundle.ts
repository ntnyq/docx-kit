import { readdirSync, statSync } from 'node:fs'
import { resolve } from 'node:path'

const assets = resolve(import.meta.dirname, '../docs/.vitepress/dist/assets')
const files = readdirSync(assets).filter(name => name.endsWith('.js'))
for (const name of files) {
  if (
    /monaco|\.worker-|playground|theme-studio|^(?:generate-error-docs|nostics|catalog)[.-]/i.test(
      name,
    )
  ) {
    throw new Error(
      `Application code leaked into the documentation bundle: ${name}`,
    )
  }
  if (statSync(resolve(assets, name)).size > 3 * 1024 * 1024) {
    throw new Error(`Documentation chunk exceeds 3 MiB: ${name}`)
  }
}
console.log(
  'Documentation bundle audit passed; standalone app code is excluded.',
)
