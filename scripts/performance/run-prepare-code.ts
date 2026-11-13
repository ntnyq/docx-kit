import { mkdir, writeFile } from 'node:fs/promises'
import { cpus } from 'node:os'
import { dirname, resolve } from 'node:path'
import process from 'node:process'
import { pathToFileURL } from 'node:url'
import { benchmarkPrepareCode, PREPARE_CODE_SCENARIOS } from './prepare-code'

const moduleIndex = process.argv.indexOf('--module')
const outputIndex = process.argv.indexOf('--output')
const modulePath = resolve(
  moduleIndex === -1
    ? 'apps/playground/src/utils/prepareCode.ts'
    : process.argv[moduleIndex + 1],
)
const outputPath = resolve(
  outputIndex === -1
    ? 'output/performance/prepare-code.json'
    : process.argv[outputIndex + 1],
)
const startedAt = performance.now()
const runtime: unknown = await import(pathToFileURL(modulePath).href)
const importMs = performance.now() - startedAt
if (
  !runtime
  || typeof runtime !== 'object'
  || !('prepareCode' in runtime)
  || typeof runtime.prepareCode !== 'function'
) {
  throw new TypeError('Benchmark module must export prepareCode(source)')
}
const prepare = runtime.prepareCode
const prepareCode = (source: string): string => {
  const result: unknown = prepare(source)
  if (typeof result !== 'string') {
    throw new TypeError('prepareCode must return a string')
  }
  return result
}
const firstStartedAt = performance.now()
prepareCode(PREPARE_CODE_SCENARIOS[0].source)
const firstConversionMs = performance.now() - firstStartedAt
const report = {
  cpu: cpus()[0]?.model,
  firstConversionMs,
  importMs,
  node: process.version,
  scenarios: benchmarkPrepareCode(prepareCode),
}
await mkdir(dirname(outputPath), { recursive: true })
await writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`)
console.table(
  report.scenarios.map(({ medianMs, name, sourceBytes }) => ({
    milliseconds: medianMs.toFixed(3),
    scenario: name,
    sourceBytes,
  })),
)
console.log(`Report: ${outputPath}`)
