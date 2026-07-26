import { performance } from 'node:perf_hooks'
import process from 'node:process'
// eslint-disable-next-line antfu/no-import-dist -- the benchmark intentionally measures the built package
import { calloutPlugin, createDocx } from '../../packages/docx-kit/dist/node.js'
import { buildPerformanceScenario } from './scenarios.mjs'

const scenarioName = process.argv[2]

if (!scenarioName) {
  throw new Error('A performance scenario name is required')
}

globalThis.gc?.()

const baselineRssBytes = process.memoryUsage().rss
const baselinePeakRssBytes = process.resourceUsage().maxRSS * 1024
const startedAt = performance.now()
const document = buildPerformanceScenario(
  { calloutPlugin, createDocx },
  scenarioName,
)
const output = await document.toBuffer()
const durationMs = performance.now() - startedAt
const peakRssBytes = process.resourceUsage().maxRSS * 1024

process.stdout.write(
  `${JSON.stringify({
    baselineRssBytes,
    durationMs,
    outputBytes: output.byteLength,
    peakRssBytes,
    peakRssDeltaBytes: Math.max(0, peakRssBytes - baselinePeakRssBytes),
    scenario: scenarioName,
  })}\n`,
)
