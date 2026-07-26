import { calloutPlugin, createDocx } from '../../packages/docx-kit/src/browser'
import {
  buildPerformanceScenario,
  PERFORMANCE_SCENARIOS,
} from '../../scripts/performance/scenarios.mjs'

interface ChromiumPerformanceMemory {
  usedJSHeapSize: number
}

interface PerformanceWithMemory extends Performance {
  memory?: ChromiumPerformanceMemory
}

const resultElement =
  document.querySelector<HTMLPreElement>('#benchmark-result')!
const statusElement = document.querySelector<HTMLParagraphElement>('#status')!
const performanceWithMemory = performance as PerformanceWithMemory
const results = []
const runsPerScenario = Math.max(
  1,
  Number(new URLSearchParams(location.search).get('runs') ?? 3),
)

for (const scenario of PERFORMANCE_SCENARIOS) {
  const samples = []

  for (let run = 0; run < runsPerScenario; run += 1) {
    const heapBeforeBytes = performanceWithMemory.memory?.usedJSHeapSize
    const startedAt = performance.now()
    const documentBuilder = buildPerformanceScenario(
      { calloutPlugin, createDocx },
      scenario.name,
    )
    const output = await documentBuilder.toBlob()

    samples.push({
      durationMs: performance.now() - startedAt,
      outputBytes: output.size,
      heapDeltaBytes:
        heapBeforeBytes == null
          ? undefined
          : Math.max(
              0,
              (performanceWithMemory.memory?.usedJSHeapSize ?? heapBeforeBytes)
                - heapBeforeBytes,
            ),
    })
  }

  results.push({
    durationMedianMs: median(samples.map(sample => sample.durationMs)),
    durationSamplesMs: samples.map(sample => sample.durationMs),
    outputMedianBytes: median(samples.map(sample => sample.outputBytes)),
    scenario: scenario.name,
    heapDeltaMaxBytes: Math.max(
      ...samples.map(sample => sample.heapDeltaBytes ?? 0),
    ),
  })
}

const report = {
  generatedAt: new Date().toISOString(),
  runsPerScenario,
  scenarios: results,
  userAgent: navigator.userAgent,
}

resultElement.dataset.report = encodeURIComponent(JSON.stringify(report))
resultElement.textContent = JSON.stringify(report, null, 2)
statusElement.textContent = 'Complete'
document.title = 'docx-kit browser benchmark: complete'

function median(values: number[]) {
  const sorted = [...values].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle]
}
