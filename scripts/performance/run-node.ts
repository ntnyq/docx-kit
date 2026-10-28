import { execFile } from 'node:child_process'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { arch, cpus, platform, release } from 'node:os'
import { dirname, resolve } from 'node:path'
import process from 'node:process'
import { promisify } from 'node:util'
import { PERFORMANCE_SCENARIOS } from './scenarios'

interface PerformanceReport {
  scenarios: ScenarioSummary[]
}

interface PerformanceSample {
  durationMs: number
  outputBytes: number
  peakRssBytes: number
  peakRssDeltaBytes: number
}

type ScenarioSummary = ReturnType<typeof summarizeSamples>

const execFileAsync = promisify(execFile)
const repositoryRoot = resolve(import.meta.dirname, '../..')
const defaultOutputPath = resolve(
  repositoryRoot,
  'output/performance/node-latest.json',
)
const baselinePath = resolve(repositoryRoot, 'benchmarks/baselines/node.json')
const workerPath = resolve(import.meta.dirname, 'node-worker.ts')
const updateBaseline = process.argv.includes('--update-baseline')
const runsArgumentIndex = process.argv.indexOf('--runs')
const runs = Math.max(
  1,
  Number(
    runsArgumentIndex === -1
      ? (process.env.DOCX_KIT_BENCH_RUNS ?? 3)
      : process.argv[runsArgumentIndex + 1],
  ),
)
const outputArgumentIndex = process.argv.indexOf('--output')
const outputPath =
  outputArgumentIndex === -1
    ? defaultOutputPath
    : resolve(repositoryRoot, process.argv[outputArgumentIndex + 1])

const results = []

for (const scenario of PERFORMANCE_SCENARIOS) {
  const samples: PerformanceSample[] = []

  for (let run = 0; run < runs; run += 1) {
    const { stdout } = await execFileAsync(
      process.execPath,
      ['--expose-gc', '--import', 'tsx', workerPath, scenario.name],
      {
        cwd: repositoryRoot,
        maxBuffer: 1024 * 1024,
      },
    )
    samples.push(JSON.parse(stdout.trim()))
  }

  results.push(summarizeSamples(scenario.name, samples))
}

const report = {
  generatedAt: new Date().toISOString(),
  runsPerScenario: runs,
  scenarios: results,
  machine: {
    arch: arch(),
    cpu: cpus()[0]?.model ?? 'unknown',
    node: process.version,
    platform: platform(),
    release: release(),
  },
}

await writeJson(outputPath, report)
if (updateBaseline) {
  await writeJson(baselinePath, report)
}

printResults(report)
await printComparison(report)
console.log(`Report: ${outputPath}`)
if (updateBaseline) {
  console.log(`Baseline updated: ${baselinePath}`)
}

function formatBytes(bytes: number) {
  const units = ['B', 'KiB', 'MiB', 'GiB']
  let value = bytes
  let unitIndex = 0
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024
    unitIndex += 1
  }
  return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`
}

function formatDelta(current: number, previous: number | undefined) {
  if (!previous) {
    return 'new'
  }
  const percentage = ((current - previous) / previous) * 100
  return `${percentage >= 0 ? '+' : ''}${percentage.toFixed(1)}%`
}

function median(values: number[]) {
  const sorted = [...values].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle]
}

async function printComparison(current: PerformanceReport) {
  let baseline: PerformanceReport
  try {
    baseline = JSON.parse(await readFile(baselinePath, 'utf8'))
  } catch {
    return
  }

  console.log('\nChange from recorded baseline (informational):')
  console.table(
    current.scenarios.map(scenario => {
      const previous = baseline.scenarios.find(
        item => item.scenario === scenario.scenario,
      )
      return {
        scenario: scenario.scenario,
        duration: formatDelta(
          scenario.durationMedianMs,
          previous?.durationMedianMs,
        ),
        output: formatDelta(
          scenario.outputMedianBytes,
          previous?.outputMedianBytes,
        ),
        peakRSS: formatDelta(
          scenario.peakRssMaxBytes,
          previous?.peakRssMaxBytes,
        ),
      }
    }),
  )
}

function printResults(report: PerformanceReport) {
  console.table(
    report.scenarios.map(scenario => ({
      duration: `${scenario.durationMedianMs.toFixed(1)} ms`,
      output: formatBytes(scenario.outputMedianBytes),
      peakRSS: formatBytes(scenario.peakRssMaxBytes),
      peakRSSDelta: formatBytes(scenario.peakRssDeltaMaxBytes),
      scenario: scenario.scenario,
    })),
  )
}

function summarizeSamples(scenario: string, samples: PerformanceSample[]) {
  return {
    durationMedianMs: median(samples.map(sample => sample.durationMs)),
    durationSamplesMs: samples.map(sample => sample.durationMs),
    outputMedianBytes: median(samples.map(sample => sample.outputBytes)),
    peakRssMaxBytes: Math.max(...samples.map(sample => sample.peakRssBytes)),
    scenario,
    peakRssDeltaMaxBytes: Math.max(
      ...samples.map(sample => sample.peakRssDeltaBytes),
    ),
  }
}

async function writeJson(path: string, value: unknown) {
  await mkdir(dirname(path), { recursive: true })
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`)
}
