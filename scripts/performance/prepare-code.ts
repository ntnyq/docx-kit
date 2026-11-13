import {
  BASIC_CODE,
  COMPREHENSIVE_CODE,
} from '../../apps/playground/src/constants/templates'

interface CompilerMeasurement {
  iterations: number
  medianMs: number
  name: string
  outputLength: number
  samplesMs: number[]
  sourceBytes: number
  warmups: number
}

interface CompilerScenario {
  name: string
  source: string
}

export const PREPARE_CODE_SCENARIOS: readonly CompilerScenario[] = [
  { name: 'basic', source: BASIC_CODE },
  { name: 'comprehensive', source: COMPREHENSIVE_CODE },
  { name: 'typed-100', source: createTypedSource(100) },
  { name: 'typed-1000', source: createTypedSource(1000) },
]

/**
 * Measure complete source preparation without executing DOCX generation.
 * Each sample is the mean duration of 100 conversions, in milliseconds.
 */
export function benchmarkPrepareCode(
  prepareCode: (source: string) => string,
): CompilerMeasurement[] {
  const warmups = 100
  const iterations = 100
  return PREPARE_CODE_SCENARIOS.map(({ name, source }) => {
    let outputLength = 0
    for (let index = 0; index < warmups; index += 1) {
      outputLength = prepareCode(source).length
    }
    const samplesMs = Array.from({ length: 9 }, () => {
      const startedAt = performance.now()
      for (let index = 0; index < iterations; index += 1) {
        outputLength = prepareCode(source).length
      }
      return (performance.now() - startedAt) / iterations
    })
    const sorted = samplesMs.toSorted((left, right) => left - right)
    return {
      iterations,
      medianMs: sorted[Math.floor(sorted.length / 2)],
      name,
      outputLength,
      samplesMs,
      sourceBytes: new TextEncoder().encode(source).length,
      warmups,
    }
  })
}

function createTypedSource(count: number): string {
  return [
    "import { createDocx, type DocxKitConfig } from 'docx-kit'",
    'const config: DocxKitConfig = {}',
    ...Array.from(
      { length: count },
      (_, index) =>
        `const row${index}: { title: string; value: number } = { title: 'Row ${index}', value: ${index} }`,
    ),
    `createDocx(config).p(row${count - 1}.title).toBlob()`,
  ].join('\n')
}
