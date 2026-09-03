import type {
  calloutPlugin,
  createDocx,
} from '../../packages/docx-kit/src/browser'

/**
 * Platform-specific document and plugin factories used to build a workload.
 */
export interface PerformanceRuntime {
  calloutPlugin: typeof calloutPlugin
  createDocx: typeof createDocx
}

/**
 * Workload configuration shared by the Node.js and browser benchmarks.
 */
export interface PerformanceScenario {
  count: number
  kind: 'images' | 'paragraphs' | 'plugins' | 'table'
  name: string
}

/**
 * Repeatable document workloads and their content counts.
 */
export const PERFORMANCE_SCENARIOS: PerformanceScenario[]

/**
 * Build a named workload with the supplied platform runtime.
 *
 * @param runtime - Document and callout factories for the target platform
 * @param scenarioName - A name from PERFORMANCE_SCENARIOS
 * @returns The document's cross-platform binary export methods
 * @throws If the scenario name or workload kind is unknown
 */
export function buildPerformanceScenario(
  runtime: PerformanceRuntime,
  scenarioName: string,
): Pick<ReturnType<typeof createDocx>, 'toBlob' | 'toBuffer'>
