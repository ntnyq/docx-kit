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

const PIXEL_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=',
  ),
  character => character.codePointAt(0) ?? 0,
)

/**
 * Repeatable document workloads and their content counts.
 */
export const PERFORMANCE_SCENARIOS: PerformanceScenario[] = [
  { count: 100, kind: 'paragraphs', name: 'paragraphs-100' },
  { count: 1_000, kind: 'paragraphs', name: 'paragraphs-1000' },
  { count: 10_000, kind: 'paragraphs', name: 'paragraphs-10000' },
  { count: 1_000, kind: 'table', name: 'table-1000-rows' },
  { count: 100, kind: 'images', name: 'images-100' },
  { count: 500, kind: 'plugins', name: 'plugins-500' },
]

/**
 * Build a named workload with the supplied platform runtime.
 *
 * @throws If the scenario name or workload kind is unknown
 */
export function buildPerformanceScenario(
  runtime: PerformanceRuntime,
  scenarioName: string,
): Pick<ReturnType<typeof createDocx>, 'toBlob' | 'toBuffer'> {
  const scenario = PERFORMANCE_SCENARIOS.find(
    item => item.name === scenarioName,
  )
  if (!scenario) {
    throw new Error(`Unknown performance scenario: ${scenarioName}`)
  }

  switch (scenario.kind) {
    case 'images':
      return buildImageDocument(runtime, scenario.count)
    case 'paragraphs':
      return buildParagraphDocument(runtime, scenario.count)
    case 'plugins':
      return buildPluginDocument(runtime, scenario.count)
    case 'table':
      return buildTableDocument(runtime, scenario.count)
    default:
      throw new Error(`Unsupported performance scenario: ${scenario.kind}`)
  }
}

function buildImageDocument({ createDocx }: PerformanceRuntime, count: number) {
  const document = createDocx().h1(`Image benchmark: ${count} images`)

  for (let index = 0; index < count; index += 1) {
    document.image({
      alt: `Benchmark pixel ${index + 1}`,
      data: PIXEL_PNG,
      height: 32,
      imageType: 'png',
      width: 32,
    })
  }

  return document
}

function buildParagraphDocument(
  { createDocx }: PerformanceRuntime,
  count: number,
) {
  const document = createDocx().h1(`Paragraph benchmark: ${count} paragraphs`)

  for (let index = 0; index < count; index += 1) {
    document.p(
      `Paragraph ${index + 1}: docx-kit performance baseline content with bold and colored spans.`,
      {
        style: {
          color: index % 2 === 0 ? '#1F2937' : '#374151',
          marginBottom: '4pt',
        },
      },
    )
  }

  return document
}

function buildPluginDocument(
  { calloutPlugin, createDocx }: PerformanceRuntime,
  count: number,
) {
  const document = createDocx()
    .use(calloutPlugin())
    .h1(`Plugin benchmark: ${count} callouts`)
  const types = ['danger', 'info', 'success', 'warning'] as const

  for (let index = 0; index < count; index += 1) {
    document.plugin('callout', {
      content: `Plugin-rendered content ${index + 1}`,
      title: `Callout ${index + 1}`,
      type: types[index % types.length],
    })
  }

  return document
}

function buildTableDocument({ createDocx }: PerformanceRuntime, count: number) {
  const rows = Array.from({ length: count }, (_, index) => ({
    amount: (index + 1) * 17.25,
    category: `Category ${(index % 12) + 1}`,
    id: `ROW-${String(index + 1).padStart(5, '0')}`,
    owner: `Owner ${(index % 40) + 1}`,
    status: index % 7 === 0 ? 'At risk' : 'On track',
  }))

  return createDocx()
    .h1(`Table benchmark: ${count} rows`)
    .table({
      data: rows,
      header: true,
      striped: true,
      columns: [
        { key: 'id', title: 'ID' },
        { key: 'category', title: 'Category' },
        { key: 'owner', title: 'Owner' },
        { key: 'status', title: 'Status' },
        {
          key: 'amount',
          title: 'Amount',
          render: value => Number(value).toFixed(2),
        },
      ],
    })
}
