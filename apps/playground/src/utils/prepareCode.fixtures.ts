import { PRESETS } from '../constants/templates'

interface CompilerFixture {
  name: string
  source: string
}

/**
 * Exercise the production-only Sucrase profile against its normal public API.
 */
export const COMPILER_FIXTURES: CompilerFixture[] = [
  ...PRESETS.map(preset => ({
    name: preset.id,
    source: preset.code,
  })),
  {
    name: 'numeric and string enums',
    source:
      'enum A { One = 1, Two }; enum B { Ready = "ready" }; [A.Two, B.Ready]',
  },
  {
    name: 'constructor properties and class fields',
    source:
      'class Value { doubled = this.value * 2; constructor(public value: number) {} }; new Value(21).doubled',
  },
  {
    name: 'async generic arrow',
    source:
      'const fn = async <\nT\n>(value: T): Promise<T> => value; await fn<number>(42)',
  },
  {
    name: 'optional chaining and numeric separators',
    source: 'const value: { count?: number } = {}; value?.count ?? 1_000',
  },
  {
    name: 'type-only and aliased imports',
    source:
      'import { createDocx as make, type DocxKitConfig } from "docx-kit"; const config: DocxKitConfig = {}; make(config).toBlob()',
  },
  {
    name: 'UTF-16 and quoted import names',
    source:
      'const label = "文档📄"; import { "HeadingLevel" as headings } from "docx"; [label, headings.HEADING_1]',
  },
  {
    name: 'dynamic import and explicit return',
    source:
      'return (await import("data:text/javascript,export default 42")).default',
  },
  { name: 'trailing erased import', source: '42; import "external-package"' },
  {
    name: 'trailing runtime import',
    source: '42; import { createDocx } from "docx-kit"',
  },
]
