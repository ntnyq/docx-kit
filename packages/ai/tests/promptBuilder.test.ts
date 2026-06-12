/**
 * Tests for AI prompt builder and schema guide.
 *
 * @module tests/ai/promptBuilder
 */

import { describe, expect, it } from 'vitest'
import { buildFreeformPrompt, buildPrompt } from '../src/promptBuilder'
import {
  generateSchemaGuide,
  generateToolDefinitions,
} from '../src/schemaGuide'
import { reportTemplate } from '../src/templates'

describe('buildPrompt', () => {
  it('includes system prompt from template', () => {
    const prompt = buildPrompt(reportTemplate as any)
    expect(prompt).toContain(reportTemplate.systemPrompt)
  })

  it('includes template schema', () => {
    const prompt = buildPrompt(reportTemplate as any)
    expect(prompt).toContain('Template: report')
    expect(prompt).toContain('Input Schema')
  })

  it('includes user context when provided', () => {
    const prompt = buildPrompt(
      reportTemplate as any,
      'Focus on quarterly metrics',
    )
    expect(prompt).toContain('User Requirements')
    expect(prompt).toContain('Focus on quarterly metrics')
  })

  it('includes output format reminder', () => {
    const prompt = buildPrompt(reportTemplate as any)
    expect(prompt).toContain('Output Format')
    expect(prompt).toContain('DocxSchema')
  })
})

describe('buildFreeformPrompt', () => {
  it('includes document type in prompt', () => {
    const prompt = buildFreeformPrompt('business proposal')
    expect(prompt).toContain('business proposal')
  })

  it('lists all available node types', () => {
    const prompt = buildFreeformPrompt('letter')
    expect(prompt).toContain('heading')
    expect(prompt).toContain('paragraph')
    expect(prompt).toContain('table')
    expect(prompt).toContain('bulletList')
  })

  it('lists all available plugins', () => {
    const prompt = buildFreeformPrompt('report')
    expect(prompt).toContain('coverPage')
    expect(prompt).toContain('callout')
    expect(prompt).toContain('qrcode')
  })

  it('includes user context when provided', () => {
    const prompt = buildFreeformPrompt('invoice', 'Include tax calculation')
    expect(prompt).toContain('User Requirements')
    expect(prompt).toContain('Include tax calculation')
  })
})

describe('generateSchemaGuide', () => {
  it('returns a non-empty string', () => {
    const guide = generateSchemaGuide()
    expect(guide.length).toBeGreaterThan(100)
  })

  it('includes DocxSchema structure', () => {
    const guide = generateSchemaGuide()
    expect(guide).toContain('DocxSchema')
    expect(guide).toContain('content')
  })

  it('includes all block node types', () => {
    const guide = generateSchemaGuide()
    expect(guide).toContain('heading')
    expect(guide).toContain('paragraph')
    expect(guide).toContain('table')
    expect(guide).toContain('bulletList')
    expect(guide).toContain('numberedList')
    expect(guide).toContain('pageBreak')
    expect(guide).toContain('plugin')
  })

  it('includes built-in plugins', () => {
    const guide = generateSchemaGuide()
    expect(guide).toContain('callout')
    expect(guide).toContain('qrcode')
    expect(guide).toContain('watermark')
    expect(guide).toContain('echarts')
  })
})

describe('generateToolDefinitions', () => {
  it('returns create_document tool by default', () => {
    const defs = generateToolDefinitions()
    const createDoc = defs.find(d => d.function.name === 'create_document')
    expect(createDoc).toBeDefined()
    expect(createDoc!.function.description).toContain('docx-kit')
    expect(createDoc!.type).toBe('function')
  })

  it('returns validate_schema tool by default', () => {
    const defs = generateToolDefinitions()
    const validate = defs.find(d => d.function.name === 'validate_schema')
    expect(validate).toBeDefined()
  })

  it('returns list_templates tool by default', () => {
    const defs = generateToolDefinitions()
    const list = defs.find(d => d.function.name === 'list_templates')
    expect(list).toBeDefined()
  })

  it('includes list_plugins when plugins option provided', () => {
    const defs = generateToolDefinitions({ plugins: ['qrcode'] })
    const list = defs.find(d => d.function.name === 'list_plugins')
    expect(list).toBeDefined()
  })

  it('excludes list_plugins when no plugins option', () => {
    const defs = generateToolDefinitions()
    const list = defs.find(d => d.function.name === 'list_plugins')
    expect(list).toBeUndefined()
  })

  it('includes template tool when template option provided', () => {
    const defs = generateToolDefinitions({ template: 'report' })
    const templateTool = defs.find(
      d => d.function.name === 'apply_template_report',
    )
    expect(templateTool).toBeDefined()
  })

  it('skips template tool for unknown template name', () => {
    const defs = generateToolDefinitions({ template: 'nonexistent' })
    const templateTool = defs.find(d =>
      d.function.name.startsWith('apply_template_'),
    )
    expect(templateTool).toBeUndefined()
  })

  it('all tool definitions have valid structure', () => {
    const defs = generateToolDefinitions({
      plugins: ['qrcode'],
      template: 'report',
    })
    for (const def of defs) {
      expect(def.type).toBe('function')
      expect(def.function.name).toBeTypeOf('string')
      expect(def.function.name.length).toBeGreaterThan(0)
      expect(def.function.description).toBeTypeOf('string')
      expect(def.function.parameters).toBeTypeOf('object')
      expect(def.function.parameters.type).toBe('object')
    }
  })
})
