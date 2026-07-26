import { describe, expect, it } from 'vitest'

describe('shared.ts — shared exports', () => {
  it('exports createDocx and renderDocx', async () => {
    const mod = await import('../src/browser')
    expect(mod.createDocx).toBeDefined()
    expect(mod.renderDocx).toBeDefined()
    expect(mod.DocxBuilder).toBeDefined()
  })

  it('exports defineStyles and definePlugin', async () => {
    const mod = await import('../src/browser')
    expect(mod.defineStyles).toBeDefined()
    expect(mod.definePlugin).toBeDefined()
  })

  it('exports ERROR_CODES and DocxKitError', async () => {
    const mod = await import('../src/browser')
    expect(mod.ERROR_CODES).toBeDefined()
    expect(mod.DocxKitError).toBeDefined()
  })

  it('exports plugins', async () => {
    const mod = await import('../src/browser')
    expect(mod.qrcodePlugin).toBeDefined()
    expect(mod.echartsPlugin).toBeDefined()
    expect(mod.badgePlugin).toBeDefined()
    expect(mod.invoicePlugin).toBeDefined()
    expect(mod.tocPlugin).toBeDefined()
  })

  it('exports dataUrlToUint8Array', async () => {
    const mod = await import('../src/browser')
    expect(mod.dataUrlToUint8Array).toBeDefined()
  })

  it('exports all type definitions', async () => {
    const mod = await import('../src/browser')
    // Types are compile-time only, but the module should load clean
    expect(mod).toBeDefined()
    // These are documented exports from shared.ts
    expect(typeof mod.createDocx).toBe('function')
  })
})

describe('node.ts — platform exports', () => {
  it('exports Node output and data URL helpers', async () => {
    const mod = await import('../src/node')
    expect(mod.saveDocument).toBeDefined()
    expect(mod.streamDocument).toBeDefined()
    expect(mod.dataUrlToUint8Array).toBeDefined()
  })

  it('installs save and stream output on fluent and JSON builders', async () => {
    const mod = await import('../src/node')
    const fluentBuilder = mod.createDocx().use({
      name: 'test',
      render: () => [],
    })
    const jsonBuilder = await mod.renderDocx({ content: [] })

    expect(typeof fluentBuilder.save).toBe('function')
    expect(typeof fluentBuilder.toStream).toBe('function')
    expect(typeof jsonBuilder.save).toBe('function')
    expect(typeof jsonBuilder.toStream).toBe('function')
  })
})

describe('browser.ts — platform exports', () => {
  it('exports dataUrlToUint8Array and normalizeImageData', async () => {
    const mod = await import('../src/browser')
    expect(mod.dataUrlToUint8Array).toBeDefined()
    expect(mod.normalizeImageData).toBeDefined()
  })
})

describe('dist exports — published package smoke tests', () => {
  it('loads browser dist entrypoint', async () => {
    const mod = await import('../dist/browser.js')
    expect(mod.badgePlugin).toBeDefined()
    expect(mod.invoicePlugin).toBeDefined()
    expect(mod.tocPlugin).toBeDefined()
  })

  it('loads node dist entrypoint', async () => {
    const mod = await import('../dist/node.js')
    expect(mod.saveDocument).toBeDefined()
    expect(mod.streamDocument).toBeDefined()
    expect(mod.letterheadPlugin).toBeDefined()
    expect(typeof mod.createDocx().save).toBe('function')
    expect(typeof mod.createDocx().toStream).toBe('function')
  })

  it.each([
    ['docx-kit/ai', 'buildPrompt'],
    ['docx-kit/loader', 'createPluginLoader'],
    ['docx-kit/loader/browser', 'loadUrlPlugin'],
    ['docx-kit/loader/node', 'loadNpmPlugin'],
    ['docx-kit/mcp', 'TOOL_DEFINITIONS'],
    ['docx-kit/pdk', 'createPluginTestContext'],
    ['docx-kit/registry', 'createPluginRegistry'],
  ] as const)('loads package export %s', async (specifier, exportName) => {
    const mod = await import(specifier)
    expect(mod[exportName]).toBeDefined()
  })
})
