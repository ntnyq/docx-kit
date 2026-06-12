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
  it('exports saveDocument and dataUrlToUint8Array', async () => {
    const mod = await import('../src/node')
    expect(mod.saveDocument).toBeDefined()
    expect(mod.dataUrlToUint8Array).toBeDefined()
  })
})

describe('browser.ts — platform exports', () => {
  it('exports dataUrlToUint8Array and normalizeImageData', async () => {
    const mod = await import('../src/browser')
    expect(mod.dataUrlToUint8Array).toBeDefined()
    expect(mod.normalizeImageData).toBeDefined()
  })
})
