import { describe, expect, it } from 'vitest'
import packageJson from '../package.json'
import { renderEslintConfig } from '../src/templates/plugin/eslint-config'
import { renderManifest } from '../src/templates/plugin/manifest'
import { renderPackageJson } from '../src/templates/plugin/package-json'
import { renderReadme } from '../src/templates/plugin/readme-md'
import { renderPluginSource } from '../src/templates/plugin/src-index'
import { renderPluginTest } from '../src/templates/plugin/test-index'
import { renderTsconfigJson } from '../src/templates/plugin/tsconfig-json'
import { renderTsdownConfig } from '../src/templates/plugin/tsdown-config'

describe('renderManifest', () => {
  it('generates valid manifest JSON', () => {
    const result = renderManifest(
      'docx-kit-plugin-chart',
      '0.1.0',
      '^0.3.0',
      'A chart plugin',
    )
    const parsed = JSON.parse(result)
    expect(parsed.name).toBe('docx-kit-plugin-chart')
    expect(parsed.version).toBe('0.1.0')
    expect(parsed.docxKit).toBe('^0.3.0')
    expect(parsed.main).toBe('./dist/index.js')
    expect(parsed.plugin.name).toBe('chart')
    expect(parsed.description).toBe('A chart plugin')
  })

  it('uses defaults for optional fields', () => {
    const result = renderManifest('docx-kit-plugin-chart')
    const parsed = JSON.parse(result)
    expect(parsed.version).toBe('0.1.0')
    expect(parsed.docxKit).toBe(`^${packageJson.version}`)
    expect(parsed.description).toBe('')
  })
})

describe('renderPluginSource', () => {
  it('generates source with correct plugin name', () => {
    const result = renderPluginSource('chart')
    expect(result).toContain("definePlugin<'chart'")
    expect(result).toContain("name: 'chart'")
    expect(result).toContain('ChartOptions')
  })

  it('capitalizes options interface name', () => {
    const result = renderPluginSource('myWidget')
    expect(result).toContain('MyWidgetOptions')
  })

  it('creates valid identifiers for kebab-case plugin names', () => {
    const result = renderPluginSource('my-widget')
    expect(result).toContain('MyWidgetOptions')
    expect(result).toContain('myWidgetPlugin')
    expect(result).toContain("name: 'my-widget'")
  })
})

describe('renderPluginTest', () => {
  it('generates test with correct plugin name', () => {
    const result = renderPluginTest('chart')
    expect(result).toContain('chartPlugin')
    expect(result).toContain("'chart'")
    expect(result).toContain("renderPlugin(plugin, { text: 'Hello' })")
  })
})

describe('renderPackageJson', () => {
  it('generates package.json with all fields', () => {
    const result = renderPackageJson(
      'docx-kit-plugin-chart',
      'chart',
      'A chart plugin',
      'John Doe',
      'MIT',
      '0.1.0',
    )
    const parsed = JSON.parse(result)
    expect(parsed.name).toBe('docx-kit-plugin-chart')
    expect(parsed.description).toBe('A chart plugin')
    expect(parsed.author).toBe('John Doe')
    expect(parsed.license).toBe('MIT')
    expect(parsed.version).toBe('0.1.0')
    expect(parsed.keywords).toContain('docx-kit-plugin')
    expect(parsed.keywords).toContain('chart')
    expect(parsed.exports['.'].default).toBe('./dist/index.js')
    expect(parsed.files).toContain('docx-kit.plugin.json')
    expect(parsed.peerDependencies.docx).toBe('^9.7.1')
    expect(parsed.peerDependencies['docx-kit']).toBe(`^${packageJson.version}`)
    expect(parsed.devDependencies['docx-kit']).toBe(`^${packageJson.version}`)
    expect(parsed.devDependencies.tsdown).toBe('^0.22.9')
    expect(parsed.devDependencies.vitest).toBe('^4.1.10')
  })

  it('uses defaults for optional fields', () => {
    const result = renderPackageJson('docx-kit-plugin-chart', 'chart')
    const parsed = JSON.parse(result)
    expect(parsed.author).toBe('')
    expect(parsed.license).toBe('MIT')
    expect(parsed.version).toBe('0.1.0')
  })
})

describe('renderTsconfigJson', () => {
  it('generates valid tsconfig', () => {
    const result = renderTsconfigJson()
    const parsed = JSON.parse(result)
    expect(parsed.compilerOptions.strict).toBe(true)
    expect(parsed.compilerOptions.target).toBe('ESNext')
    expect(parsed.compilerOptions.module).toBe('ESNext')
    expect(parsed.include).toContain('src')
    expect(parsed.include).toContain('tests')
  })
})

describe('build config templates', () => {
  it('generates tsdown and ESLint configs', () => {
    expect(renderTsdownConfig()).toContain("entry: ['src/index.ts']")
    expect(renderEslintConfig()).toContain('defineESLintConfig')
  })
})

describe('renderReadme', () => {
  it('generates README with plugin name', () => {
    const result = renderReadme(
      'docx-kit-plugin-chart',
      'chart',
      'A chart plugin',
      'John Doe',
    )
    expect(result).toContain('# docx-kit-plugin-chart')
    expect(result).toContain('A chart plugin')
    expect(result).toContain('npm install docx-kit-plugin-chart')
    expect(result).toContain('chartPlugin')
    expect(result).toContain(".plugin('chart', { text: 'Hello' })")
    expect(result).toContain('await doc.toUint8Array()')
    expect(result).toContain('John Doe')
  })

  it('omits author when empty', () => {
    const result = renderReadme('docx-kit-plugin-chart', 'chart')
    expect(result).toContain('MIT')
    expect(result).not.toContain('MIT — ')
  })
})
