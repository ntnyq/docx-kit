/**
 * Package.json template — generates the plugin's package.json.
 *
 * @module templates/plugin/package-json
 */

/**
 * Render a package.json for a docx-kit plugin.
 *
 * @param pluginName - — npm package name (e.g. `docx-kit-plugin-chart`)
 * @param shortName - — plugin name without prefix (e.g. `chart`)
 * @param description - — Plugin description
 * @param author - — Author name
 * @param license - — License (default `MIT`)
 * @param version - — Initial version (default `0.1.0`)
 */
export function renderPackageJson(
  pluginName: string,
  shortName: string,
  description: string = 'A docx-kit plugin',
  author: string = '',
  license: string = 'MIT',
  version: string = '0.1.0',
): string {
  return JSON.stringify(
    {
      author,
      description,
      keywords: ['docx-kit-plugin', shortName],
      license,
      main: './dist/index.js',
      name: pluginName,
      type: 'module',
      types: './dist/index.d.ts',
      version,
      devDependencies: {
        docx: '^9.7.1',
        'docx-kit': '^0.2.0',
        typescript: '^6.0.3',
        vitest: '^3.2.3',
      },
      peerDependencies: {
        'docx-kit': '^0.2.0',
      },
      scripts: {
        build: 'tsdown',
        lint: 'eslint',
        test: 'vitest run',
        typecheck: 'tsc --noEmit',
      },
    },
    null,
    2,
  )
}
