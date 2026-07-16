/**
 * Package.json template — generates the plugin's package.json.
 *
 * @module templates/plugin/package-json
 */

/* eslint-disable perfectionist/sort-objects */

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
      name: pluginName,
      type: 'module',
      version,
      description,
      keywords: ['docx-kit-plugin', shortName],
      license,
      author,
      exports: {
        '.': {
          types: './dist/index.d.ts',
          default: './dist/index.js',
        },
      },
      main: './dist/index.js',
      types: './dist/index.d.ts',
      files: ['README.md', 'dist', 'docx-kit.plugin.json'],
      publishConfig: {
        access: 'public',
      },
      sideEffects: false,
      scripts: {
        build: 'tsdown',
        lint: 'eslint',
        test: 'vitest run',
        typecheck: 'tsc --noEmit',
      },
      peerDependencies: {
        docx: '^9.7.1',
        'docx-kit': '^0.3.0',
      },
      devDependencies: {
        '@ntnyq/eslint-config': '^6.1.5',
        docx: '^9.7.1',
        'docx-kit': '^0.3.0',
        eslint: '^10.7.0',
        tsdown: '^0.22.9',
        typescript: '^6.0.3',
        vitest: '^4.1.10',
      },
    },
    null,
    2,
  )
}
