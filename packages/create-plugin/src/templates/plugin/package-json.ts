/**
 * Package.json template — generates the plugin's package.json.
 *
 * @module templates/plugin/package-json
 */

/* eslint-disable perfectionist/sort-objects */

import { DOCX_KIT_RANGE } from '../../version'

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
        docx: '^9.15.0',
        'docx-kit': DOCX_KIT_RANGE,
      },
      devDependencies: {
        '@ntnyq/eslint-config': '^7.0.0-beta.28',
        '@typescript/native': 'npm:typescript@^7.0.2',
        docx: '^9.15.0',
        'docx-kit': DOCX_KIT_RANGE,
        eslint: '^10.13.0',
        tsdown: '^0.23.3',
        typescript: 'npm:@typescript/typescript6@^6.0.2',
        vitest: '^5.0.5',
      },
    },
    null,
    2,
  )
}
