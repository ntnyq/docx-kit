import { fileURLToPath, URL } from 'node:url'
import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import sucrasePackage from 'sucrase/package.json' with { type: 'json' }
import { defineConfig } from 'vite'
import type { Plugin } from 'vite'

export default defineConfig({
  build: { chunkSizeWarningLimit: 7 * 1024 },
  plugins: [vue(), tailwindcss(), sucraseTypeScriptProfile()],
  server: { port: 5173 },
  worker: { format: 'es' },
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
    dedupe: ['vue', 'vue-i18n', '@vueuse/core'],
  },
})

/**
 * The playground always requests only TypeScript. Fold Sucrase's transform
 * selection checks so tree shaking removes the other transformer classes.
 * Parser and ES/TypeScript runtime transforms stay on the upstream code path.
 */
export function sucraseTypeScriptProfile(): Plugin {
  const transformedModules = new Set<string>()
  const modes = new Set([
    'flow',
    'imports',
    'jest',
    'jsx',
    'react-hot-loader',
    'typescript',
  ])
  return {
    apply: 'build',
    enforce: 'pre',
    name: 'docxkit:sucrase-typescript-profile',
    buildStart() {
      if (sucrasePackage.version !== '3.35.1') {
        throw new Error(
          'Review the TypeScript compiler profile before upgrading Sucrase',
        )
      }
    },
    generateBundle() {
      if (transformedModules.size !== 2) {
        throw new Error(
          'The Sucrase TypeScript profile was not applied to both modules',
        )
      }
    },
    transform(code, id) {
      const path = id.replaceAll('\\', '/')
      if (
        path.endsWith('/src/utils/prepareCode.ts')
        && (!/\btransform\(\s*raw,\s*\{\s*keepUnusedImports:\s*true,\s*transforms:\s*\[\s*['"]typescript['"]\s*(?:,\s*)?\]\s*(?:,\s*)?\}\s*(?:,\s*)?\)/u.test(
          code,
        )
          || code.match(/\btransform\s*\(/gu)?.length !== 1)
      ) {
        throw new Error(
          'Review the Sucrase profile when changing compiler transforms',
        )
      }
      const isEntryModule = path.endsWith('/sucrase/dist/esm/index.js')
      // Audited counts catch changes to the pinned dependency's source layout.
      const expectedCount = isEntryModule
        ? 7
        : path.endsWith('/sucrase/dist/esm/transformers/RootTransformer.js')
          ? 14
          : undefined
      if (expectedCount === undefined) {
        return
      }
      let count = 0
      let result = code.replace(
        /\b(?:options\.)?transforms\.includes\("([a-z-]+)"\)/gu,
        (_match: string, mode: string) => {
          if (!modes.has(mode)) {
            throw new Error(`Unknown Sucrase transform: ${mode}`)
          }
          count += 1
          return mode === 'typescript' ? 'true' : 'false'
        },
      )
      if (count !== expectedCount) {
        throw new Error(
          'Sucrase internals changed; review the TypeScript compiler profile',
        )
      }
      if (isEntryModule) {
        // The caller above supplies a typed, fixed options object. It never
        // requests source maps or accepts user-supplied compiler options.
        const unusedPaths = [
          ['import {validateOptions} from "./Options";', ''],
          ['  validateOptions(options);', ''],
          ['import computeSourceMap, {} from "./computeSourceMap";', ''],
          ['    if (options.sourceMapOptions) {', '    if (false) {'],
        ]
        for (const [original, replacement] of unusedPaths) {
          if (result.split(original).length !== 2) {
            throw new Error(
              'Sucrase option handling changed; review the compiler profile',
            )
          }
          result = result.replace(original, () => replacement)
        }
      }
      transformedModules.add(path)
      return {
        code: result,
        map: null,
      }
    },
  }
}
