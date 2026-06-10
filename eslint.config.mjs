// @ts-check

import { defineESLintConfig } from '@ntnyq/eslint-config'

export default defineESLintConfig({
  ignores: ['dist/**', 'docs/.vitepress/cache/**', 'docs/.vitepress/dist/**'],
  perfectionist: {
    all: true,
  },
})
