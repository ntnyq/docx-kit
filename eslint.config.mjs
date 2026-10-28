// @ts-check

import { defineESLintConfig } from '@ntnyq/eslint-config'

export default defineESLintConfig(
  {
    perfectionist: {
      all: true,
    },
  },
  // Auto-generated artifacts — content is produced by build scripts, not humans.
  {
    ignores: ['apps/playground/src/utils/monacoTypes.generated.ts'],
  },
)
