/**
 * ESLint config template for generated plugin projects.
 *
 * @returns The ESLint configuration source for a generated plugin project
 */
export function renderEslintConfig(): string {
  return `import { defineESLintConfig } from '@ntnyq/eslint-config'

export default defineESLintConfig()
`
}
