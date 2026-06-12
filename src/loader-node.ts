/**
 * docx-kit Plugin Loader — Node.js platform entry.
 *
 * Re-exports the shared loader plus Node.js-specific
 * npm package loader.
 *
 * @module docx-kit/loader/node
 * @packageDocumentation
 */

// Shared loader
export * from './loader'

// Node.js-specific npm source loader
export { loadNpmPlugin } from './loader/sources/npm'
