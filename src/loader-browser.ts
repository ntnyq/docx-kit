/**
 * docx-kit Plugin Loader — Browser platform entry.
 *
 * Re-exports the shared loader plus browser-specific
 * URL plugin loader.
 *
 * @module docx-kit/loader/browser
 * @packageDocumentation
 */

// Shared loader
export * from './loader'

// Browser-specific URL source loader
export { loadUrlPlugin } from './loader/sources/url'
