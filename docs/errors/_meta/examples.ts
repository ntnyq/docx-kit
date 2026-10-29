import type { ErrorCode } from '../../../packages/errors/src/index'
import type { ErrorExample } from './types'

export const examples = {
  EXPORT_FAILED: {
    guide: '/guide/platforms',
    after:
      "import { createDocx, DocxKitError, serializeDocxKitError } from 'docx-kit'\n\ntry {\n  const blob = await createDocx().p('Hello').toBlob()\n} catch (error) {\n  if (error instanceof DocxKitError) {\n    console.error(serializeDocxKitError(error))\n    // Inspect error.cause locally when needed.\n  } else {\n    throw error\n  }\n}",
    before:
      '// Packaging can fail when a referenced resource is invalid.\n// Inspect the original cause at the export boundary.',
  },
  HEADER_FOOTER_INVALID_CONTENT: {
    guide: '/guide/plugins',
    after:
      "import { Paragraph, TextRun } from 'docx'\n\n// Valid output from a plugin used in a header/footer:\nconst render = () => new Paragraph({ children: [new TextRun('Header')] })",
    before:
      "import { TextRun } from 'docx'\n\n// Invalid output from a plugin used in a header/footer:\nconst render = () => new TextRun('Header')",
  },
  IMAGE_INVALID_DATA: {
    guide: '/guide/images',
    after:
      "import { createDocx } from 'docx-kit'\n\nconst data = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aX1cAAAAASUVORK5CYII='\nconst blob = await createDocx().image({ data }).toBlob()",
    before:
      "import { createDocx } from 'docx-kit'\n\nawait createDocx().image({ data: '' }).toBlob()",
  },
  MANIFEST_INVALID: {
    guide: '/guide/plugins',
    after:
      "import { validateManifest } from 'docx-kit'\n\nconst manifest = {\n  name: 'example-plugin', version: '1.0.0', docxKit: '^0.7.0',\n  main: './dist/index.js', plugin: { name: 'example' },\n}\n\nvalidateManifest(manifest)",
    before:
      "import { validateManifest } from 'docx-kit'\n\nconst manifest = {\n  name: 'example-plugin', version: '1.0.0', docxKit: '^0.7.0',\n  main: './dist/index.js', plugin: { name: 'example' },\n}\n\nvalidateManifest({ ...manifest, main: '../outside.js' })",
  },
  MANIFEST_MISSING: {
    before: '// Package contents: dist/index.js only',
    guide: '/guide/plugins',
    after:
      "// Package contents:\n// docx-kit.plugin.json\n// dist/index.js\nimport { validateManifest } from 'docx-kit'\n\nconst manifest = {\n  name: 'example-plugin', version: '1.0.0', docxKit: '^0.7.0',\n  main: './dist/index.js', plugin: { name: 'example' },\n}\n\nvalidateManifest(manifest)",
  },
  MICROSOFT_URL_REQUIRED: {
    guide: '/guide/preview',
    after:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\n// Local documents use the DOM renderer.\nconst preview = createDocxPreview(container, { renderer: 'dom' })\ntry {\n  await preview.render(await createDocx().p('Hello').toBlob())\n} finally { preview.destroy() }",
    before:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container, { renderer: 'microsoft' })\ntry {\n  await preview.render(await createDocx().p('Hello').toBlob())\n} finally { preview.destroy() }",
  },
  PLUGIN_DEPENDENCY_MISSING: {
    guide: '/guide/plugins',
    after:
      "// Install a compatible dependency declared by the plugin manifest.\n// Built-in loaders do not currently raise this code automatically.\nconst manifestDependencies = { 'example-dependency': '^1.0.0' }",
    before:
      "import { createDocxKitError } from 'docx-kit'\n\n// Explicit reporting by a custom dependency validator.\nthrow createDocxKitError('PLUGIN_DEPENDENCY_MISSING', {\n  dependency: 'example-dependency',\n})",
  },
  PLUGIN_LOAD_FAILED: {
    before: "// An invalid module export:\nexport default { name: 'example' }",
    guide: '/guide/plugin-security',
    after:
      "import { definePlugin } from 'docx-kit'\nimport { Paragraph } from 'docx'\n\nexport default definePlugin({\n  name: 'example',\n  render: () => new Paragraph('Hello'),\n})",
  },
  PLUGIN_NOT_REGISTERED: {
    guide: '/guide/plugins',
    after:
      "import { createDocx, qrcodePlugin } from 'docx-kit'\n\nconst doc = createDocx()\n  .use(qrcodePlugin())\n  .plugin('qrcode', { text: 'https://example.com' })\nconst blob = await doc.toBlob()",
    before:
      "import { createDocx } from 'docx-kit'\n\nawait createDocx()\n  .plugin('qrcode', { text: 'https://example.com' })\n  .toBlob()",
  },
  PLUGIN_RENDER_FAILED: {
    guide: '/guide/plugins',
    after:
      "import { createDocx, definePlugin } from 'docx-kit'\nimport { Paragraph } from 'docx'\n\nconst plugin = definePlugin({\n  name: 'example', render: () => new Paragraph('Hello'),\n})\nconst blob = await createDocx().use(plugin).plugin('example', {}).toBlob()",
    before:
      "import { createDocx, definePlugin } from 'docx-kit'\n\nconst plugin = definePlugin({\n  name: 'example', render() { throw new Error('Resource unavailable') },\n})\nawait createDocx().use(plugin).plugin('example', {}).toBlob()",
  },
  PLUGIN_VERSION_MISMATCH: {
    before: "// With docx-kit 0.7.x installed:\nconst docxKit = '^0.6.0'",
    guide: '/guide/plugins',
    after:
      "// Use a plugin release tested with the installed docx-kit version.\nconst docxKit = '^0.7.0'",
  },
  PREVIEW_FETCH_FAILED: {
    guide: '/guide/preview',
    after:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container)\ntry {\n  const blob = await createDocx().p('Hello').toBlob()\n  await preview.render(blob)\n} finally { preview.destroy() }",
    before:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container)\ntry {\n  await preview.render('https://example.com/missing.docx')\n} finally { preview.destroy() }",
  },
  PREVIEW_INPUT_INVALID: {
    before: '// Invalid external input:\nconst input = null',
    guide: '/guide/preview',
    after:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container)\ntry {\n  await preview.render(await createDocx().p('Hello').toBlob())\n} finally { preview.destroy() }",
  },
  PREVIEW_RENDER_FAILED: {
    guide: '/guide/preview',
    after:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container)\ntry {\n  await preview.render(await createDocx().p('Hello').toBlob())\n} finally { preview.destroy() }",
    before:
      "import { createDocx, createDocxPreview } from 'docx-kit'\n\nconst container = document.createElement('div')\ndocument.body.append(container)\n\nconst preview = createDocxPreview(container)\ntry {\n  await preview.render(new Blob(['plain text']))\n} finally { preview.destroy() }",
  },
  STYLE_CIRCULAR_EXTENDS: {
    guide: '/guide/styling',
    after:
      "import { createDocx, defineStyles } from 'docx-kit'\n\nconst styles = defineStyles({\n  body: { fontSize: 12 }, title: { extends: 'body', fontSize: 24 },\n})\nconst blob = await createDocx({ styles })\n  .p('Hello', { className: 'title' }).toBlob()",
    before:
      "import { createDocx, defineStyles } from 'docx-kit'\n\nconst styles = defineStyles({\n  body: { extends: 'title' }, title: { extends: 'body' },\n})\nawait createDocx({ styles }).p('Hello', { className: 'body' }).toBlob()",
  },
  STYLE_UNKNOWN_CLASS: {
    guide: '/guide/styling',
    after:
      "import { createDocx, defineStyles } from 'docx-kit'\n\nconst styles = defineStyles({ body: { fontSize: 12 } })\nconst blob = await createDocx({ styles })\n  .p('Hello', { className: 'body' })\n  .toBlob()",
    before:
      "import { defineStyles } from 'docx-kit'\n\ndefineStyles({ body: { extends: 'missing' } })",
  },
  TABLE_INVALID_COLUMNS: {
    before: '// Invalid table grid: no columns.\nconst columns = []',
    guide: '/guide/tables',
    after:
      "import { createDocx } from 'docx-kit'\n\nconst blob = await createDocx().table({\n  columns: [{ key: 'name', title: 'Name' }],\n  data: [{ name: 'Ada' }],\n}).toBlob()",
  },
  UNKNOWN_NODE_TYPE: {
    guide: '/guide/json-dsl',
    after:
      "import { renderDocx } from 'docx-kit'\n\nconst doc = await renderDocx({\n  content: [{ type: 'paragraph', text: 'Hello' }],\n})\nconst blob = await doc.toBlob()",
    before:
      "// Invalid DSL fragment:\nconst node = { type: 'paragaph', text: 'Hello' }",
  },
} satisfies Record<ErrorCode, ErrorExample>
