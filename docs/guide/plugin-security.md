# Plugin Security Guide

docx-kit's plugin system allows loading external code from npm packages, URLs, and local files. This flexibility comes with security considerations. This guide explains the trust model, permission hooks, and best practices for safe plugin consumption and authoring.

## Trust Levels

Plugins fall into four trust tiers based on their origin and verification:

| Level        | Source                                    | Verification                              | Risk    |
|--------------|-------------------------------------------|-------------------------------------------|---------|
| **Built-in** | Included in docx-kit source               | Tested & reviewed with each release       | Minimal |
| **Verified** | Published by the docx-kit team            | Manifest validated, version checked       | Low     |
| **Community**| Published by third parties on npm         | Manifest validated, keyword tagged        | Medium  |
| **Untrusted**| Loaded from arbitrary URLs or local paths | No guaranteed verification                | High    |

Built-in plugins (callout, watermark, echarts, qrcode, etc.) are compiled and tested as part of the docx-kit release cycle. Verified plugins follow the same review process but live in separate packages. Community plugins are only validated at the manifest level. Untrusted plugins carry the highest risk since they may bypass manifest validation.

## Permission Hooks

The `PluginLoader` provides two security hooks that let consumers control what gets loaded and executed:

### `allowLoad(source)`

Called **before** loading begins. Receives the `PluginSource` descriptor. Return `false` to block the load entirely.

```ts
import { createPluginLoader } from 'docx-kit/loader'

const loader = createPluginLoader({
  security: {
    // Only allow npm packages from trusted registries
    allowLoad(source) {
      if (source.type === 'url') {
        const domain = new URL(source.url).origin
        return domain === 'https://trusted.example.com'
      }
      return true // allow npm and inline sources
    },
  },
})
```

This hook is ideal for whitelisting domains, restricting npm package names, or blocking URL-based loading in environments where network fetches are prohibited.

### `allowExecute(manifest, source)`

Called **after** the manifest is resolved but **before** the plugin code is executed. Receives the parsed `PluginManifest` and the original `PluginSource`. Return `false` to block execution.

```ts
const loader = createPluginLoader({
  security: {
    // Only execute plugins with compatible version ranges
    allowExecute(manifest, source) {
      // Block plugins that require exact versions (no semver ranges)
      if (!manifest.docxKit.startsWith('^') && manifest.docxKit !== '*') {
        console.warn(`Blocked plugin with pinned version: ${manifest.plugin.name}`)
        return false
      }
      return true
    },
  },
})
```

This hook gives you access to the manifest metadata (author, version range, dependencies) before the plugin code runs. Use it to enforce version policies, reject unmaintained packages, or block plugins with suspicious dependency lists.

## Manifest Validation

Every plugin loaded from npm, URL, or local sources goes through `validateManifest()` by default. This ensures:

- Required fields (`name`, `version`, `docxKit`, `plugin.name`) are present and correctly typed
- `version` is valid semver (e.g. `1.0.0`)
- `docxKit` is a valid semver range (e.g. `^0.2.0`, `~0.3.0`, `*`)
- Optional fields (`main`, `types`, `dependencies`, `peerDependencies`) have correct types

You can disable manifest validation when loading trusted or test plugins:

```ts
const loader = createPluginLoader({
  validateManifest: false, // skip manifest checks
})
```

**Warning:** Disabling validation is only appropriate for local development or when loading plugins you fully control. Never disable validation for URL-based or community plugins.

## Version Compatibility

The loader checks each plugin's `docxKit` range against the installed docx-kit version:

| Range       | Meaning                                      | Example                |
|-------------|----------------------------------------------|------------------------|
| `*`         | Any version — always compatible              | Wildcard               |
| `^0.2.0`    | Same major version, >= minor                 | Loose caret            |
| `~0.2.0`    | Same major.minor, >= patch                   | Tight tilde            |

If a plugin declares `^0.2.0` but the installed version is `1.0.0`, the loader throws `PLUGIN_VERSION_MISMATCH`. This prevents plugins from silently breaking when the host library evolves.

## Best Practices for Consumers

1. **Always validate manifests.** Keep `validateManifest: true` (the default) for any plugin source you don't fully control.

2. **Restrict URL sources.** Use `allowLoad()` to whitelist specific domains. Never load plugins from arbitrary URLs in production.

3. **Check version ranges.** Use `allowExecute()` to reject plugins with overly specific or wildcard ranges that could mask compatibility issues.

4. **Load from npm when possible.** npm packages are at least validated by the registry and tagged with `keyword:docx-kit-plugin`. Use the registry search to discover plugins:

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const plugins = await registry.search('chart')
   ```

5. **Use `loadAll()` for batch loading.** It collects successes and warns on failures without throwing, making it safe for startup sequences:

   ```ts
   const results = await loader.loadAll([
     { plugin: calloutPlugin(), type: 'inline' },
     { package: 'docx-kit-plugin-chart', type: 'npm' },
   ])
   ```

## Best Practices for Authors

1. **Always include a manifest.** Create a `docx-kit.plugin.json` in your package root. This allows the loader to validate and check compatibility.

2. **Declare a caret range for `docxKit`.** Use `^0.2.0` (or the current major.minor) rather than `*` or pinned versions. Caret ranges communicate intent and allow the compatibility check to protect users.

3. **Tag your npm package.** Add `keyword:docx-kit-plugin` to your `package.json` keywords. This makes your plugin discoverable via the registry search.

4. **Declare peer dependencies.** List `docx-kit` as a peerDependency (not a regular dependency) to avoid version conflicts in consumer projects.

5. **Write tests using the PDK.** Use `renderPlugin()` and `assertRendersParagraph()` from `docx-kit/pdk` to test your plugin in isolation:

   ```ts
   import { renderPlugin, assertRendersParagraph } from 'docx-kit/pdk'
   import { myPlugin } from './src'

   const result = await renderPlugin(myPlugin(), { text: 'Hello' })
   assertRendersParagraph(result, 'Hello')
   ```

## Sandboxing (Future)

Current plugin loading executes code in the same context as the host application. Future versions may introduce sandboxed execution via:

- **Web Workers** for browser environments
- **VM modules** (`node:vm`) for Node.js
- **Content Security Policy** integration for iframe-based isolation

Until sandboxing is available, rely on the permission hooks and manifest validation as your primary defense layers.