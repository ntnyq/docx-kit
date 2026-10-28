# 插件安全指南

docx-kit 的插件系统支持从 npm 包、URL 与本地文件加载外部代码。本指南介绍其信任模型、权限钩子，以及安全使用和编写插件的实践。

## 信任级别

按来源与验证方式，插件分为四个信任级别：

| 级别 | 来源 | 验证方式 | 风险 |
|--------------|-------------------------------------------|-------------------------------------------|---------|
| **内置** | 包含在 docx-kit 源码中 | 每次发布均经过测试和审查 | 极低 |
| **已验证** | 由 docx-kit 团队发布 | 校验清单与版本 | 低 |
| **社区** | 第三方发布到 npm | 校验清单与关键字标签 | 中 |
| **不受信任** | 任意 URL 或本地路径 | 校验清单，但不信任发布者 | 高 |

内置插件（callout、watermark、echarts、qrcode 等）随 docx-kit 发布周期一起编译和测试。已验证插件遵循相同审查流程，但位于独立包。社区、URL 和本地插件接受清单级校验。URL 与本地插件风险最高，因为清单有效并不代表可执行代码可信。

## 权限钩子

`PluginLoader` 提供两个安全钩子，让使用者控制加载和执行哪些插件：

### `allowLoad(source)`

在加载开始**之前**调用，接收 `PluginSource` 描述符。返回 `false` 可完全阻止加载。

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

适合设置域名白名单、限制 npm 包名，或在禁止网络请求的环境中阻止 URL 加载。

### `allowExecute(manifest, source)`

在解析清单**之后**、执行插件代码**之前**调用，接收解析后的 `PluginManifest` 与原始 `PluginSource`。返回 `false` 可阻止执行。

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

可在代码执行前检查清单元数据（作者、版本范围、依赖），以执行版本策略、拒绝无人维护的包或阻止依赖列表可疑的插件。

## 清单校验

默认情况下，从 npm、URL 或本地来源加载的每个插件都经过 `validateManifest()`，确保：

- 必填字段（`name`、`version`、`docxKit`、`main`、`plugin.name`）存在且类型正确
- `version` 为有效 semver（如 `1.0.0`）
- `docxKit` 为有效 semver 范围（如 `^0.2.0`、`~0.3.0`、`*`）
- `main` 和可选 `types` 为以 `./` 开头的安全包内相对路径
- 可选依赖映射与导出映射只包含字符串值

加载可信或测试插件时，可以关闭清单校验：

```ts
const loader = createPluginLoader({
  validateManifest: false, // skip manifest checks
})
```

**警告：** 关闭校验仅适用于本地开发或完全由你控制的插件。不要为 URL 或社区插件关闭校验。

## 版本兼容性

加载器会检查插件的 `docxKit` 范围是否匹配已安装的 docx-kit 版本：

| 范围 | 含义 | 示例说明 |
|-------------|----------------------------------------------|------------------------|
| `*` | 任意版本，始终兼容 | 通配符 |
| `^0.2.0` | >= 0.2.0 且 < 0.3.0 | 脱字符范围（0.x 版本保持次版本兼容） |
| `~0.2.0` | >= 0.2.0 且 < 0.3.0 | 波浪号范围，允许补丁版本更新 |

若插件声明 `^0.2.0`，但已安装版本为 `1.0.0`，加载器会抛出 `PLUGIN_VERSION_MISMATCH`，避免宿主库升级后插件静默失效。

## 使用者最佳实践

1. **始终校验清单。** 对于未完全控制的插件来源，保留默认的 `validateManifest: true`。

2. **限制 URL 来源。** 使用 `allowLoad()` 设置域名白名单，生产环境不要从任意 URL 加载插件。

3. **检查版本范围。** 使用 `allowExecute()` 拒绝过于精确或使用通配符的范围，避免掩盖兼容性问题。

4. **尽量从 npm 加载。** npm 包至少经过注册表校验，并可通过 `keyword:docx-kit-plugin` 发现。使用注册表搜索插件：

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const plugins = await registry.search('chart')
   ```

5. **使用 `loadAll()` 批量加载。** 它收集成功结果，对失败发出警告而不抛出异常，适合启动流程：

   ```ts
   const results = await loader.loadAll([
     { plugin: calloutPlugin(), type: 'inline' },
     { package: 'docx-kit-plugin-chart', type: 'npm' },
   ])
   ```

## 作者最佳实践

1. **始终提供清单。** 在包根目录创建 `docx-kit.plugin.json`，以便加载器校验并检查兼容性。

2. **为 `docxKit` 声明脱字符范围。** 使用 `^0.2.0`（或当前主版本、次版本），避免 `*` 或固定版本，以表达兼容意图并保护使用者。

3. **添加 npm 关键字。** 在 `package.json` 的 keywords 中添加 `docx-kit-plugin`，让注册表搜索能够发现插件。

4. **声明对等依赖。** 将 `docx-kit` 放入 peerDependency，而非普通依赖，避免使用者项目出现版本冲突。

5. **使用 PDK 编写测试。** 通过 `docx-kit/pdk` 的 `renderPlugin()` 和 `assertRendersParagraph()` 独立测试插件：

   ```ts
   import { renderPlugin, assertRendersParagraph } from 'docx-kit/pdk'
   import { myPlugin } from './src'

   const result = await renderPlugin(myPlugin(), { text: 'Hello' })
   assertRendersParagraph(result, 'Hello')
   ```

## 沙箱（未来计划）

当前插件代码与宿主应用在同一上下文执行。未来版本可能通过以下方式引入沙箱：

- 浏览器环境的 **Web Workers**
- Node.js 的 **VM 模块**（`node:vm`）
- 结合**内容安全策略**，通过 iframe 隔离

沙箱支持推出前，应主要依赖权限钩子和清单校验进行防护。
