# Nostics 接入评估与实现方案

调研日期：2026-10-29。对象：当前 docx-kit 仓库与 npm 已发布的 `nostics@1.3.0`。
本文是待实施方案；本次仅新增调研文档，没有安装依赖或修改运行时代码。

## 1. 结论

**适合接入，推荐在 `@docxkit/errors` 内组合使用 Nostics，保留现有公共错误契约，并同时交付错误索引与详情页。**

docx-kit 已有跨包共享错误类，Node.js、浏览器、插件、预览和 MCP 又需要不同的错误展示方式，正好适合把“错误定义”和“展示错误”分开。接入的主要收益是统一带类型的参数、修复建议和文档链接；仅增加一个漂亮的控制台 formatter，收益不足以支撑迁移。

首期边界：

- 保留 `DocxKitError`、构造器前三个参数、`name === 'DocxKitError'`、现有 `code` 字符串和 `cause`。
- 内部使用 Nostics catalog 生成结构化诊断，再适配为 `DocxKitError`。
- 库内部创建错误时保持静默，由调用方 catch 或产品边界负责展示。
- 新增三语 `/errors/` 索引和每码一个详情页，旧 `/guide/errors` 继续介绍捕获方式。
- 先迁移现有错误，新增警告码和更细的错误分类放在后续步骤。

不推荐直接把 `DocxKitError` 替换成 `Diagnostic`：Nostics 的 `name` 等于诊断码，直接替换会改变已有异常名称、类身份和构造方式。v1.3.0 虽然也有 `.code`，其 `toJSON()` 仍采用 `name`、`why` 等字段，需要产品自己的兼容序列化层。详见 [v1.3.0 Diagnostic 源码](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/diagnostic.ts)。

## 2. 当前仓库的适配基础与缺口

| 位置 | 当前行为 | 接入时处理 |
| --- | --- | --- |
| `packages/errors/src/index.ts` | 12 个公共常量；`DocxKitError(code, message, cause?)` | 作为适配层入口，兼容原构造器 |
| `packages/errors/tests/errors.test.ts` | 验证错误类、名称、码、cause 和 key/value 一致 | 作为公共契约回归基线 |
| `packages/core/src/compiler/nodes/compilePlugin.ts` | 缺插件与渲染失败分别包装错误 | 改为带 `pluginName` 等参数的 factory，保留 cause |
| `packages/core/src/style/normalizeStyle.ts` | `STYLE_CIRCULAR_EXTENDS` 未列入公共常量 | 纳入目录并补齐文档 |
| `packages/core/src/compiler/compileDocument.ts` | `HEADER_FOOTER_INVALID_CONTENT` 未列入公共常量 | 纳入目录并补齐文档 |
| `packages/renderer/src/errors.ts` | 4 个预览错误码与独立 factory | 保留原导出，接入共享 metadata，避免依赖 core |
| `packages/core/src/builder/PluginManager.ts`、loader | 重复注册、加载失败等直接 `console.warn` | 后续统一 warning 定义，保持原继续执行行为 |
| `packages/mcp/src/index.ts` | 失败响应只保留 message，校验失败另走一条路径 | 输出可识别的 code、fix、docs，保留 `isError` |
| `apps/playground/src/composables/useDocxPlayground.ts` | 错误被 `String(error)` 压成字符串 | 展示错误码、修复建议与详情页链接 |
| `docs/.vitepress/site.ts`、`config.ts` | 英文、简中、繁中导航及本地搜索 | 增加三语错误入口，复用现有主题与搜索 |

现有三语错误指南各列 9 个码。其中 `IMAGE_UNKNOWN_TYPE` 和 `STYLE_INVALID_UNIT` 没有对应的共享常量，本次源码扫描也未发现对应抛出点；不要据此新增虚假的运行时错误。应删除或订正这两项，并以实际目录生成表格。

上述 inventory 来自当前源码，扫描范围为 `packages` 与 `packages-plugins` 的非测试 TypeScript 文件。普通 `Error`、`TypeError`、插件 setup 原始异常和 schema 校验数组仍需逐类评估，不能把正则扫描结果当成完整异常语义分析。

## 3. 参考 x18n 的哪些做法

[x18n 错误索引](https://x18n.dev/errors) 按错误码分组，表格包含 Code、Title、Exit code。每个码有详情页，依次解释 Message、What to do、Exit code 和 Source，例如 [X18N1001](https://x18n.dev/errors/x18n1001)。

它的关键实现是：诊断定义与人工说明分开维护，生成器把两者合并为 Markdown；`--check` 检查产物是否同步。源码可见 [错误定义](https://github.com/x18n-dev/x18n/blob/main/packages/core/src/errors.ts)、[文档生成器](https://github.com/x18n-dev/x18n/blob/main/docs/scripts/generate.ts) 和 [人工说明](https://github.com/x18n-dev/x18n/blob/main/docs/scripts/notes.ts)。这些 main 分支链接会随上游演进，复核记录见配套调研文件。

docx-kit 采用同样的维护模式，页面列改为「错误码 / 问题 / 分类 / 所属包」。当前产品没有统一 CLI 退出码契约，不照搬 Exit code，也不为接入强行把语义化错误码改成数字编号。

## 4. 运行时设计

### 4.1 依赖与模块边界

```text
nostics（固定已核对版本）
  └─ @docxkit/errors
       ├─ core / loader / renderer / pdk / plugins
       └─ 产品边界：Playground、Theme Studio、MCP

packages/errors/src/catalog.ts + docs/errors/_meta/
  └─ scripts/generate-error-docs.ts
       └─ 三语错误索引、详情页 Markdown
```

`nostics@1.3.0` 是 ESM 包，清单没有运行时 dependencies，与本仓库 ESM、neutral 平台构建方向一致。源码和已发布包仍需通过浏览器构建与体积检查确认实际影响；不能把“无 dependencies”等同于“零体积”。来源：[发布包清单](https://unpkg.com/nostics@1.3.0/package.json)。

建议只让 `@docxkit/errors` 直接依赖 Nostics，应用和其他包继续依赖 docx-kit 自己的 API。首期固定精确版本并提交 lockfile。不要把 docs 的翻译正文、Node 文件 reporter 或 dev collector 引入库公共入口。

### 4.2 catalog 与错误适配

建议新增模块：

```text
packages/errors/src/
  catalog.ts          # why / fix / docs，带类型的参数定义
  metadata.ts         # 码的分类、归属、slug；不包含三语长篇正文
  createError.ts      # catalog → DocxKitError，保持静默
  serializeError.ts   # 稳定、可安全输出的产品 DTO
  index.ts            # 保留旧导出，并导出新增 API
```

catalog 示例，展示拟采用的 Nostics 调用方式；这是方案片段，尚未接入或编译验证：

```ts
import { defineDiagnostics } from 'nostics'

export const diagnostics = defineDiagnostics({
  docsBase: code =>
    `https://docx-kit.ntnyq.dev/errors/${code.toLowerCase()}/`,
  reporters: [],
  codes: {
    PLUGIN_NOT_REGISTERED: {
      why: (params: { pluginName: string }) =>
        `Plugin not registered: ${params.pluginName}`,
      fix: 'Register the plugin with .use(pluginFactory()) before exporting.',
    },
  },
})
```

`defineDiagnostics` 的 handle 会返回 `Diagnostic` 并调用已配置 reporters；v1.3.0 默认 reporters 为空。适配层明确使用空数组，避免底层抛错时额外打印。其运行机制应以 [v1.3.0 实现](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/diagnostic.ts) 为准，接入时再次核对文件与导出。

对外新增的目标调用方式：

```ts
throw createDocxKitError('PLUGIN_NOT_REGISTERED', {
  pluginName: node.name,
})
```

`createDocxKitError` 是本产品拟新增的 API，参数必须根据 code 推导。不要把所有参数退化成 `Record<string, unknown>`，也不要使用没有参数类型的万能兜底重载绕过检查。

factory 内部先获得 diagnostic，再创建原来的 `DocxKitError`，把诊断的 `fix`、`docs`、`sources` 和允许公开的 `data` 加到错误上。原三参构造器继续接受任意自定义 string code；内部可增加第四个可选 metadata 参数。旧调用保留原 message 和 cause，未知自定义码不生成不存在的详情链接。

公共兼容要求：

- `instanceof DocxKitError`、`instanceof Error` 和 `name` 保持。
- `error.code === ERROR_CODES.PLUGIN_NOT_REGISTERED` 等比较保持。
- `cause` 保留同一对象引用，包括非 Error 的值。
- 首轮迁移尽量保持现有 message 文本，新增 fix/docs 采用独立字段。
- 栈顶应尽量指向调用位置；Node 可按能力使用 `Error.captureStackTrace`，浏览器允许正常退化，不能依赖 Node 全局。
- `ERROR_CODES` 与 metadata 从同一 catalog 推导或生成，保持 literal 类型及现有 key/value 相等约定。
- `PREVIEW_ERROR_CODES` 保留原名字和四个值，作为共享目录的子集导出，renderer 不增加 core 依赖。

原生 `formatDiagnostic` 使用 `.name` 作为输出码，不能直接传入名称固定为 `DocxKitError` 的适配对象；应传内部 `Diagnostic`，或由产品 formatter 显式读取 `.code`。

### 4.3 展示、序列化与报告策略

库默认只抛异常。Nostics 的 console/file/fetch reporters 是可选能力，不需要为了首期接入而启用任何远程收集。参见 [reporters 与 formatters](https://nostics.dev/guide/reporters-and-formatters)。

新增 serializer 的目标形状：

```json
{
  "code": "PLUGIN_NOT_REGISTERED",
  "message": "Plugin not registered: qrcode",
  "fix": "Register the plugin with .use(pluginFactory()) before exporting.",
  "docs": "https://docx-kit.ntnyq.dev/errors/plugin_not_registered/"
}
```

serializer 由 docx-kit 自己实现，明确输出 `code` 和 `message`；不直接依赖 `Diagnostic.toJSON()` 的字段。默认不输出 stack、cause、整份 schema、文档正文、base64 图片或带凭据 URL。可公开的节点路径用 `data.nodePath`，仅在确实有来源映射时填 `sources`，不要伪造用户文件行号。

应用边界的改动：

1. **Playground**：保存结构化错误状态，呈现 code、message、fix 和「查看解决方法」，按当前语言映射站内 docs URL；未知异常继续显示可读文本。当前 runner 在同一 JS realm 执行，首期无需引入 Worker 或跨窗口协议。
2. **Theme Studio**：仅为生成与导出异常复用展示组件；主题字段校验继续采用现有表单反馈，不把所有校验都转换为异常。
3. **MCP**：继续返回 `isError: true` 和 text content，text 内序列化以上 DTO；schema 校验失败保留原 path/message 明细。`structuredContent` 是否加入，应结合当前 SDK 及 outputSchema 单独验证，不能直接替换既有响应协议。
4. **warnings**：后续新增 `PLUGIN_ALREADY_REGISTERED` 等 warning 定义，并通过每个 builder/loader 的局部 reporter 输出；避免全局可变 reporter 导致并发文档互相串日志。reporter 的异常不能遮蔽真实错误。

不要首期启用 production stripping。当前主要诊断需要被抛出，生产环境也需要错误码、修复建议与链接；官方 strip 能力主要面向 report-only 调用，并非把所有异常安全移除。参见 [production builds](https://nostics.dev/guide/production-builds)。

## 5. 错误文档页面

### 5.1 路由

| 页面 | 英文 | 简中 / 繁中 |
| --- | --- | --- |
| 错误索引 | `/errors/` | `/zh-CN/errors/`、`/zh-TW/errors/` |
| 单码详情 | `/errors/plugin_not_registered/` | 对应语言前缀 |
| 捕获教程 | 原 `/guide/errors` | 保留原三语路径 |

建议生成 `docs/errors/<code>/index.md`，让错误详情使用目录型永久 URL，减少对扩展名重写规则的依赖。VitePress 的 Markdown 文件路由和目录索引是原生能力，见 [官方路由说明](https://vitepress.dev/guide/routing)。上线仍要实测直达路由、尾斜杠以及未知码的 HTTP 404。

运行时的 docs 链接固定指向英文规范 URL；展示层切换语言。错误码与 slug 一经发布保持稳定，删除功能也保留历史错误说明。旧指南添加索引链接，不删除或重命名原入口。

### 5.2 索引页

沿用现有 VitePress 主题。首版用静态分组表格，全部码及链接包含在初始 HTML 中，本地搜索即可按错误码检索。18 个码暂不需要额外的客户端搜索组件。

```text
错误参考
根据错误码查找原因与修复步骤。

插件与加载
错误码                    问题             所属包
PLUGIN_NOT_REGISTERED     插件未注册       @docxkit/core
PLUGIN_RENDER_FAILED      插件渲染失败     @docxkit/core
...

样式 / 图片 / 表格与节点 / 导出 / 预览
```

修改 `docs/.vitepress/site.ts` 的三语 nav/sidebar，添加「错误参考」。三语 `guide/errors.md` 引用生成表格或链接到索引，避免继续手写第二份错误码列表。

### 5.3 详情页内容模板

每页固定顺序：码与标题 → 一句话说明 → 修复步骤 → 错误/正确代码 → 常见原因 → 诊断字段 → 源码与相关指南。

例如「`PLUGIN_NOT_REGISTERED`：插件未注册」：

> 文档引用了未注册的插件。导出前，先调用该插件工厂，并将返回值传给 `.use()`。

修复示例应自包含，并进入现有文档样例校验：

```ts
import { createDocx, qrcodePlugin } from 'docx-kit'

const doc = createDocx()
  .use(qrcodePlugin())
  .plugin('qrcode', { text: 'https://example.com' })

const blob = await doc.toBlob()
```

页面还要解释：JSON DSL 需要注册相应插件；`.use()` 接收插件实例而非未调用的工厂；插件名称必须匹配。不要为动态插件名机械生成 `${name}Plugin()`，第三方插件的工厂名未必遵循这个规则。

为同一码的多种原因保留细分说明。例如 `PLUGIN_LOAD_FAILED` 当前同时用于环境不支持、模块导出不正确和安全策略拦截，修复建议必须由显式 reason 参数选择；不能统一建议“重新安装”或“关闭安全策略”。

Nostics 官方也建议保持逐码 URL、初始 HTML 可读、清晰的修复步骤，以及新增码必须有文档的 CI 约束，见 [diagnostic registry](https://nostics.dev/recipes/diagnostic-registry)。

### 5.4 生成与同步

生成器读取运行时 catalog 与 `docs/errors/_meta/{en,zh-CN,zh-TW}.ts`。人工 metadata 保存 title、原因、修复步骤、示例、相关链接及首次支持版本；动态 why 不通过函数源码或正则拼成人类文档，必要时提供明确的 sample params。

生成器输出索引和详情 Markdown，并在文件头标记生成来源；详情正文从人工 metadata 生成，直接编辑产物会被覆盖。`docs/errors/_meta` 不产生页面，需在 VitePress `srcExclude` 明确排除该目录，并确认浏览器文档包没有打入 generator 或 runtime catalog。

新增项目命令：

```bash
pnpm docs:errors:generate
pnpm docs:errors:check
```

`--check` 只在内存生成并比较，不修改文件；检查缺失、过期和多余产物。放入 `docs:verify`，覆盖直接 docs 构建所需的生成状态，避免生产构建临时生成未审查文案。

页面源码链接由「所属包 + 定义位置 + 仓库 release tag」生成，不把本机路径或本机行号写入发布文档。维护分支可链接 main，发布时应与对应 release 对齐。

## 6. 首批覆盖范围

先覆盖现有 18 个已定义或已使用的不同码：

| 分类 | 错误码 |
| --- | --- |
| 导出 | `EXPORT_FAILED` |
| 图片 | `IMAGE_INVALID_DATA` |
| 样式 | `STYLE_UNKNOWN_CLASS`、`STYLE_CIRCULAR_EXTENDS` |
| 表格 | `TABLE_INVALID_COLUMNS` |
| 节点 | `UNKNOWN_NODE_TYPE`、`HEADER_FOOTER_INVALID_CONTENT` |
| 插件及清单 | `MANIFEST_INVALID`、`MANIFEST_MISSING`、`PLUGIN_DEPENDENCY_MISSING`、`PLUGIN_LOAD_FAILED`、`PLUGIN_NOT_REGISTERED`、`PLUGIN_RENDER_FAILED`、`PLUGIN_VERSION_MISMATCH` |
| 预览 | `PREVIEW_INPUT_INVALID`、`PREVIEW_FETCH_FAILED`、`PREVIEW_RENDER_FAILED`、`MICROSOFT_URL_REQUIRED` |

`PLUGIN_DEPENDENCY_MISSING` 已是公共常量，但本次未找到直接抛出点；页面应说明适用范围，不能编造触发行为。新增 STYLE/HEADER 常量和共享预览 metadata 属于补齐公开目录，不改已有字符串。

后续再为 Markdown 不支持的 token、MCP 路径问题和重复插件注册引入新码。插件 setup 原样透传、`loadAll()` 失败后继续执行、嵌套插件错误包装等行为均保留原语义；诊断迁移与行为变更分开处理。

## 7. 分阶段实施与验收

| 阶段 | 交付 | 验收重点 |
| --- | --- | --- |
| A：共享适配 | 依赖、catalog、factory、serializer、原导出兼容 | 类身份、name/code/message/cause、参数类型、静默构造 |
| B：调用点迁移 | core、loader、renderer、PDK 与相关插件 | 不改变抛错时机、包装层级、错误码或继续执行行为 |
| C：错误文档 | 三语索引与 18 × 3 详情页、生成器、导航、旧指南订正 | 目录/页面一致，示例有效，所有错误链接可直接访问 |
| D：产品展示 | Playground、Theme Studio、MCP，随后统一 warnings | 错误信息不再只剩字符串；并发互不干扰；无重复输出 |

建议最小可发布组合为 A + B + C；如果输出 docs URL，目标页面必须先于或伴随 npm 版本上线。D 可以后续交付，先保留应用原来的 message 展示。

实现阶段需要的关键验证：

- 错误包契约测试；known code 的 params 类型测试；自定义 code 的旧构造器测试。
- 保留现有插件加载、嵌套编译、预览和导出测试，检查原异常及 cause 链。
- Node 与浏览器消费 smoke test，确认无 `node:fs` 等 reporter 依赖进入浏览器入口。
- 生成器 `--check`，三语导航与详情 HTML、搜索索引、未知码 404；详情示例走文档样例检查。
- Playground 实际触发未注册插件，确认 fix 与当前语言详情链接；MCP 失败响应保留 code 和修复建议。
- 对比接入前后 errors/core/browser 入口的原始与压缩体积，记录差值，再决定是否拆分展示入口。

完整发布前沿用仓库现有命令：

```bash
pnpm build
pnpm lint
pnpm typecheck
pnpm test --run
pnpm docs:verify
pnpm apps:verify
pnpm release:artifacts
pnpm release:consumer
```

本次没有执行这些实现阶段验证，因为目前只交付方案。外部 API、发布版本及参考实现的核对记录见 [来源调研](./nostics-source-research.md)。
