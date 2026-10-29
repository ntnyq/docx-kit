# nostics 官方资料与 x18n 错误文档调研

核对日期：2026-10-29。本文仅记录依赖能力、外部来源与限制；docx-kit 的具体改造见接入方案。本次没有安装依赖或实施任何产品改造。

## 版本与依赖

本次查询 npm `latest` 为 **nostics 1.3.0**，并核对了发布 tarball 的 `dist/index.mjs` 与 `v1.3.0` 标签源码。包为 MIT、ESM only，根入口为 `.mjs`，声明为 `.d.mts`，包清单没有运行时 `dependencies`。这适合 ESM 库，但实际包体增长仍须在 docx-kit 构建后测量，不能把零依赖等同于零成本。[npm 元数据](https://registry.npmjs.org/nostics/1.3.0)、[版本源码](https://github.com/vercel-labs/nostics/tree/v1.3.0)

## 对错误模型的影响

`Diagnostic` 继承 `Error`。1.3.0 的实例同时具有 `code` 与 `name`，两者初始值均为诊断码；`message` 保存解释，`why` 是它的 getter。`fix`、`docs`、`sources`、`data` 为附加信息，`cause` 交给原生 Error 构造器保存。官网 API 表未列出 `code`，以已发布代码核对结果为准。[Diagnostic 1.3.0 源码](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/diagnostic.ts)

这意味着直接用 `Diagnostic` 替换现有 `DocxKitError` 会影响 `instanceof DocxKitError`、`name === 'DocxKitError'` 及构造签名。接入时应保留旧类和旧错误码，用适配层承接新元数据。不要为了模仿示例而给全部错误重新编号；nostics 接受任意字符串作为代码。[错误码约定](https://nostics.dev/guide/defining-diagnostics#codes)

`toJSON()` 输出 `name/why/fix/docs/sources/cause/data/stack`，没有单独的 `code/message` 字段，且 `cause` 未递归规范化，不能直接当作既有 MCP 或 HTTP 协议的替代。若传输层需要 `code/message`，应显式映射；如需可靠序列化，另外处理循环引用和 Error 的不可枚举字段。[JSON 类型与实现](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/diagnostic.ts)

## 构造、报告与类型

`defineDiagnostics` 将每个代码变为函数：调用后创建新实例，执行已配置 reporters，再返回实例；`throw` 是调用者的控制流。默认 reporters 是空数组，因此可以静默使用。也可直接 `new Diagnostic(init)`，但那会绕过目录的参数类型推导。没有额外的 `.create()` 静默 API。reporter 调用循环没有统一异常隔离，不能把业务可靠性寄托在任意外部 reporter 上。[实现](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/diagnostic.ts)

动态 `why/fix/data` 函数的参数类型合并为调用参数；`cause/sources` 在调用处传入。文档链接可用 `docsBase` 字符串或解析函数生成，单个定义可覆盖或关闭。代码、运行时文本、文档 URL 因此可集中维护。[参数与文档配置](https://nostics.dev/guide/defining-diagnostics)

根入口只依赖内部 JS，未导入 Node 内置模块；文件 reporter 独立在 `nostics/reporters/node` 并导入 `node:fs`。浏览器入口应避免重导出 Node reporter。console reporter 需显式配置，默认打印到 `console.warn`；这不意味着默认目录会自动打印。[根入口](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/index.ts)、[Node reporter](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/reporters/node.ts)、[Reporter 文档](https://nostics.dev/guide/reporters-and-formatters)

## 生产构建的边界

`defineProdDiagnostics` 是精简目录：保留代码、docs 与调用处 cause/sources，去掉目录 why/fix/data。它的 message 为 docs URL 或空字符串，默认也不报告。选用它会改变生产环境可观察字段，不能仅作为包体优化无条件开启。[生产目录源码](https://github.com/vercel-labs/nostics/blob/v1.3.0/src/prod-diagnostics.ts)

`@nostics/unplugin` 是独立构建依赖。strip 插件针对仅报告的语句，保留被 throw、return、赋值或作为参数使用的调用；与精简目录解决的是不同问题。文档给出 tsdown 的 Rolldown 适配方式，并要求保持 nostics external。跨文件分析只跟随一层相对导入，复杂包装、深层 barrel 或当前库的适配器不一定能识别。首期可以保留完整错误信息，后续实测产物再决定是否引入 strip。[生产构建说明](https://nostics.dev/guide/production-builds)

## x18n 页面与生成流程

参考站的索引实际包含 **Code / Title / Exit code** 表，按代码数字分组，链接到独立详情页。详情页包含 **Message、What to do、Exit code、Source**，其中修复步骤靠前，来源链接定位抛错文件。这是 CLI 的模型，docx-kit 应把退出码栏目改成分类或所属包。[索引](https://x18n.dev/errors)、[详情示例](https://x18n.dev/errors/x18n1001)

x18n 在 core 中集中定义动态 why 与静态 fix，使用 nostics 且不配置 reporters，由消费端输出捕获的错误。错误类型收窄了 `name/code`，对外 JSON 另设转换函数，主动排除 stack/cause/data。这是可以借鉴的分层方式。[x18n 错误实现](https://github.com/x18n-dev/x18n/blob/main/packages/core/src/errors.ts)

站点由自己的 `docs/scripts/generate.ts` 生成，并非 nostics 内置文档生成器：它把错误定义与 `notes.ts` 的标题、处理说明组合，生成索引和每码页，扫描源码得到 Source 链接。生成器检查缺少或多余的 notes；`--check` 比较已提交文件并检测过期页面。docx-kit 可采用同样的“运行时目录 + 文档内容 + 生成器 + CI 校验”流程。[生成器](https://github.com/x18n-dev/x18n/blob/main/docs/scripts/generate.ts)、[手写说明](https://github.com/x18n-dev/x18n/blob/main/docs/scripts/notes.ts)

nostics 的 registry recipe 仅提供 URL、页面内容与 CI 一致性检查的建议；并不提供 x18n 风格的网站组件。它建议静态 HTML 可读、未知代码返回 404，并长期维护已发布 URL。[官方 registry recipe](https://nostics.dev/recipes/diagnostic-registry)

## 本次未验证的部分

- 已核对发布代码，但未将 nostics 安装进 docx-kit，也未测量实际 bundle 增量。
- 已获取参考站 Markdown 与生成源码，未做浏览器视觉验收；不对页面的视觉样式作推断。
- Node、浏览器、Web Worker 与 MCP 的实际兼容性，需要在实现阶段通过本仓库的消费端测试与打包检查确认。
- 线上文档与 GitHub main 可继续变化；实现应固定依赖版本并重新核对相关 API。
