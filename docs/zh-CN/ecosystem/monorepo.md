# Monorepo 结构

docx-kit 以 **pnpm workspace monorepo** 维护。文档涵盖当前工作区中的核心包、22 个插件包、3 个预设与 3 个主题。

## 目录结构

```text
docx-kit/
├── apps/
│   ├── playground/
│   └── theme-studio/
├── packages/
│   ├── ai/
│   ├── app-shared/
│   ├── core/
│   ├── create-plugin/
│   ├── docx-kit/
│   ├── errors/
│   ├── loader/
│   ├── markdown/
│   ├── mcp/
│   ├── pdk/
│   ├── registry/
│   ├── renderer/
│   └── types/
├── packages-plugins/
│   ├── badge/
│   ├── barcode/
│   ├── callout/
│   ├── changelog/
│   ├── checklist/
│   ├── code-block/
│   ├── cover-page/
│   ├── data-table/
│   ├── divider/
│   ├── echarts/
│   ├── figure/
│   ├── invoice/
│   ├── letterhead/
│   ├── meeting-minutes/
│   ├── metric-card/
│   ├── page-number/
│   ├── property-table/
│   ├── qrcode/
│   ├── signature-block/
│   ├── timeline/
│   ├── toc/
│   └── watermark/
├── packages-presets/
│   ├── academic/
│   ├── classic/
│   └── modern/
├── packages-themes/
│   ├── minimal/
│   ├── ocean/
│   └── warm/
└── docs/
```

## 工作区包分布

| 范围 | 数量 | 说明 |
| ------------------- | ----- | ------------------------------------------------------------ |
| `packages/` | 13 | 运行时、工具、AI、渲染器、Markdown 与共享类型 |
| `packages-plugins/` | 22 | 每个内置插件都有独立的包、测试与 README |
| `packages-presets/` | 3 | 常见文档类型的样式预设 |
| `packages-themes/` | 3 | 基于令牌的配色主题 |

`apps/` 中的两个私有应用独立构建和部署。[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-CN) 负责 TypeScript 编辑器和示例；[Theme Studio](https://docx-kit-theme-studio.ntnyq.dev/?lang=zh-CN) 负责主题编辑。`packages/app-shared` 是共享 UI、偏好设置和预览生命周期的私有源码包，不发布到 npm。

在仓库根目录先运行 `pnpm build`，再运行 `pnpm playground:dev` 或 `pnpm theme-studio:dev`。`pnpm apps:build` 构建两个应用，`pnpm docs:build` 仅构建文档站点。

## `packages/` 中的内容

| 包 | 职责 |
| ------------------------ | --------------------------------------------------- |
| `docx-kit` | 重新导出公共 API 的聚合包 |
| `@docxkit/core` | 构建器、编译器、样式系统与插件契约 |
| `@docxkit/loader` | 从多种来源动态加载插件 |
| `@docxkit/pdk` | 插件开发与测试辅助工具 |
| `@docxkit/registry` | 基于 npm 的插件发现 |
| `@docxkit/ai` | 模板、schema 指引与提示词构建 |
| `@docxkit/mcp` | 面向 AI Agent 集成的 MCP 服务端 |
| `@docxkit/create-plugin` | 插件脚手架 CLI |
| `@docxkit/renderer` | 浏览器端 `.docx` 预览 |
| `@docxkit/errors` | 共享错误类与错误码 |
| `@docxkit/types` | 共享的文档、样式、插件与工具类型 |

[包总览](/zh-CN/ecosystem/packages)提供各个包的详细使用建议。

## 插件打包规则

内置插件遵循统一模式：

- 包名：`@docxkit/plugin-{name}`
- 目录：`packages-plugins/{name}`
- 将 `@docxkit/core` 作为对等依赖
- 使用 `tsdown` 构建，配置 `platform: 'neutral'`
- 每个插件都有独立 README 与测试

插件可能有不同的对等依赖和运行时限制，因此保持独立包很重要。

## 预设与主题

- 预设提供 `classic`、`modern`、`academic` 等明确的基础样式。
- 主题提供 `minimal`、`ocean`、`warm` 等可复用令牌。
- 聚合包会重新导出它们，除非需要按包控制依赖，否则可统一使用 `docx-kit`。

## 贡献流程

使用 `.node-version` 指定的 Node.js LTS，以及根目录
`packageManager` 字段固定的 pnpm 版本。首次安装前启用 Corepack。
完整的审查与兼容性策略见[仓库
贡献
指南](https://github.com/ntnyq/docx-kit/blob/main/CONTRIBUTING.md)。

在仓库根目录运行：

```sh
corepack enable
corepack install
pnpm install --frozen-lockfile
pnpm run release:verify
```

常用的专项检查命令：

```sh
pnpm exec vitest run packages/core/tests/compiler/ooxml-contracts.test.ts
pnpm exec vitest run tests/e2e/plugin-lifecycle.test.ts
pnpm run visual:check
pnpm run bench:node
pnpm run bench:browser
```

视觉和性能检查的临时输出写入已忽略的 `tmp/` 与
`output/` 目录。只有确认变更符合预期后，
才提交视觉或基准测试基线。

## 文档维护要求

工作区变化时，同步更新以下三个位置：

- 首页功能摘要
- 插件与包总览
- 所有支持语言的导航

这样文档才能与实际 monorepo 保持一致。
