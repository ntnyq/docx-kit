# Monorepo 结构

仓库采用 pnpm workspace monorepo，当前文档已按最新结构覆盖以下目录：

- `packages/`：10 个基础包，含 `renderer` 与 `types`
- `packages-plugins/`：18 个内置插件包
- `packages-presets/`：3 个样式预设
- `packages-themes/`：3 个主题包

## 核心包

| 包名 | 作用 |
| --- | --- |
| `docx-kit` | 总入口，重导出公共 API |
| `@docxkit/core` | builder、compiler、样式系统、插件契约 |
| `@docxkit/loader` | 动态插件加载 |
| `@docxkit/pdk` | 插件开发与测试辅助 |
| `@docxkit/registry` | 插件发现 |
| `@docxkit/ai` | AI 模板与 prompt 构建 |
| `@docxkit/mcp` | MCP 服务端 |
| `@docxkit/create-plugin` | 插件脚手架 |
| `@docxkit/renderer` | 浏览器预览 |
| `@docxkit/types` | 共享类型定义 |

详细说明见 [包总览](/zh-CN/ecosystem/packages)。
