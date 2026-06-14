# 包总览

这个页面按 monorepo 最新状态总结所有公开包。

## 基础运行时

| 包名 | 说明 |
| --- | --- |
| `docx-kit` | 默认总入口 |
| `@docxkit/core` | 核心文档构建与编译能力 |
| `@docxkit/types` | 共享类型 |
| `@docxkit/renderer` | 浏览器内 `.docx` 预览 |

## 扩展与工具

| 包名 | 说明 |
| --- | --- |
| `@docxkit/loader` | 从 inline、npm、URL、本地加载插件 |
| `@docxkit/pdk` | 自定义插件测试辅助 |
| `@docxkit/registry` | 基于 npm 关键字的插件发现 |
| `@docxkit/create-plugin` | 新插件脚手架 |

## AI 与 Agent

| 包名 | 说明 |
| --- | --- |
| `@docxkit/ai` | report、invoice、resume、letter 模板与 prompt 工具 |
| `@docxkit/mcp` | 面向 AI Agent 的 MCP 接口 |

## 预设与主题

- 预设：`@docxkit/preset-classic`、`@docxkit/preset-modern`、`@docxkit/preset-academic`
- 主题：`@docxkit/theme-minimal`、`@docxkit/theme-ocean`、`@docxkit/theme-warm`

## 插件包

所有内置插件包都在 [插件总览](/zh-CN/plugins/) 中逐个列出并附带独立文档页。
