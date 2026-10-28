# 包总览

本页按包展示当前 docx-kit monorepo 的运行时、工具、预设和主题，与插件参考文档互补。

## 核心运行时包

| 包 | 用途 | 使用场景 |
| --- | --- | --- |
| `docx-kit` | 聚合包 | 希望统一导入 API 的应用开发者的默认选择 |
| `@docxkit/core` | 核心构建器、编译器、DSL、样式与工具 | 需要精简运行时或构建高级集成 |
| `@docxkit/types` | 共享类型定义 | 仅需要类型的包或应用 |
| `@docxkit/markdown` | 将 Markdown 转换为文档 DSL | 使用 Markdown 输入复用文档样式、主题与插件 |
| `@docxkit/renderer` | 浏览器预览渲染 | 需要浏览器内 `.docx` 预览或嵌入 Office Online |

## 扩展与工具

| 包 | 用途 | 核心能力 |
| --- | --- | --- |
| `@docxkit/loader` | 动态加载插件 | 支持 inline、npm、URL 与本地插件来源 |
| `@docxkit/pdk` | 插件开发工具包 | 插件渲染测试与可复用测试上下文 |
| `@docxkit/registry` | 插件注册表搜索 | 发现为 docx-kit 发布的插件 |
| `@docxkit/create-plugin` | 插件脚手架 | 按约定目录结构创建插件包 |

## AI 与 Agent 包

| 包 | 用途 | 核心能力 |
| --- | --- | --- |
| `@docxkit/ai` | 适合 AI 的模板系统 | 报告、发票、简历、信函模板与提示词构建 |
| `@docxkit/mcp` | MCP 服务端 | 向 AI Agent 和助手提供文档生成工具 |

## 内置插件包

| 包 | 功能 |
| --- | --- |
| `@docxkit/plugin-badge` | 状态标签 |
| `@docxkit/plugin-barcode` | 一维条形码 |
| `@docxkit/plugin-callout` | 提示块 |
| `@docxkit/plugin-changelog` | 更新日志表格 |
| `@docxkit/plugin-code-block` | 源代码排版 |
| `@docxkit/plugin-cover-page` | 报告与文档封面 |
| `@docxkit/plugin-data-table` | 结构化对象数组表格 |
| `@docxkit/plugin-divider` | 章节分隔线 |
| `@docxkit/plugin-echarts` | 图表渲染 |
| `@docxkit/plugin-invoice` | 账单文档 |
| `@docxkit/plugin-letterhead` | 品牌信头 |
| `@docxkit/plugin-meeting-minutes` | 会议纪要 |
| `@docxkit/plugin-page-number` | 页码域 |
| `@docxkit/plugin-property-table` | 键值表格 |
| `@docxkit/plugin-qrcode` | 二维码 |
| `@docxkit/plugin-checklist` | 分组 Word 原生复选框 |
| `@docxkit/plugin-figure` | 带题注与来源的图片 |
| `@docxkit/plugin-metric-card` | 多列指标摘要 |
| `@docxkit/plugin-signature-block` | 审批与签名布局 |
| `@docxkit/plugin-timeline` | 里程碑时间线 |
| `@docxkit/plugin-toc` | 目录域 |
| `@docxkit/plugin-watermark` | 水印文字 |

每个插件的完整选项参考见[插件章节](/zh-CN/plugins/)。

## 预设包

| 包 | 样式方向 |
| --- | --- |
| `@docxkit/preset-classic` | 正式公文样式 |
| `@docxkit/preset-modern` | 适合商务文档的默认样式 |
| `@docxkit/preset-academic` | 论文与学术写作排版 |

## 主题包

| 包 | 主题方向 |
| --- | --- |
| `@docxkit/theme-minimal` | 中性简约配色 |
| `@docxkit/theme-ocean` | 蓝色主导的配色系统 |
| `@docxkit/theme-warm` | 琥珀色与暖色配色系统 |

## 推荐入口

- 应用开发选择 `docx-kit`。
- 适配器、插件或受限运行时包选择 `@docxkit/core`。
- 浏览器预览选择 `@docxkit/renderer`。
- 扩展生态选择 `@docxkit/pdk` 与 `@docxkit/create-plugin`。
- AI 驱动的文档流程选择 `@docxkit/ai` 或 `@docxkit/mcp`。
