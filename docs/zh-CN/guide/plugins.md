# 插件

插件是 docx-kit 的主要扩展机制，无需修改核心编译器即可添加新节点类型，同时保留统一的构建流程与类型安全的选项对象。

## 工作方式

每个内置或自定义插件都遵循相同生命周期：

1. 导入插件工厂函数或插件实例。
2. 使用 `.use(...)` 注册。
3. 使用 `.plugin(name, options)` 调用。

```ts
import { calloutPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(calloutPlugin())
  .plugin('callout', {
    type: 'success',
    title: 'Published',
    content: 'The release notes were exported successfully.',
  })
```

## 内置插件范围

当前工作区包含 **22 个内置插件**：

- 内容块：`badge`、`callout`、`checklist`、`codeBlock`、`divider`、`watermark`
- 报告结构：`coverPage`、`letterhead`、`toc`、`pageNumber`
- 结构化业务数据：`dataTable`、`metricCard`、`propertyTable`、`timeline`、`invoice`、`meetingMinutes`、`changelog`
- 媒体与嵌入：`barcode`、`echarts`、`figure`、`qrcode`
- 审批流程：`signatureBlock`

完整参考见[插件章节](/zh-CN/plugins/)。

## 包结构

每个内置插件都在 `packages-plugins/` 下作为独立包发布：

```text
packages-plugins/
  badge/
  barcode/
  callout/
  changelog/
  checklist/
  code-block/
  cover-page/
  data-table/
  divider/
  echarts/
  figure/
  invoice/
  letterhead/
  meeting-minutes/
  metric-card/
  page-number/
  property-table/
  qrcode/
  signature-block/
  timeline/
  toc/
  watermark/
```

这样可以独立维护插件的发布、测试、对等依赖与文档。

## 对等依赖

部分插件将体积较大的运行时依赖设为可选：

| 插件 | 可选原因 |
| --- | --- |
| `@docxkit/plugin-barcode` | 仅生成条形码图片时需要 `bwip-js` |
| `@docxkit/plugin-code-block` | 仅语法着色时需要 `highlight.js` |
| `@docxkit/plugin-echarts` | 仅使用图表渲染时需要 `echarts` |
| `@docxkit/plugin-qrcode` | 仅生成二维码图片时需要 `qrcode` |

缺少可选对等依赖时，插件会优雅降级，或抛出带有解决建议的运行时错误。

## 自定义插件

需要业务或文档流程专属节点类型时，使用 `definePlugin()`。

```ts
import { definePlugin } from 'docx-kit'
import { Paragraph } from 'docx'

interface DisclaimerOptions {
  text: string
}

const disclaimerPlugin = definePlugin<'disclaimer', DisclaimerOptions>({
  name: 'disclaimer',
  render(options) {
    return new Paragraph(options.text)
  },
})
```

好的自定义插件应当：

- 接收小而明确的选项对象
- 返回常规 `docx` 节点或节点数组
- 在渲染函数附近完成校验
- 避免隐藏无关的文档修改

## 插件配套工具

- `@docxkit/create-plugin`：创建插件包脚手架
- `@docxkit/pdk`：插件测试辅助工具
- `@docxkit/loader`：从 inline、npm、URL 或本地来源加载插件
- `@docxkit/registry`：发现带 docx-kit 关键字的插件
- `@docxkit/mcp`：向 AI Agent 提供支持插件的文档生成工具

各部分所在位置见 [Monorepo 结构](/zh-CN/ecosystem/monorepo)与[包总览](/zh-CN/ecosystem/packages)。
