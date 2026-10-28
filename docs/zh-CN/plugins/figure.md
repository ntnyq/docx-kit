# 插图（Figure）

组合行内图片、可选题注与可选来源，使用核心图片编译器处理格式检测、尺寸、替代文本和图片解析。

## 导入

```ts
import { figurePlugin, type FigureOptions } from 'docx-kit'
```

独立包 `@docxkit/plugin-figure` 导出相同工厂函数与类型。

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `data` | `string \| Uint8Array \| ArrayBuffer \| Blob` | 必填 | 图片字节、Blob、data URL、原始 base64 或已解析路径 |
| `caption` | `string` | — | 图片下方的原样题注 |
| `source` | `string` | — | 题注下方的原样来源，不添加前缀 |
| `alt` | `string` | — | 无障碍图片描述 |
| `width` | `UnitValue` | 原始宽度 | 纯数字按像素处理，支持 CSS 长度单位 |
| `height` | `UnitValue` | 原始高度 | 仅提供一个维度时保持宽高比 |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'` | 自动检测 | 显式图片格式提示 |
| `alignment` | `'center' \| 'left' \| 'right'` | `'center'` | 图片、题注与来源的对齐方式 |

同时提供宽高时直接使用指定尺寸。无法读取原始尺寸时，核心图片编译器采用 300 × 180 像素的回退尺寸与宽高比。省略题注和来源不会生成额外段落。

## 示例

```ts
import { createDocx, figurePlugin } from 'docx-kit/node'

await createDocx()
  .use(figurePlugin())
  .plugin('figure', {
    data: './architecture.png',
    width: '12cm',
    alt: 'Services and their dependencies',
    caption: 'Figure 1. Service architecture',
    source: 'Source: Engineering team',
    alignment: 'center',
  })
  .save('architecture.docx')
```

浏览器中从 `docx-kit` 导入，提供字节、Blob 或 data URL，通过 `.toBlob()` 导出。本地路径需要 `docx-kit/node` 或显式 `config.resolveImage` 适配器。URL 也需要显式解析器，插件不会自动获取。详见[图片](/zh-CN/guide/images)。

图片与题注段落在后面存在题注或来源时，会请求与下一段保持同页。调用样式可以覆盖此行为和默认对齐方式。来源使用主题的 `muted` 颜色。

题注按输入原样渲染。自动编号、Word 序列域和交叉引用不在当前插件范围内。
