# Figure

组合图片、题注和来源，复用核心图片编译器。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `data` | `string \| Uint8Array \| ArrayBuffer \| Blob` | 必填；支持字节、Blob、data URL、base64 或已配置解析器的路径 |
| `caption` | `string` | 可选题注，按原文输出 |
| `source` | `string` | 可选来源，显示在题注下方，不自动添加前缀 |
| `alt` | `string` | 图片替代文本 |
| `width` / `height` | `UnitValue` | 数字按像素处理；支持 CSS 长度单位，只传一边会保持比例 |
| `imageType` | `'bmp' \| 'gif' \| 'jpeg' \| 'jpg' \| 'png'` | 默认自动检测 |
| `alignment` | `'center' \| 'left' \| 'right'` | 图片、题注与来源对齐，默认居中 |

不传尺寸时使用图片原始尺寸；无法读取时回退到 300 × 180 像素及其比例。两边都传时使用指定尺寸。题注和来源均为可选，不生成多余空段落。

浏览器从 `docx-kit` 导入，传入字节、Blob 或 data URL，并用 `.toBlob()` 导出。本地路径需要 `docx-kit/node` 或 `config.resolveImage`，URL 也需要显式解析器，不会自动下载。图片与后续题注、来源默认尽量保持同页，调用样式可以覆盖对齐等设置。来源使用主题 `muted` 颜色。

首版不生成自动编号、Word 序列域或交叉引用，需要的编号请写入 `caption`。

独立包：`@docxkit/plugin-figure`，导出 `figurePlugin`、`FigureOptions`。主包 `docx-kit` 和 `docx-kit/node` 同样导出。

## 用法

```ts
import { createDocx, figurePlugin } from 'docx-kit/node'

await createDocx().use(figurePlugin()).plugin('figure', {
  data: './architecture.png',
  width: '12cm',
  alt: '服务及其依赖关系',
  caption: '图 1. 服务架构',
  source: '来源：工程团队',
}).save('architecture.docx')
```
