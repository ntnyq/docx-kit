# 交叉引用

`crossReferencePlugin` 生成引用段落，使用 Word 原生 `REF` / `PAGEREF` 域，指向标题、题注或其他文字上的现有书签。

```ts
import { crossReferencePlugin, type CrossReferenceOptions } from 'docx-kit'
// 独立包：import { crossReferencePlugin } from '@docxkit/plugin-cross-reference'
```

独立包、`docx-kit`、`docx-kit/node` 均导出工厂和 `CrossReferenceOptions`、`CrossReferenceConfig`、`CrossReferenceStyles`。

## 选项

| 选项 | 类型 | 默认值与行为 |
| --- | --- | --- |
| `target` | `string` | 必填；1–40 个 ASCII 字母、数字或下划线，以字母开头 |
| `kind` | `'text' \| 'number' \| 'page'` | 默认 `text`；分别使用 `REF`、带 `\w` 完整编号上下文的 `REF`、`PAGEREF` |
| `fallback` | `string` | 默认 `target`；更新域之前显示的缓存文字 |
| `hyperlink` | `boolean` | 默认 `true`；添加 `\h` 跳转到目标 |
| `prefix`、`suffix` | `string` | 默认空字符串；域前后的文字 |
| `styles` | `CrossReferenceStyles` | 分部样式 |

用 `.bookmark('overview', ['概述'])` 创建目标，也可以在段落内放置书签节点。`number` 要求书签位于真正的编号段落中。普通题注文字中的数字不是自动段落编号；这种情况可只为所需文字添加书签，再用 `text` 引用。

插件只生成引用，不创建或检查目标是否存在。支持引用后面的书签；同一文档中不要重复使用书签名。空名称、非法名称和未知 `kind` 在渲染时抛错；目标不存在时，Word 更新域会显示引用错误。

## 平台与更新域

支持 Node.js 和浏览器，仅需常规的 `@docxkit/core`、`docx` 对等依赖，无额外依赖。通过 `.use(crossReferencePlugin()).plugin('crossReference', options)` 使用；浏览器导出用 `.toBlob()`，Node 文件保存从 `docx-kit/node` 调用 `.save()`。

输出是块级段落。浏览器预览显示 `fallback`，不计算页码或编号。需要在 Word 中更新域以刷新目标文字、编号和页码；具体结果取决于阅读器支持。

## 样式

文字部位：`text`；段落部位：`paragraph`。缓存结果和前后文字都支持样式。

工厂 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。优先级：内置/主题默认 → 工厂样式 → 通用节点样式 → 单次部位样式。支持单位和主题 token；文字、段落、单元格分别使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整示例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-CN&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
