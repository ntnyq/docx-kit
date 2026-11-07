# 术语与缩略语表

`glossaryPlugin` 生成术语、全称、解释三列表格，支持排序与分组。

```ts
import { glossaryPlugin, type GlossaryEntry, type GlossaryOptions } from 'docx-kit'
// 独立包：import { glossaryPlugin } from '@docxkit/plugin-glossary'
```

独立包、`docx-kit`、`docx-kit/node` 均导出工厂和 `GlossaryOptions`、`GlossaryEntry`、`GlossaryConfig`、`GlossaryStyles`。

## 选项

| 选项 | 类型 | 默认值与行为 |
| --- | --- | --- |
| `entries` | `GlossaryEntry[]` | 必填；空数组不输出任何块，包括标题 |
| `title` | `string` | 默认 `Glossary`；空字符串省略标题 |
| `sort` | `'term' \| 'none'` | 默认 `term`，不区分大小写排序术语与组名；`none` 保留首次出现顺序 |
| `locale` | `string` | 默认 `en`；传给 `Intl.Collator` 的排序语言 |
| `groupBy` | `'group' \| 'none'` | 默认 `none`；按组分别生成标题和表格 |
| `labels` | `{ term?, fullForm?, definition? }` | 字符串列名，默认 `Term`、`Full form`、`Definition` |
| `styles` | `GlossaryStyles` | 分部样式 |

`GlossaryEntry` 必填 `term`（术语或缩写）和 `definition`（解释），两项均不得为空白；可选 `fullForm`（全称）和 `group`（组名），均为字符串。没有全称时保留空白单元格。

输入数组不会被修改。重复术语保留为独立行；相同排序结果保持输入顺序。开启分组后，不相邻的同组条目也会合并。默认排序时无组名的条目排在最前，且不显示组标题；`sort: 'none'` 时组按首次出现顺序，组内条目保持原序。表头跨页重复，长行允许跨页。

空白术语或解释、非法 `sort` / `groupBy` 在渲染时抛错；非法排序语言由 `Intl.Collator` 抛出 `RangeError`。编译器包装为 `PLUGIN_RENDER_FAILED`。

## 平台与样式

支持 Node.js 和浏览器，仅需常规的 `@docxkit/core`、`docx` 对等依赖，无额外依赖。浏览器导出用 `.toBlob()`，Node 文件保存从 `docx-kit/node` 调用 `.save()`。预览支持标题和表格，最终分页由阅读器决定。

文字部位：`definition`、`fullForm`、`group`、`headerText`、`term`、`title`；段落部位：`bodyParagraph`、`groupParagraph`、`headerParagraph`、`titleParagraph`；单元格部位：`bodyCell`、`headerCell`。标题和组标题使用主题 `primary`，表头底色使用 `surface`，边框使用 `border`。

工厂 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。优先级：内置/主题默认 → 工厂样式 → 通用节点样式 → 单次部位样式。支持单位和主题 token；文字、段落、单元格分别使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整示例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-CN&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
