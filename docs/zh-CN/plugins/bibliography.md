# 参考文献

`bibliographyPlugin` 生成数字编号文献列表和可跳转的引用段落。编号从 1 开始，按显式文献目录的顺序确定；重复引用、提前引用和重复导出保持相同编号。

```ts
import { bibliographyPlugin, type BibliographyEntry, type BibliographyOptions } from 'docx-kit'
// 独立包：import { bibliographyPlugin } from '@docxkit/plugin-bibliography'
```

独立包、`docx-kit`、`docx-kit/node` 均导出工厂和 `BibliographyConfig`、`BibliographyOptions`、`BibliographyBaseOptions`、`BibliographyListOptions`、`BibliographyCitationOptions`、`BibliographyEntry`、`BibliographyStyles`。

## 选项

| 选项 | 类型 | 默认值与行为 |
| --- | --- | --- |
| `entries` | `BibliographyEntry[]` | 沿用工厂目录，否则 `[]`；单次传入会整体替换目录，数组顺序决定编号 |
| `mode` | `'list' \| 'citation'` | 默认 `list`；文献列表或引用段落 |
| `bookmarkPrefix` | `string` | 默认 `bib_`；列表书签与引用目标的共同前缀 |
| `styles` | `BibliographyStyles` | 分部样式 |
| `title` | `string` | 仅列表模式，默认 `References`；空字符串省略标题 |
| `ids` | `string[]` | 引用模式必填；按传入顺序显示，重复 ID 去重 |
| `hyperlink` | `boolean` | 仅引用模式，默认 `true`；编号链接到列表条目 |
| `prefix`、`suffix` | `string` | 仅引用模式，默认空字符串；放在 `[1, 2]` 前后 |

工厂接受 `{ entries, styles }`。所有引用和对应列表须使用相同目录与前缀；JSON/MCP 在各节点的 `options.entries` 中传入目录。列表模式空目录、引用模式空 `ids` 均不输出任何块，包括标题。`entries: []` 可清空工厂目录。

`BibliographyEntry` 必填 `id`（唯一非空字符串）与 `title`（非空白标题）；可选 `authors: string[]`、`year`、`publisher`、`doi`、`url`，后四项均为字符串。`bookmarkPrefix + id` 须为 1–40 个 ASCII 字母、数字或下划线，以字母开头。

格式为 `[1] 作者. 标题. 年份. 出版者. DOI: 标识符. URL`，缺失字段省略，元数据按普通文字显示。首版提供简单数字格式，不包含 CSL、APA/GB/T 标准实现、Word 来源管理器或原生 `CITATION` / `BIBLIOGRAPHY` 域；修改目录后需重新生成文档。

引用输出为块级段落，通过 `prefix` / `suffix` 搭配正文，不是行内 DSL 节点。可跳转引用需要文档中存在对应列表，列表可以稍后出现。每份列表渲染一次；独立列表使用不同 `bookmarkPrefix`，避免书签重名。

缺失或类型不正确的引用 ID 数组、未知引用 ID、重复来源 ID、非法书签名称和空白标题在渲染时抛错；编译器包装为 `PLUGIN_RENDER_FAILED`。

## 平台与样式

支持 Node.js 和浏览器，仅需常规的 `@docxkit/core`、`docx` 对等依赖，无额外依赖。浏览器导出用 `.toBlob()`，Node 文件保存从 `docx-kit/node` 调用 `.save()`。预览显示普通文字与内部链接，无需更新 Word 域。

文字部位：`citation`、`entry`、`title`；段落部位：`citationParagraph`、`entryParagraph`、`titleParagraph`。标题使用主题 `primary`，文献条目默认悬挂缩进。

工厂 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。优先级：内置/主题默认 → 工厂样式 → 通用节点样式 → 单次部位样式。支持单位和主题 token；文字、段落、单元格分别使用 `TextStyleRule`、`ParagraphStyleRule`、`Omit<CellStyleRule, 'height'>`。

## 完整示例

[Playground](https://docx-kit-playground.ntnyq.dev/?lang=zh-CN&example=references)

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
