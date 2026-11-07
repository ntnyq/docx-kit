# 目录（Table of Contents）

插入根据标题层级生成的 Word 目录域。

## 导入

```ts
import { tocPlugin, type TocOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `title` | `string` | `'Contents'` | 可见章节标题 |
| `maxLevel` | `number` | `3` | 纳入目录的最大标题层级，限制在 `1..9` |

## 示例

```ts
import { createDocx, tocPlugin } from 'docx-kit'

const doc = createDocx()
  .use(tocPlugin())
  .plugin('toc', { title: 'Document Contents', maxLevel: 4 })
```

## 说明

- 插件输出真正的 DOCX 目录域，而非静态文本快照。
- Microsoft Word 打开文件后可能提示更新域。

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `TocConfig` 和 `TocStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `title` | `TextStyleRule` | 文字排版 |
| `titleParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

此处仅定制显式标题。更新目录域时，Word 使用文档的 TOC 样式生成目录条目。

```ts
tocPlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
