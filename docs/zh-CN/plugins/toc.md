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
