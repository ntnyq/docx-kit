# 分隔线（Divider）

在章节之间插入横线，或从视觉上分隔内容块。

## 导入

```ts
import { dividerPlugin, type DividerOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | `'solid'` | 分隔线边框样式 |
| `color` | `string` | `'D9D9D9'` | 十六进制分隔线颜色 |
| `spacingBefore` | `number` | `200` | 分隔线段前间距 |
| `spacingAfter` | `number` | `200` | 分隔线段后间距 |

## 示例

```ts
import { createDocx, dividerPlugin } from 'docx-kit'

const doc = createDocx()
  .use(dividerPlugin())
  .h1('Quarterly Review')
  .plugin('divider', { style: 'double', color: '4472C4' })
  .p('Executive summary starts here.')
```
