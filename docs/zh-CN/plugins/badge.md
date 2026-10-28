# 徽章（Badge）

渲染 `DRAFT`、`APPROVED`、`INTERNAL` 等简短状态标签。

## 导入

```ts
import { badgePlugin, type BadgeOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `string` | 必填 | 徽章文字 |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | `'neutral'` | 预设色调名称或自定义文字颜色 |
| `backgroundColor` | `string` | 由预设决定 | 自定义徽章背景色 |

## 示例

```ts
import { badgePlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(badgePlugin())
  .plugin('badge', {
    text: 'DRAFT',
    color: 'warning',
  })
```

## 说明

- 预设提供配套的前景色与背景色。
- 自定义 `color` 若不属于预设名称，会按字面文字颜色处理。
