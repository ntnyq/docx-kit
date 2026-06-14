# Badge

用于生成 `DRAFT`、`APPROVED` 这类短状态标签。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `text` | `string` | 标签文本，必填 |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | 预设语义色或自定义文字颜色 |
| `backgroundColor` | `string` | 自定义背景色 |

## 用法

```ts
createDocx().use(badgePlugin()).plugin('badge', {
  text: 'DRAFT',
  color: 'warning',
})
```
