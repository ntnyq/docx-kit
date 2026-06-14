# Callout

用于信息、警告、成功、危险提示块。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | 提示类型 |
| `title` | `string` | 可选标题 |
| `content` | `string` | 正文内容 |

## 用法

```ts
createDocx().use(calloutPlugin()).plugin('callout', {
  type: 'warning',
  title: '注意',
  content: '此操作不可撤销。',
})
```
