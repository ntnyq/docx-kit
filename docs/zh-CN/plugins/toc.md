# TOC

插入 Word 原生目录字段。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `title` | `string` | 目录标题 |
| `maxLevel` | `number` | 最大标题层级，范围 `1..9` |

## 用法

```ts
createDocx().use(tocPlugin()).plugin('toc', {
  title: '目录',
  maxLevel: 3,
})
```

文档在 Word 中打开后，通常需要更新字段才能显示最终页码。
