# Page Number

输出页码字段，通常用于页眉或页脚。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `alignment` | `AlignmentType` | 对齐方式 |
| `fontSize` | `number` | 字号，单位为 points |
| `showTotal` | `boolean` | 是否显示 `Page X of Y` |

## 用法

```ts
createDocx().use(pageNumberPlugin()).plugin('pageNumber', {
  alignment: 'center',
  showTotal: true,
})
```
