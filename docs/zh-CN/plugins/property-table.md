# Property Table

用于展示属性名与属性值的双列表格。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `items` | `{ key: string; value: string }[]` | 属性列表 |
| `keyBold` | `boolean` | Key 列是否加粗 |
| `striped` | `boolean` | 是否使用斑马纹 |

## 用法

```ts
createDocx().use(propertyTablePlugin()).plugin('propertyTable', {
  items: [{ key: 'Author', value: 'Jane' }],
})
```
