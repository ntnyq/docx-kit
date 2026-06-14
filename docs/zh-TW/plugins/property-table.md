# Property Table

用於展示屬性名與屬性值的雙欄表格。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `items` | `{ key: string; value: string }[]` | 屬性列表 |
| `keyBold` | `boolean` | Key 欄是否加粗 |
| `striped` | `boolean` | 是否使用斑馬紋 |

## 用法

```ts
createDocx().use(propertyTablePlugin()).plugin('propertyTable', {
  items: [{ key: 'Author', value: 'Jane' }],
})
```
