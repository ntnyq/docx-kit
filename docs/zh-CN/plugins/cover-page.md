# Cover Page

生成报告、方案书、白皮书常见的封面页。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `title` | `string` | 主标题，必填 |
| `subtitle` | `string` | 副标题 |
| `author` | `string` | 作者或部门 |
| `organization` | `string` | 机构名称 |
| `date` | `string` | 日期 |
| `backgroundColor` | `string` | 标题背景色 |
| `showRule` | `boolean` | 是否显示装饰线 |

## 用法

```ts
createDocx().use(coverPagePlugin()).plugin('coverPage', {
  title: '年度报告',
  subtitle: '2026',
  author: '战略发展部',
})
```
