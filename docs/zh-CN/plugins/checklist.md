# Checklist

生成验收、巡检、发布清单，使用 Word 原生复选框。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `items` | `ChecklistItem[]` | 按输入顺序渲染；空数组显示 `(no items)` |
| `title` | `string` | 可选标题 |

`ChecklistItem` 包含必填的 `text`，以及可选的 `checked`（默认 `false`）、`owner`（负责人）、`note`（下方备注）、`group`（分组标题）。连续相同分组只显示一次标题，分组再次出现会重新显示标题；不排序。无分组条目会重置当前分组。

支持复选框内容控件的 Word 阅读器可以切换状态；浏览器预览可能只显示初始状态。负责人和备注使用主题 `muted` 颜色，调用样式可以覆盖默认样式。

独立包：`@docxkit/plugin-checklist`，导出 `checklistPlugin`、`ChecklistOptions`、`ChecklistItem`。主包 `docx-kit` 和 `docx-kit/node` 同样导出。

## 用法

```ts
import { checklistPlugin, createDocx } from 'docx-kit'

const doc = createDocx().use(checklistPlugin()).plugin('checklist', {
  title: '发布检查',
  items: [
    { text: '代码审核', checked: true, owner: 'Alice', group: '发布前' },
    { text: '验收测试', owner: 'Bob', note: '附上测试报告', group: '发布前' },
    { text: '监控错误', group: '发布后' },
  ],
})

await doc.toBlob()
```
