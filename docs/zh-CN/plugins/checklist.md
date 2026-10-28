# Checklist

生成验收、巡检、发布清单，默认使用 GitHub 风格的图形复选框。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `items` | `ChecklistItem[]` | 按输入顺序渲染；空数组显示 `(no items)` |
| `checkboxStyle` | `'github' \| 'native'` | 默认 `'github'`，显示静态圆角图形；`'native'` 使用可切换状态的 Word 原生控件 |
| `title` | `string` | 可选标题 |

`ChecklistItem` 包含必填的 `text`，以及可选的 `checked`（默认 `false`）、`owner`（负责人）、`note`（下方备注）、`group`（分组标题）。连续相同分组只显示一次标题，分组再次出现会重新显示标题；不排序。无分组条目会重置当前分组。

分组标题默认加粗，使用主题的 `accent` 颜色（回退到 `primary`），字号约为正文的 120%，组间留白随正文字号缩放。可选的清单总标题使用 `primary`，字号约为正文的 140%，保持高于分组的层级。调用时显式设置的文字颜色可覆盖这些默认颜色。

默认的 `'github'` 样式使用 SVG 图形，并提供高分辨率 PNG 回退。选中框使用主题的 `accent` 颜色（缺省时回退到 `primary`），搭配白色对勾；未选中框的边线使用 `muted`（回退到 `border`），底色使用 `background`。两种图片格式使用相同的主题颜色。图形复选框会随任务的实际字号缩放，并与文字视觉居中对齐；换行文字与备注保持相同的文字起点。这些是静态图片，不依赖符号字体。

设置 `checkboxStyle: 'native'` 后，支持复选框内容控件的 Word 阅读器可以切换状态；浏览器预览可能只显示初始状态。负责人和备注使用主题 `muted` 颜色，调用样式可以覆盖默认的文字和段落样式。备注字号随任务正文缩放，约为正文的 80%，按 Word 的半磅精度取整。负责人字号约为正文的 85%，即使正文显式设置了颜色，负责人仍保持主题的 `muted` 颜色。

独立包：`@docxkit/plugin-checklist`，导出 `checklistPlugin`、`ChecklistOptions`、`ChecklistItem`。主包 `docx-kit` 和 `docx-kit/node` 同样导出。

## 完整示例

[在 Playground 中打开完整示例](/zh-CN/playground?example=checklist)，可直接修改代码、预览效果并下载 DOCX。

这份 Ocean 主题示例包含发布前后两个分组、负责人、备注，以及已勾选和未勾选的任务。正文为 12pt，分组标题为 14.5pt，负责人为 10pt，备注为 9.5pt。

![包含分组标题、负责人和备注的完整 Checklist 效果](/examples/checklist.png)

<<< @/.vitepress/constants/templates/checklist.ts#checklist-example
