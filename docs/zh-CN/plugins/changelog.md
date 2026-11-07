# 更新日志（Changelog）

生成包含标题与四列版本表格的发布说明章节。

## 导入

```ts
import { changelogPlugin, type ChangelogOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `entries` | `ChangelogEntry[]` | 必填 | 版本记录列表 |
| `title` | `string` | `'Changelog'` | 表格上方的标题 |

`ChangelogEntry` 字段：

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `version` | `string` | 版本标签 |
| `date` | `string` | 发布日期 |
| `type` | `'added' \| 'changed' \| 'fixed' \| 'removed'` | 带有独立高亮颜色的记录分类 |
| `changes` | `string` | 可读的变更摘要 |

## 示例

```ts
import { changelogPlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(changelogPlugin())
  .plugin('changelog', {
    title: 'Release Notes',
    entries: [
      { version: '1.4.0', date: '2026-06-14', type: 'added', changes: 'Added zh-TW docs locale.' },
      { version: '1.4.1', date: '2026-06-15', type: 'fixed', changes: 'Corrected plugin sidebar links.' },
    ],
  })
```

## 说明

- `entries` 为空时仍渲染标题，并显示居中的 `(no entries)` 占位内容。
- 适合发布说明、迁移日志或合规变更历史。

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `ChangelogConfig` 和 `ChangelogStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `bodyText`, `emptyText`, `headerText`, `title` | `TextStyleRule` | 文字排版 |
| `bodyParagraph`, `emptyParagraph`, `headerParagraph`, `titleParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |
| `bodyCell`, `headerCell` | `Omit<CellStyleRule, 'height'>` | 单元格背景、边框、宽度和内边距 |

```ts
changelogPlugin({
  styles: { headerText: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
