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
