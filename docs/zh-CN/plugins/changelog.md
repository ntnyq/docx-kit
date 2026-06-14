# Changelog

输出版本变更表，适合 release notes 或发布记录。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `title` | `string` | 标题，默认 `Changelog` |
| `entries` | `ChangelogEntry[]` | 版本条目数组 |

`ChangelogEntry` 包含 `version`、`date`、`type`、`changes` 四个字段，`type` 支持 `added`、`changed`、`fixed`、`removed`。

## 用法

```ts
createDocx().use(changelogPlugin()).plugin('changelog', {
  entries: [
    { version: '1.0.0', date: '2026-06-14', type: 'added', changes: '新增中文文档。' },
  ],
})
```
