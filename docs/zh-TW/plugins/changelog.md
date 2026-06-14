# Changelog

輸出版本變更表，適合 release notes 或發布記錄。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `title` | `string` | 標題，預設 `Changelog` |
| `entries` | `ChangelogEntry[]` | 版本條目陣列 |

`ChangelogEntry` 包含 `version`、`date`、`type`、`changes` 四個欄位，`type` 支援 `added`、`changed`、`fixed`、`removed`。

## 用法

```ts
createDocx().use(changelogPlugin()).plugin('changelog', {
  entries: [
    { version: '1.0.0', date: '2026-06-14', type: 'added', changes: '新增中文文件。' },
  ],
})
```
