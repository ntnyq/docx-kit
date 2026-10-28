# 提示块（Callout）

带图标的彩色信息、警告、成功和危险提示框。

## 导入

```ts
import { calloutPlugin, type CalloutOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `content` | `string` | _必填_ | 提示块正文 |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | _必填_ | 提示块样式，控制图标、颜色与色调 |
| `title` | `string` | — | 正文前的可选加粗标题 |

### 类型预设

| 类型 | 图标 | 背景色 | 边框颜色 |
|------|------|------------|--------------|
| `info` | ℹ️ | `#D6E4F0` | `#4472C4` |
| `warning` | ⚠️ | `#FFF2CC` | `#FFC000` |
| `success` | ✅ | `#E2F0D9` | `#70AD47` |
| `danger` | ⚠️ | `#FCE4D6` | `#FF0000` |

## 示例

### 信息提示

```ts
import { createDocx, calloutPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(calloutPlugin())
  .h1('System Maintenance')
  .plugin('callout', {
    type: 'info',
    content: 'The system will be upgraded tonight at 22:00. No downtime is expected.',
  })
  .save('callout.docx')
```

### 带标题的警告

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Important Notice')
  .plugin('callout', {
    type: 'warning',
    title: '注意',
    content: '此操作不可撤销，请确认后再提交。',
  })
  .save('warning.docx')
```

### 成功确认

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Deployment Status')
  .plugin('callout', {
    type: 'success',
    title: 'Deployment Successful',
    content: 'Version 2.4.1 has been deployed to production. All health checks passed.',
  })
  .save('success.docx')
```

### 危险提醒

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Security Alert')
  .plugin('callout', {
    type: 'danger',
    title: 'Critical Vulnerability',
    content: 'A critical security vulnerability (CVE-2026-XXXX) has been detected. Immediate action required.',
  })
  .save('danger.docx')
```

### 多个提示块

```ts
const doc = createDocx()
  .use(calloutPlugin())
  .h1('Weekly Summary')

  .plugin('callout', {
    type: 'success',
    title: 'Completed',
    content: 'Sprint goals met on time. 12 features shipped.',
  })

  .plugin('callout', {
    type: 'warning',
    title: 'At Risk',
    content: 'API rate limiting needs attention before next sprint.',
  })

  .plugin('callout', {
    type: 'info',
    title: 'Upcoming',
    content: 'Team retro scheduled for Friday at 3 PM.',
  })

  .save('multiple-callouts.docx')
```
