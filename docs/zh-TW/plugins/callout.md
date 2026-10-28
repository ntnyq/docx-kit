# 提示區塊（Callout）

帶圖標的彩色資訊、警告、成功和危險提示框。

## 匯入

```ts
import { calloutPlugin, type CalloutOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `content` | `string` | _必填_ | 提示區塊正文 |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | _必填_ | 提示區塊樣式，控制圖標、顏色與色調 |
| `title` | `string` | — | 正文前的可選加粗標題 |

### 型別預設

| 型別 | 圖標 | 背景色 | 邊框顏色 |
|------|------|------------|--------------|
| `info` | ℹ️ | `#D6E4F0` | `#4472C4` |
| `warning` | ⚠️ | `#FFF2CC` | `#FFC000` |
| `success` | ✅ | `#E2F0D9` | `#70AD47` |
| `danger` | ⚠️ | `#FCE4D6` | `#FF0000` |

## 範例

### 資訊提示

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

### 帶標題的警告

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

### 成功確認

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

### 危險提醒

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

### 多個提示區塊

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
