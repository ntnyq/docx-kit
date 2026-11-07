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
| `styles` | `CalloutStyles` | — | 本次调用的部位样式 |

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

## 样式定制

给插件工厂传入 `{ styles }`，可设置当前插件实例的默认样式；通过 `options.styles` 覆盖单次调用。插件包和 `docx-kit` 的 Node/browser 入口均导出 `CalloutConfig` / `CalloutStyles` 与 `DataTableConfig` / `DataTableStyles`。

优先级从低到高：**内置默认值（含主题取值）→ 初始化 `styles` → 调用通用 `style`（`.plugin()` 第三个参数）→ 调用部位 `styles`**。缺省或 `undefined` 字段沿用默认值；嵌套普通对象按字段合并，数组整体替换，`0` 和 `false` 均生效。每次调用独立解析，不修改配置对象。单位和 `$colors.primary`、`$fontSize.lg`、`$spacing.md` 等主题 token 复用现有样式编译器。

| 部位 | 样式类型 | 作用范围 |
| --- | --- | --- |
| `cell` | `Omit<CellStyleRule, 'height'>` | 容器背景、边框、宽度及内边距 |
| `paragraph` | `ParagraphStyleRule` | 共享段落的对齐、间距及分页控制 |
| `title` | `TextStyleRule` | 标题文字 |
| `content` | `TextStyleRule` | 正文文字 |

```ts
import { calloutPlugin, createDocx } from 'docx-kit/node'

const doc = createDocx({
  theme: { colors: { primary: '#2563EB' }, spacing: { md: 8 } },
})
  .use(calloutPlugin({
    styles: {
      cell: { margin: '8pt 10pt', borderLeft: { width: 3 } },
      title: { fontSize: 16, color: '$colors.primary' },
      content: { fontSize: 11 },
    },
  }))
  .plugin('callout', {
    type: 'info',
    title: 'Note',
    content: 'Custom body text',
    styles: {
      cell: { marginBottom: 0 },
      title: { bold: false, color: '#DC2626' },
      paragraph: { lineHeight: 1.5 },
    },
  }, { fontFamily: 'Arial' })
```

### JSON DSL、AI 与 MCP

JSON DSL 中在调用的 `options.styles` 内填写与 Builder 相同的可序列化部位样式。AI 工具定义和 schema 指南已包含这些选项；MCP 通过 `docx-kit://schema` 和 `get_plugin_help().optionsSchema` 暴露对应结构。`validate_schema` 会拒绝未知部位、放错部位的属性及错误的样式值。其他内置插件也提供对应的部位样式契约。

```json
{
  "type": "plugin",
  "name": "callout",
  "options": {
    "type": "info",
    "title": "Note",
    "content": "Body",
    "styles": { "title": { "color": "#2563EB" }, "cell": { "margin": 0 } }
  }
}
```
