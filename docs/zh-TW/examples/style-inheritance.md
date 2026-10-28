# 範例：樣式繼承

展示 `defineStyles()` 的 `extends` 功能如何實現 CSS 風格樣式繼承，以及基本樣式如何層疊到派生樣式。

## 完整程式碼

```ts
import { createDocx, defineStyles } from 'docx-kit/node'

// 1. Define styles with extends (inheritance chain)
const styles = defineStyles({
  // ─── Base styles (no extends) ────────────────────────────────────
  baseText: {
    fontSize: 11,
    fontFamily: 'Calibri',
    color: '#333333',
    lineHeight: 1.5,
  },

  // ─── Inherit from baseText, override selectively ───────────────────
  body: {
    extends: 'baseText',
    marginTop: 4,
    marginBottom: 4,
  },

  // ─── Inherit from body, add heading traits ────────────────────────
  h2: {
    extends: 'body',
    fontSize: 18,
    fontWeight: 'bold',
    color: '#1a1a2e',
    marginTop: 16,
    marginBottom: 8,
  },

  // ─── Inherit from h2, but smaller ───────────────────────────────
  h3: {
    extends: 'h2',
    fontSize: 14,
    marginTop: 12,
    marginBottom: 6,
  },

  // ─── Multiple inheritance (array) ────────────────────────────────
  // Inherits from both 'baseText' and 'callout' — later keys win
  warning: {
    extends: ['baseText', 'callout'],
    color: '#92400e',
    border: { style: 'single', width: 1, color: '#f59e0b' },
  },

  // ─── Standalone style ─────────────────────────────────────────────
  callout: {
    backgroundColor: '#fef3c7',
    padding: 8,
    borderRadius: 4,
  },

  // ─── Table cell that extends baseText ─────────────────────────────
  cell: {
    extends: 'baseText',
    fontSize: 10,
    verticalAlign: 'middle',
  },
  tableHeader: {
    extends: 'cell',
    fontWeight: 'bold',
    backgroundColor: '#1e293b',
    color: '#ffffff',
  },
})

// 2. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
})

interface FeatureRow {
  feature: string
  inherits: string
  resolvesTo: string
}

doc
  .h1('Style Inheritance via extends')
  .p(
    'Styles can extend other styles using the extends keyword. ' +
    'The inheritance chain is resolved at compile time — child styles ' +
    'inherit all parent properties and can override selectively.',
    { className: 'body' },
  )

  .h2('How extends Works')
  .p(
    'Think of extends as a CSS pre-processor mixin: the final style ' +
    'is the result of merging the parent into the child, with child ' +
    'values winning on conflict.',
    { className: 'body' },
  )

  // ─── Demonstrate each style ───────────────────────────────────────
  .h2('Heading Level 2')
  .p(
    'This paragraph uses body class, which extends baseText. ' +
    'It inherits fontSize: 11 and fontFamily from baseText.',
    { className: 'body' },
  )

  .h3('Heading Level 3')
  .p(
    'h3 extends h2, which extends body, which extends baseText. ' +
    'That is a 3-level inheritance chain.',
    { className: 'body' },
  )

  .p(
    '⚠ Warning: This is a warning box that extends both baseText and callout. ' +
    'The result merges both: background from callout, base font from baseText, ' +
    'and its own color/textAlign overrides.',
    { className: 'warning' },
  )

  // ─── Table showing inheritance resolution ─────────────────────────
  .h2('Inheritance Resolution Table')
  .table<FeatureRow>({
    columns: [
      { key: 'feature', title: 'Style Class', width: '25%' },
      { key: 'inherits', title: 'Extends', width: '25%' },
      { key: 'resolvesTo', title: 'Effective Properties', width: '50%' },
    ],
    data: [
      {
        feature: 'baseText',
        inherits: '(none)',
        resolvesTo: 'fontSize:11, fontFamily:Calibri, color:#333, lineHeight:1.5',
      },
      {
        feature: 'body',
        inherits: 'baseText',
        resolvesTo: 'baseText + marginTop:4, marginBottom:4',
      },
      {
        feature: 'h2',
        inherits: 'body',
        resolvesTo: 'body + fontSize:18, bold, color:#1a1a2e',
      },
      {
        feature: 'h3',
        inherits: 'h2',
        resolvesTo: 'h2 - fontSize:14, marginTop:12, marginBottom:6',
      },
      {
        feature: 'tableHeader',
        inherits: 'cell → baseText',
        resolvesTo: 'baseText + fontSize:10, bold, bg:#1e293b, color:#fff',
      },
      {
        feature: 'warning',
        inherits: 'baseText + callout',
        resolvesTo: 'baseText + callout + color:#92400e, border',
      },
    ],
    bordered: true,
    striped: true,
    headerCellStyle: { fontWeight: 'bold', fontSize: 10 },
    cellStyle: { fontSize: 10, verticalAlign: 'middle' },
  })

  .save('style-inheritance-demo.docx')
```

## 本例展示的能力

| 功能 | 使用位置 |
|---|---|
| `extends: 'parent'`（單繼承） | `body`、`h2`、`h3`、`cell`、`tableHeader` |
| `extends: ['a', 'b']`（多繼承） | `warning` 合併 `baseText` 與 `callout` |
| 繼承鏈（3 層） | `h3` → `h2` → `body` → `baseText` |
| 選擇性覆蓋 | `h3` 保留 `h2` 的屬性，但修改 `fontSize` |
| `defineStyles()` 驗證 | 在定義時報告無效的 `extends` 引用 |

## `extends` 語法

```ts
const styles = defineStyles({
  // Single parent
  child: {
    extends: 'parent',
    fontSize: 14,     // override
    // ...inherits all other properties from 'parent'
  },

  // Multiple parents (later array entries win on conflict)
  merged: {
    extends: ['base', 'override'],
    // ...inherits from both, 'override' wins on conflicts
  },
})
```

## 繼承解析順序

```
defineStyles({
  A: { color: 'red',   fontSize: 12 },
  B: { extends: 'A',   fontSize: 14 },    // B = { color: 'red', fontSize: 14 }
  C: { extends: 'B',   fontWeight: 'bold' }, // C = A + B + { fontWeight: 'bold' }
})
```

解析採用**深度優先**遍歷，能夠**檢測循環引用**，並在文件**編譯階段**完成。
