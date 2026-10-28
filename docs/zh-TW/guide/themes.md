# 主題

主題是**語義化設計令牌**，包含顏色、字型、字級與間距體系，可透過 `$category.key` 在樣式規則中引用，將_樣式定義_與_具體顏色和字型_解耦。主題字型還會提供 Word 原生預設值，讓未設置樣式的文字和外掛自動繼承文件字型。

## 什麼是主題？

`DocxTheme` 是設計令牌的集合：

```ts
interface DocxTheme {
  id?: string
  name?: string
  colors?: Record<string, string> // palette
  fonts?: Record<string, string> // font stacks
  fontSize?: Record<string, number> // scale (pt)
  spacing?: Record<string, number> // scale (pt)
}
```

**為什麼使用主題？**

- 修改一個欄位即可切換整份文件的外觀
- 讓多份文件保持配色一致
- 按語義重用 `primary`、`accent`、`success` 等令牌
- 與[預設](/zh-TW/guide/presets)配合，得到完整初始設定

## 內建主題

docx-kit 提供 **3 個內建主題**：

| 主題 | ID | 配色 | 適用場景 | 字型 |
| -------------- | --------- | ----------------------------------- | ---------------------- | ----------------- |
| `minimalTheme` | `minimal` | 簡潔灰階與藍色強調色 | 通用商務文件 | Calibri |
| `oceanTheme` | `ocean` | 海軍藍、青綠色與淺灰底色 | 報告、文章 | Cambria / Calibri |
| `warmTheme` | `warm` | 暖灰與棕色 | 信函、隨筆 | Georgia |

輔助函式：

```ts
import {
  useTheme,
  THEME_LIST,
  minimalTheme,
  oceanTheme,
  warmTheme,
} from 'docx-kit'

// All three are exported by name:
const ocean = oceanTheme
const minimal = minimalTheme
const warm = warmTheme

// Or look up by ID:
const found = useTheme('ocean') // returns DocxTheme | undefined

// Or iterate all built-ins:
for (const t of THEME_LIST) {
  console.log(t.id, t.name)
}
// minimal  ocean  warm
```

## 使用主題

將主題傳入 `createDocx()`：

```ts
import { createDocx, defineStyles, useTheme } from 'docx-kit'

const doc = createDocx({
  theme: useTheme('ocean'),
  styles: defineStyles({
    title: {
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      fontSize: '$fontSize.xl',
      marginBottom: '$spacing.lg',
    },
  }),
})

doc.h1('Annual Report 2026', { className: 'title' })
```

樣式規則中以 `$` 開頭的字串會作為令牌引用，在文件編譯階段解析。

## 令牌分類

### `$colors.*` — 調色板

每個內建主題提供 12 個語義化顏色令牌：

| 令牌 | 用途 | `minimal` | `ocean` | `warm` |
| ------------ | ---------------------- | --------- | --------- | --------- |
| `primary` | 標題、主要文字 | `#111827` | `#0f172a` | `#44403c` |
| `secondary` | 次要文字 | `#374151` | `#1e293b` | `#57534e` |
| `accent` | 強調色與連結顏色 | `#2563eb` | `#0d9488` | `#8b5e34` |
| `success` | 成功狀態 | `#2e7d32` | `#14765c` | `#15803d` |
| `warning` | 警告狀態 | `#9a6700` | `#946200` | `#a34c13` |
| `danger` | 錯誤狀態 | `#b91c1c` | `#e11d48` | `#b91c1c` |
| `info` | 資訊狀態 | `#245a96` | `#24758a` | `#6b6578` |
| `muted` | 弱化文字 | `#6b7280` | `#64748b` | `#78716c` |
| `background` | 頁面背景 | `#ffffff` | `#ffffff` | `#ffffff` |
| `surface` | 卡片或面板底色 | `#f9fafb` | `#f0f5f7` | `#f6f3ee` |
| `text` | 正文 | `#1f2937` | `#1e293b` | `#292524` |
| `border` | 邊框顏色 | `#e5e7eb` | `#cbd9df` | `#ddd6ca` |

### `$fonts.*` — 字型棧

每個內建主題提供正文、標題和程式碼字型，以及分別用於中日韓文字的 `bodyEastAsia` 和 `headingEastAsia` 字型。Word 對每種文字系統僅接受一種字型，CSS 備援字型棧會取第一個字型。

| 令牌 | `minimal` | `ocean` | `warm` |
| --------- | ---------- | ---------- | ---------- |
| `heading` | `Calibri`  | `Cambria`  | `Georgia`  |
| `body`    | `Calibri`  | `Calibri`  | `Georgia`  |
| `code`    | `Consolas` | `Consolas` | `Consolas` |

Minimal 和 Ocean 的 `bodyEastAsia`、`headingEastAsia` 使用 Microsoft YaHei。
Warm 的正文使用 SimSun，標題使用 SimHei。這些是 Word 字型名稱，
打開文件的應用會在字型缺失時選擇替代字型。

### `$fontSize.*` — 字級體系

每個內建主題提供 5 級字級，單位為 pt：

| 令牌 | `minimal` | `ocean` | `warm` |
| ------ | --------- | ------- | ------ |
| `xs`   | 8         | 8       | 8      |
| `sm`   | 9         | 10      | 10     |
| `base` | 11        | 12      | 12     |
| `lg`   | 14        | 16      | 15     |
| `xl`   | 24        | 26      | 24     |

### `$spacing.*` — 間距體系

單位為 pt，可用於 `margin*`、`padding*`、`borderWidth` 等：

| 令牌 | `minimal` | `ocean` | `warm` |
| ----- | --------- | ------- | ------ |
| `xs`  | 4         | 5       | 4      |
| `sm`  | 8         | 10      | 9      |
| `md`  | 16        | 18      | 16     |
| `lg`  | 24        | 28      | 26     |
| `xl`  | 32        | 36      | 34     |

## 自訂主題

將自訂 `DocxTheme` 物件傳入 `createDocx()`：

```ts
import { createDocx, defineStyles, type DocxTheme } from 'docx-kit'

const brand: DocxTheme = {
  id: 'brand-2026',
  name: 'Brand 2026',
  colors: {
    primary: '#1a56db',
    accent: '#f59e0b',
    danger: '#dc2626',
    success: '#16a34a',
    text: '#1f2937',
    muted: '#6b7280',
    background: '#ffffff',
    surface: '#f3f4f6',
    border: '#e5e7eb',
  },
  fonts: {
    heading: 'Poppins, Arial, sans-serif',
    body: 'Inter, Arial, sans-serif',
    code: 'JetBrains Mono, monospace',
  },
  fontSize: {
    xs: 8,
    sm: 10,
    base: 11,
    lg: 14,
    xl: 18,
  },
  spacing: {
    xs: 4,
    sm: 8,
    md: 16,
    lg: 24,
    xl: 32,
  },
}

const doc = createDocx({
  theme: brand,
  styles: defineStyles({
    title: {
      color: '$colors.primary',
      fontFamily: '$fonts.heading',
      fontSize: '$fontSize.xl',
      marginBottom: '$spacing.lg',
    },
    body: {
      color: '$colors.text',
      fontFamily: '$fonts.body',
      fontSize: '$fontSize.base',
    },
  }),
})
```

## 令牌引用語法

可以在樣式規則的任意欄位中引用令牌：

```ts
defineStyles({
  card: {
    color: '$colors.text', // → '#1f2937'
    backgroundColor: '$colors.surface',
    fontFamily: '$fonts.body',
    fontSize: '$fontSize.base',
    margin: '$spacing.md', // shorthand — sets all sides
    border: {
      color: '$colors.border',
      style: 'single',
      width: 1,
    },
  },
  calloutSuccess: {
    color: '$colors.success',
    backgroundColor: '$colors.surface',
    borderColor: '$colors.success',
  },
})
```

## 結合預設

預設定義具體樣式規則，主題定義令牌。可以將預設排版與主題配色組合使用。Modern 預設使用主色與強調色令牌；Classic 和 Academic 的標題保持黑色。預設字級、正文字型和間距優先於主題預設值。

```ts
import { createDocx, modernPreset, useTheme, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  styles: {
    ...modernPreset.config.styles,
    // Use theme tokens for an extra layer
    highlight: {
      color: '$colors.accent',
      backgroundColor: '$colors.surface',
    },
  },
})
```

## 預覽常見 Word 佈局

[主題工作室](https://docx-kit-theme-studio.ntnyq.dev/?lang=zh-TW)提供商務報告、中文公文與學術佈局，每種佈局都包含代表性內容和帶樣式表格。工作室會顯式將佈局字型、主要字級與標題間距綁定到可編輯主題令牌，匯出主題側欄包含相應設定。使用**下載 Word**儲存與預覽相同的檔案。預覽保留匯出頁邊距；已安裝字型與 Word 佈局引擎仍可能使最終顯示與瀏覽器有所不同。

## 後續閱讀

- [預設](/zh-TW/guide/presets) — 預設定樣式組合
- [樣式](/zh-TW/guide/styling) — 完整樣式系統參考
- [範例：主題系統](/zh-TW/examples/theme-system) — 完整主題報告
