# 信头（Letterhead）

为正式信函渲染居中的公司页头。

## 导入

```ts
import { letterheadPlugin, type LetterheadOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `companyName` | `string` | 必填 | 主要品牌或公司名称 |
| `tagline` | `string` | — | 名称下方的可选副标题 |
| `phone` | `string` | — | 联系信息行中的电话 |
| `email` | `string` | — | 联系信息行中的邮箱 |
| `website` | `string` | — | 联系信息行中的网站 |
| `address` | `string` | — | 横线上方的最后一行地址 |

## 示例

```ts
import { createDocx, letterheadPlugin } from 'docx-kit'

const doc = createDocx()
  .use(letterheadPlugin())
  .plugin('letterhead', {
    companyName: 'Acme Corp',
    tagline: 'Designing reliable document systems',
    phone: '+1 555 0100',
    email: 'info@acme.com',
    website: 'acme.com',
    address: '1 Market Street, San Francisco, CA',
  })
```

## 说明

- 插件以蓝色下边框结束，正文可紧接其下开始。
- 联系信息自动合并为一行居中文字。

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `LetterheadConfig` 和 `LetterheadStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `address`, `companyName`, `contact`, `tagline` | `TextStyleRule` | 文字排版 |
| `addressParagraph`, `companyParagraph`, `contactParagraph`, `ruleParagraph`, `taglineParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

```ts
letterheadPlugin({
  styles: { companyName: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
