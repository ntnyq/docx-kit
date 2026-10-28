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
