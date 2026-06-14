# Letterhead

生成正式函件的公司信头。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `companyName` | `string` | 公司名称 |
| `tagline` | `string` | 标语 |
| `phone` | `string` | 电话 |
| `email` | `string` | 邮箱 |
| `website` | `string` | 网站 |
| `address` | `string` | 地址 |

## 用法

```ts
createDocx().use(letterheadPlugin()).plugin('letterhead', {
  companyName: 'Acme Corp',
  email: 'info@acme.com',
  website: 'acme.com',
})
```
