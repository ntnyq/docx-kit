# 分隔线（Divider）

在章节之间插入横线，或从视觉上分隔内容块。

## 导入

```ts
import { dividerPlugin, type DividerOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | `'solid'` | 分隔线边框样式 |
| `color` | `string` | `'D9D9D9'` | 十六进制分隔线颜色 |
| `spacingBefore` | `number` | `200` | 分隔线段前间距 |
| `spacingAfter` | `number` | `200` | 分隔线段后间距 |

## 示例

```ts
import { createDocx, dividerPlugin } from 'docx-kit'

const doc = createDocx()
  .use(dividerPlugin())
  .h1('Quarterly Review')
  .plugin('divider', { style: 'double', color: '4472C4' })
  .p('Executive summary starts here.')
```

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `DividerConfig` 和 `DividerStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `paragraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

```ts
dividerPlugin({
  styles: { paragraph: { borderBottom: { color: '#2563EB', width: 2 }, marginTop: 0 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
