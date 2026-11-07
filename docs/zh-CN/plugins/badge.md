# 徽章（Badge）

渲染 `DRAFT`、`APPROVED`、`INTERNAL` 等简短状态标签。

## 导入

```ts
import { badgePlugin, type BadgeOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `string` | 必填 | 徽章文字 |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | `'neutral'` | 预设色调名称或自定义文字颜色 |
| `backgroundColor` | `string` | 由预设决定 | 自定义徽章背景色 |

## 示例

```ts
import { badgePlugin, createDocx } from 'docx-kit'

const doc = createDocx()
  .use(badgePlugin())
  .plugin('badge', {
    text: 'DRAFT',
    color: 'warning',
  })
```

## 说明

- 预设提供配套的前景色与背景色。
- 自定义 `color` 若不属于预设名称，会按字面文字颜色处理。

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `BadgeConfig` 和 `BadgeStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `text` | `TextStyleRule` | 文字排版 |
| `paragraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

```ts
badgePlugin({
  styles: { text: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
