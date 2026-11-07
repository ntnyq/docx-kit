# 代码块（Code Block）

使用等宽字体渲染源代码，支持可选行号与 `highlight.js` 语法高亮。

## 导入

```ts
import { codeBlockPlugin, type CodeBlockOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `code` | `string` | _必填_ | 源代码字符串 |
| `language` | `string` | — | 语法高亮语言标识符，需要 `highlight.js` |
| `showLineNumbers` | `boolean` | `false` | 在每行前显示行号 |

## 示例

### 基础代码块

```ts
import { createDocx, codeBlockPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(codeBlockPlugin())
  .h1('Code Example')
  .plugin('codeBlock', {
    code: `function hello(name: string): string {
  return \`Hello, \${name}!\`
}

console.log(hello('World'))`,
    language: 'typescript',
  })
  .save('code.docx')
```

### 显示行号

```ts
const doc = createDocx()
  .use(codeBlockPlugin())
  .h1('Algorithm')
  .plugin('codeBlock', {
    code: `def fibonacci(n):
    if n <= 1:
        return n
    return fibonacci(n - 1) + fibonacci(n - 2)

for i in range(10):
    print(fibonacci(i))`,
    language: 'python',
    showLineNumbers: true,
  })
  .save('code-lines.docx')
```

### 纯文本（无高亮）

```ts
const doc = createDocx()
  .use(codeBlockPlugin())
  .h1('Configuration')
  .plugin('codeBlock', {
    code: `# Server Configuration
host = 0.0.0.0
port = 8080
max_connections = 1000
timeout = 30s`,
  })
  .save('plain-code.docx')
```

### 多个代码块

```ts
const doc = createDocx()
  .use(codeBlockPlugin())
  .h1('API Examples')

  .h2('JavaScript')
  .plugin('codeBlock', {
    code: `const response = await fetch('/api/users')
const users = await response.json()
console.log(users.length)`,
    language: 'javascript',
    showLineNumbers: true,
  })

  .h2('TypeScript')
  .plugin('codeBlock', {
    code: `interface User {
  id: number
  name: string
  email: string
}

async function getUsers(): Promise<User[]> {
  const res = await fetch('/api/users')
  return res.json()
}`,
    language: 'typescript',
    showLineNumbers: true,
  })

  .save('api-examples.docx')
```

### 语法高亮依赖

安装 `highlight.js` 启用语法高亮：

```bash
pnpm add highlight.js
```

未安装时，代码仍使用等宽字体，但不显示语法颜色。

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `CodeBlockConfig` 和 `CodeBlockStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `code`, `lineNumber` | `TextStyleRule` | 文字排版 |
| `paragraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |

`code` 同时作用于普通文本和高亮 token；显式颜色会覆盖语法配色。整行背景使用 `paragraph.backgroundColor`。

```ts
codeBlockPlugin({
  styles: { code: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
