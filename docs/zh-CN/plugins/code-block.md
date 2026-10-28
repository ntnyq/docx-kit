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
