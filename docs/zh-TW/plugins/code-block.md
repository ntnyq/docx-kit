# 程式碼區塊（Code Block）

使用等寬字型呈現原始碼，支援可選行號與 `highlight.js` 語法高亮。

## 匯入

```ts
import { codeBlockPlugin, type CodeBlockOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `code` | `string` | _必填_ | 原始碼字串 |
| `language` | `string` | — | 語法高亮語言標識符，需要 `highlight.js` |
| `showLineNumbers` | `boolean` | `false` | 在每行前顯示行號 |

## 範例

### 基本程式碼區塊

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

### 顯示行號

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

### 純文字（無高亮）

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

### 多個程式碼區塊

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

### 語法高亮相依

安裝 `highlight.js` 啟用語法高亮：

```bash
pnpm add highlight.js
```

未安裝時，程式碼仍使用等寬字型，但不顯示語法顏色。
