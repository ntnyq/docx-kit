# Code Block

Renders source code with monospaced font, optional line numbers, and syntax highlighting via `highlight.js`.

## Import

```ts
import { codeBlockPlugin, type CodeBlockOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `code` | `string` | _(required)_ | Source code string |
| `language` | `string` | — | Language identifier for syntax highlighting (requires `highlight.js`) |
| `showLineNumbers` | `boolean` | `false` | Prepend line numbers |

## Examples

### Basic Code Block

```ts
import { createDocx, codeBlockPlugin } from 'docx-kit'

const doc = createDocx()
  .use(codeBlockPlugin)
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

### With Line Numbers

```ts
const doc = createDocx()
  .use(codeBlockPlugin)
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

### Plain Text (No Highlighting)

```ts
const doc = createDocx()
  .use(codeBlockPlugin)
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

### Multiple Code Blocks

```ts
const doc = createDocx()
  .use(codeBlockPlugin)
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

### Syntax Highlighting Dependency

Install `highlight.js` for syntax highlighting:

```bash
pnpm add highlight.js
```

Without it, code renders with monospaced font but no colors.
