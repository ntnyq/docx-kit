# Code Block

输出等宽字体代码块，可选行号与语法高亮。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `code` | `string` | 源代码文本 |
| `language` | `string` | 语法高亮语言，可选 |
| `showLineNumbers` | `boolean` | 是否显示行号 |

## 用法

```ts
createDocx().use(codeBlockPlugin()).plugin('codeBlock', {
  code: 'const x = 1',
  language: 'ts',
  showLineNumbers: true,
})
```

`highlight.js` 为可选依赖，未安装时会回退为纯文本代码块。
