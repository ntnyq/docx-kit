# docx-kit

CSS-like DOCX API Kit — a type-safe, plugin-extensible DOCX generation library built on [dolanmiu/docx](https://github.com/dolanmiu/docx).

This is the umbrella package that re-exports all core APIs, built-in plugins, presets, and themes.

## Installation

```bash
npm install docx-kit
```

## Usage

```ts
import { createDocx } from 'docx-kit'

const docx = createDocx()
  .paragraph('Hello, world!')
  .build()

// Use built-in plugins
import { calloutPlugin, codeBlockPlugin } from 'docx-kit'
```
