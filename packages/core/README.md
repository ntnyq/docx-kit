# @docxkit/core

Core engine for docx-kit — a CSS-like, type-safe, plugin-extensible DOCX generation library.

## Features

- Fluent builder API (`DocxBuilder`)
- JSON schema API (`renderDocx`)
- CSS-like style DSL with cascade resolution
- Plugin system with registration and rendering
- TypeScript-first with full type inference

## Installation

```bash
npm install @docxkit/core
```

## Usage

```ts
import { createDocx } from '@docxkit/core'

const docx = createDocx()
  .paragraph('Hello, world!')
  .heading('Introduction', { level: 1 })
  .build()
```
