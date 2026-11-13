# @docxkit/ai

AI template system for docx-kit. Provides built-in document templates (report, invoice, resume, letter) with schema-based generation and prompt building utilities.

## Installation

```bash
npm install @docxkit/ai
```

## Usage

```ts
import { reportTemplate, buildPrompt } from '@docxkit/ai'

const prompt = buildPrompt(reportTemplate, { title: 'Q4 Report' })
```

Generated tool schemas reuse draft-07 `definitions` and `$ref` for common styles
instead of expanding them for every plugin part. Keep the complete schema,
including its definitions, when passing it to a validator. The exported inline
style and plugin-option metadata remains available for help and inspection.

For custom document schemas, `referenceStyleSchemas(schema)` replaces canonical
style fragments with references; include `STYLE_SCHEMA_DEFINITIONS` in the root
`definitions` object to resolve them.
