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
