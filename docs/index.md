---
layout: home
hero:
  name: Docx Kit
  text: CSS-like DOCX API Kit
  tagline: Type-safe, plugin-extensible Word document generation for Node.js & browser
  actions:
    - theme: brand
      text: Get Started
      link: /guide/getting-started
    - theme: alt
      text: View on GitHub
      link: https://github.com/ntnyq/docx-kit
  # image:
  #   src: ''
  #   alt: docx-kit

features:
  - icon: 🎨
    title: CSS-like Styling
    details: Use familiar CSS property names (fontSize, color, textAlign, margin, border...) to style your documents. Define reusable class-based stylesheets, theme tokens, and pre-built presets — no raw OOXML needed.

  - icon: ⛓️
    title: Fluent Builder API
    details: Chain .h1(), .p(), .table(), .image(), .pageBreak() calls to build documents naturally. TypeScript infers your stylesheet classes for autocomplete.

  - icon: 🧩
    title: 12 Built-in Plugins
    details: QR Code, ECharts, Code Block, Callout, Cover Page, Data Table, Meeting Minutes, Page Number, Property Table, Signature Block, Timeline, Watermark. Register with .use() and invoke with .plugin().

  - icon: 📦
    title: Presets & Themes
    details: Three style presets (classic gov-doc, modern business, academic thesis) and three themes (minimal, ocean, warm) with token-based color and font systems.

  - icon: 📄
    title: JSON DSL (AI-Friendly)
    details: Generate documents from plain JSON with renderDocx(). Ideal for AI/LLM-driven doc generation, API integrations, and serverless pipelines. Includes 4 AI templates.

  - icon: 🤖
    title: MCP Server & AI Templates
    details: Expose docx-kit as Model Context Protocol (MCP) tools. Comes with 4 built-in templates (report, invoice, resume, letter) and a prompt builder for LLM-powered doc generation.

  - icon: 🖥️
    title: Browser Preview
    details: Render .docx files directly in the browser with @docxkit/renderer. Supports DOM-based rendering (via docx-preview) and Microsoft Office Online iframe — perfect for live preview in web apps.

  - icon: 🌐
    title: Cross-Platform
    details: Works in Node.js (save to disk) and browsers (download as Blob). Platform-specific APIs are cleanly separated under docx-kit/node and docx-kit/browser.

  - icon: 🛡️
    title: Fully Typed & Tested
    details: Every API has precise generics with TSDoc comments. Structured DocxKitError with error codes. Plugin security hooks for sandboxing untrusted plugins.
---
