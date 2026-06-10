---
layout: home
hero:
  name: docx-kit
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
    details: Use familiar CSS property names (fontSize, color, textAlign, margin, border...) to style your documents. Define reusable class-based stylesheets — no raw OOXML needed.

  - icon: ⛓️
    title: Fluent Builder API
    details: Chain .h1(), .p(), .table(), .image(), .pageBreak() calls to build documents naturally. TypeScript infers your stylesheet classes for autocomplete.

  - icon: 🧩
    title: Plugin Extensible
    details: Built-in QRCode & ECharts plugins. Define custom plugins with definePlugin() — render anything from badges to signatures to charts.

  - icon: 📄
    title: JSON DSL (AI-Friendly)
    details: Generate documents from plain JSON with renderDocx(). Ideal for AI/LLM-driven doc generation, API integrations, or serverless pipelines.

  - icon: 🌐
    title: Cross-Platform
    details: Works in Node.js (save to disk) and browsers (download as Blob). Platform-specific APIs are cleanly separated under docx-kit/node and docx-kit/browser.

  - icon: 🛡️
    title: Fully Typed
    details: Every API has precise generics with TSDoc comments. Structured DocxKitError with error codes for reliable error handling.
---
