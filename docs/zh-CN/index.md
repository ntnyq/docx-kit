---
layout: home
hero:
  name: Docx Kit
  text: CSS 风格 DOCX API 工具包
  tagline: 面向 Node.js 与浏览器的类型安全、可插件扩展的 Word 文档生成方案
  image:
    src: /brand/logo.svg
    alt: Docx Kit
  actions:
    - theme: brand
      text: 快速开始
      link: /zh-CN/guide/getting-started
    - theme: alt
      text: 查看包总览
      link: /zh-CN/ecosystem/packages
    - theme: alt
      text: Playground
      link: https://docx-kit-playground.ntnyq.dev/?lang=zh-CN
    - theme: alt
      text: Theme Studio
      link: https://docx-kit-theme-studio.ntnyq.dev/?lang=zh-CN
    - theme: alt
      text: 查看 GitHub
      link: https://github.com/ntnyq/docx-kit

features:
  - icon: 🎨
    title: CSS 风格样式
    details: "使用熟悉的 `fontSize`、`color`、`margin`、`border` 和可复用类名样式，组织文档布局与文字排版。"

  - icon: ⛓️
    title: 链式 Builder API
    details: "链式调用 `.h1()`、`.p()`、`.table()`、`.image()`、`.section()` 与 `.plugin()`，以完整类型推断逐步构建文档。"

  - icon: 🧩
    title: 22 个内置插件
    details: "内置插件涵盖徽章、条形码、提示块、代码块、封面、发票、二维码、时间线、目录等常见内容。"

  - icon: 📦
    title: Monorepo 包生态
    details: "工作区包含运行时、渲染器、共享类型、插件工具、预设、主题、AI 模板与 MCP 集成，并提供包级文档。"

  - icon: 📄
    title: JSON DSL
    details: "使用 `renderDocx()` 和适合 schema 描述的文档节点，构建 API 驱动或 AI 生成的文档流程。"

  - icon: 🤖
    title: AI 模板与 MCP
    details: "通过 `@docxkit/ai` 生成报告、发票、简历和信函，或通过 `@docxkit/mcp` 将完整系统集成到 Agent 运行时。"

  - icon: 🖥️
    title: 浏览器预览
    details: "`@docxkit/renderer` 支持 DOM 和 Office Online 渲染策略，可直接在网页应用中预览生成的 `.docx` 文件。"

  - icon: 🌐
    title: 多语言文档
    details: "文档通过 VitePress locales 路由提供英文、简体中文与繁体中文版本。"

  - icon: 🛡️
    title: 类型安全、充分测试、易于扩展
    details: "公共 API 提供强类型支持，各包具有测试覆盖，插件通过显式注册与包边界保持独立。"
---
