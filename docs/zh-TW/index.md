---
layout: home
hero:
  name: Docx Kit
  text: CSS 風格 DOCX API 工具套件
  tagline: 面向 Node.js 與瀏覽器的型別安全、可外掛擴充的 Word 文件產生方案
  image:
    src: /brand/logo.svg
    alt: Docx Kit
  actions:
    - theme: brand
      text: 快速開始
      link: /zh-TW/guide/getting-started
    - theme: alt
      text: 查看套件總覽
      link: /zh-TW/ecosystem/packages
    - theme: alt
      text: 查看 GitHub
      link: https://github.com/ntnyq/docx-kit

features:
  - icon: 🎨
    title: CSS 風格樣式
    details: "使用熟悉的 `fontSize`、`color`、`margin`、`border` 和可重用類名樣式，組織文件佈局與文字排版。"

  - icon: ⛓️
    title: 鏈式 Builder API
    details: "鏈式呼叫 `.h1()`、`.p()`、`.table()`、`.image()`、`.section()` 與 `.plugin()`，以完整型別推斷逐步建置文件。"

  - icon: 🧩
    title: 22 個內建外掛
    details: "內建外掛涵蓋徽章、條碼、提示區塊、程式碼區塊、封面、發票、二維碼、時間軸、目錄等常見內容。"

  - icon: 📦
    title: Monorepo 套件生態
    details: "工作區包含執行階段、呈現器、共享型別、外掛工具、預設、主題、AI 範本與 MCP 整合，並提供套件層級文件。"

  - icon: 📄
    title: JSON DSL
    details: "使用 `renderDocx()` 和適合 schema 描述的文件節點，建置 API 驅動或 AI 產生的文件流程。"

  - icon: 🤖
    title: AI 範本與 MCP
    details: "透過 `@docxkit/ai` 產生報告、發票、履歷和信函，或透過 `@docxkit/mcp` 將完整系統整合到 Agent 執行階段。"

  - icon: 🖥️
    title: 瀏覽器預覽
    details: "`@docxkit/renderer` 支援 DOM 和 Office Online 呈現策略，可直接在網頁應用中預覽產生的 `.docx` 檔案。"

  - icon: 🌐
    title: 多語言文件
    details: "文件透過 VitePress locales 路由提供英文、簡體中文與繁體中文版本。"

  - icon: 🛡️
    title: 型別安全、充分測試、易於擴充
    details: "公共 API 提供強型別支援，各套件具有測試覆蓋，外掛透過顯式註冊與套件邊界保持獨立。"
---
