# docx-kit 项目长期记忆

## 项目定位

基于 `dolanmiu/docx` 封装的 CSS-like DOCX 生成 Kit，提供：
- CSS 风格样式配置（`defineStyles`）
- 流式 Builder DSL（`createDocx().h1().p().table().save()`）
- JSON Schema 驱动（`renderDocx(schema)`）
- 类型安全插件系统（`definePlugin` + `.use()`）

## 目录结构（单包，非 monorepo）

```
src/
  types/         # utility, style, document, plugin
  dsl/           # nodes.ts（所有 DSL 节点类型）
  style/         # normalizeStyle.ts
  compiler/      # units, compileStyle, compileNode, compileDocument
  builder/       # DocxBuilder, createDocx
  renderer/      # pack.ts
  plugins/       # qrcode/, echarts/（内置插件，peer deps 可选）
  errors.ts      # DocxKitError + ERROR_CODES
  index.ts       # 统一公共导出
```

## 关键约定

- 构建工具：`tsdown`（`pnpm build`）
- TypeScript `erasableSyntaxOnly: true`：禁止 constructor 参数 `private readonly`，需拆开写
- `echarts` / `qrcode` 为 `peerDependencies`（optional），devDependencies 里保留用于开发
- 插件通过动态 `import()` 引入 peer dep，避免无 peer dep 时崩溃
- 样式合并优先级：base → className(s) → inline style
- 单位默认：fontSize 裸数字 = pt；spacing/margin 裸数字 = pt；image 裸数字 = px

## 设计文档参考

`~/Desktop/docx-api-kit-design.md`（完整 API、类型、编译器、插件设计）
