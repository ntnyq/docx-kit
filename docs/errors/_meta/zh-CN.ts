import type { ErrorNotes } from './types'

export const notes = {
  EXPORT_FAILED: {
    causes: ['文档资源无效，或底层打包过程失败。'],
    summary: '编译后的文档打包失败，调用方仍可读取原始 cause。',
    title: '文档导出失败',
    steps: [
      '在本地检查 error.cause，定位失败资源。',
      '使用当前平台支持的导出格式，并确认引用资源有效。',
    ],
  },
  HEADER_FOOTER_INVALID_CONTENT: {
    causes: ['插件直接返回 TextRun 或其他不支持的子元素。'],
    summary: '页眉或页脚插件返回了 Paragraph、Table 以外的元素。',
    title: '页眉页脚内容无效',
    steps: [
      '页眉页脚内容返回 Paragraph 或 Table 实例。',
      '将行内元素包装在 Paragraph 内。',
    ],
  },
  IMAGE_INVALID_DATA: {
    causes: ['图片数据为空、格式不受支持，或路径未解析。'],
    summary: '图片内容为空、格式无法识别，或缺少路径解析器。',
    title: '图片数据无效',
    steps: [
      '提供非空且受支持的图片字节或 data URL。',
      '文件路径使用 docx-kit/node，浏览器配置 resolveImage。',
      '无法自动检测格式时检查 imageType。',
    ],
  },
  MANIFEST_INVALID: {
    causes: ['必填字段缺失、版本范围格式不正确，或路径越过包边界。'],
    summary: '清单字段、版本范围或入口路径无效。',
    title: '插件清单无效',
    steps: [
      '修正消息指出的字段。',
      '使用非空字符串、有效 semver，以及以 ./ 开头的安全相对路径。',
      '确保入口位于插件包内。',
    ],
  },
  MANIFEST_MISSING: {
    causes: ['发布包遗漏清单，或请求的路径返回错误。'],
    summary: '外部加载时未能找到所需的 docx-kit.plugin.json。',
    title: '插件清单缺失',
    steps: [
      '在包根目录包含清单文件。',
      '导入可执行代码前，确认清单路径或 URL 可访问。',
    ],
  },
  MICROSOFT_URL_REQUIRED: {
    causes: ['Microsoft 预览收到本地字节、Blob URL，或非 HTTP(S) URL。'],
    summary: 'Microsoft Office Online 预览需要公开可访问的 HTTP(S) DOCX URL。',
    title: 'Microsoft 预览需要 URL',
    steps: [
      '将文档放在公开可访问的 HTTP(S) URL。',
      '本地 Blob 或私有文件使用本地 docx-preview 渲染器。',
    ],
  },
  PLUGIN_DEPENDENCY_MISSING: {
    causes: ['集成显式报告依赖缺失或版本不兼容。'],
    summary: '这是公开的插件依赖错误码，当前内置加载器不会自动抛出此码。',
    title: '插件依赖缺失',
    steps: [
      '自定义集成报告此码时，检查该集成的依赖校验。',
      '安装插件要求的依赖及兼容版本。',
    ],
  },
  PLUGIN_LOAD_FAILED: {
    causes: ['来源类型不受支持、策略拒绝、导入错误，或插件导出无效。'],
    summary: '平台环境、安全策略、模块导入、模块导出或插件工厂导致加载失败。',
    title: '插件加载失败',
    steps: [
      '环境问题使用对应的浏览器或 Node.js 加载器。',
      '导入失败时检查入口、依赖和网络。',
      '导出或工厂失败时返回有效插件实例，并检查原始 cause。',
      '安全策略拒绝时核对来源和权限，继续阻止不可信代码。',
    ],
  },
  PLUGIN_NOT_REGISTERED: {
    causes: ['遗漏 .use()、未调用工厂，或运行时名称不一致。'],
    title: '插件未注册',
    steps: [
      '导出前调用插件工厂，将返回的实例传给 .use()。',
      '节点名必须匹配 plugin.name，第三方工厂名称不一定遵循内置命名方式。',
      'JSON DSL 同样需要注册插件。',
    ],
    summary:
      'plugin 节点引用了未注册的运行时名称，PDK 断言也会用此码报告无效插件实例。',
  },
  PLUGIN_RENDER_FAILED: {
    causes: ['渲染器异常、选项无效，或 PDK 输出类型不符。'],
    summary: '插件 render 方法抛错，或 PDK 渲染断言拒绝其输出。',
    title: '插件渲染失败',
    steps: [
      '在本地检查原始 cause，并核对插件选项。',
      'render 返回受支持的 docx 元素。',
      'PDK 断言失败时核对预期元素类型和数量。',
    ],
  },
  PLUGIN_VERSION_MISMATCH: {
    causes: ['发布版本组合不兼容，或运行版本、版本范围无效。'],
    summary: '清单中的 docxKit 范围不接受当前库版本。',
    title: '插件版本不兼容',
    steps: [
      '对比消息中的版本范围和已安装版本。',
      '选择兼容的插件与 docx-kit 版本组合。',
    ],
  },
  PREVIEW_FETCH_FAILED: {
    causes: ['HTTP 响应失败、网络错误，或跨域请求被阻止。'],
    summary: '未能成功从 URL 获取 DOCX 文件。',
    title: '预览文件获取失败',
    steps: [
      '检查 HTTP 响应、网络、身份验证与 CORS 设置。',
      '无法远程获取时传入本地 DOCX 字节。',
    ],
  },
  PREVIEW_INPUT_INVALID: {
    causes: ['传入 null，或不支持的对象、字符串类型。'],
    summary: '预览输入缺失，或类型不受支持。',
    title: '预览输入无效',
    steps: [
      '传入 DOCX Blob、ArrayBuffer、Uint8Array 或 URL。',
      '生成完成后再开始预览。',
    ],
  },
  PREVIEW_RENDER_FAILED: {
    causes: ['归档损坏、内容不受支持，或渲染器异常。'],
    summary: '本地 docx-preview 渲染器未能处理此文档。',
    title: '预览渲染失败',
    steps: [
      '确认输入是有效的 DOCX ZIP 归档。',
      '检查渲染器原始 cause，并尝试用 Word 打开同一文件。',
    ],
  },
  STYLE_CIRCULAR_EXTENDS: {
    causes: ['样式继承自身，或多个样式类互相继承形成闭环。'],
    summary: 'extends 链引用了正在解析的样式类。',
    title: '样式继承形成循环',
    steps: [
      '沿错误消息中的继承链检查引用。',
      '删除闭合循环的引用，共用父样式仍然可用。',
    ],
  },
  STYLE_UNKNOWN_CLASS: {
    causes: ['名称拼写错误、父样式缺失，或使用了错误的样式表。'],
    summary: 'className 或 extends 引用的名称与样式表键不匹配。',
    title: '样式类不存在',
    steps: [
      '核对引用名称和样式表键，注意拼写。',
      '调用 defineStyles 前补齐父样式定义。',
    ],
  },
  TABLE_INVALID_COLUMNS: {
    causes: ['列定义为空，或单元格跨度超出网格、互相重叠。'],
    summary: '表格没有列，或单元格跨度无效、重叠、超出边界。',
    title: '表格网格无效',
    steps: [
      '至少定义一列。',
      '将 rowSpan 和 colSpan 限制在网格范围内，移除重叠的单元格位置。',
    ],
  },
  UNKNOWN_NODE_TYPE: {
    causes: ['节点类型拼写错误、行内内容使用了块节点，或分节节点位置不正确。'],
    summary: '当前上下文没有支持此节点的编译器。',
    title: '节点类型无法编译',
    steps: [
      '使用文档列出的 DSL 类型。',
      '自定义内容使用已注册的 plugin 节点。',
      'sectionBreak 放在文档层级，不要嵌套在子内容中。',
    ],
  },
} satisfies ErrorNotes
