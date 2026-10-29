import type { ErrorNotes } from './types'

export const notes = {
  EXPORT_FAILED: {
    causes: ['文件資源無效，或底層封裝過程失敗。'],
    summary: '編譯後的文件封裝失敗，呼叫端仍可讀取原始 cause。',
    title: '文件匯出失敗',
    steps: [
      '在本機檢查 error.cause，定位失敗資源。',
      '使用目前平台支援的匯出格式，並確認引用資源有效。',
    ],
  },
  HEADER_FOOTER_INVALID_CONTENT: {
    causes: ['外掛直接回傳 TextRun 或其他不支援的子元素。'],
    summary: '頁首或頁尾外掛回傳了 Paragraph、Table 以外的元素。',
    title: '頁首頁尾內容無效',
    steps: [
      '頁首頁尾內容回傳 Paragraph 或 Table 實例。',
      '將行內元素包裝在 Paragraph 內。',
    ],
  },
  IMAGE_INVALID_DATA: {
    causes: ['圖片資料為空、格式不受支援，或路徑未解析。'],
    summary: '圖片內容為空、格式無法辨識，或缺少路徑解析器。',
    title: '圖片資料無效',
    steps: [
      '提供非空且受支援的圖片位元組或 data URL。',
      '檔案路徑使用 docx-kit/node，瀏覽器設定 resolveImage。',
      '無法自動偵測格式時檢查 imageType。',
    ],
  },
  MANIFEST_INVALID: {
    causes: ['必填欄位缺失、版本範圍格式不正確，或路徑越過套件邊界。'],
    summary: '資訊清單欄位、版本範圍或進入點路徑無效。',
    title: '外掛資訊清單無效',
    steps: [
      '修正訊息指出的欄位。',
      '使用非空字串、有效 semver，以及以 ./ 開頭的安全相對路徑。',
      '確保進入點位於外掛套件內。',
    ],
  },
  MANIFEST_MISSING: {
    causes: ['發佈套件遺漏資訊清單，或請求的路徑回傳錯誤。'],
    summary: '外部載入時未能找到所需的 docx-kit.plugin.json。',
    title: '外掛資訊清單缺失',
    steps: [
      '在套件根目錄包含資訊清單檔案。',
      '匯入可執行程式碼前，確認資訊清單路徑或 URL 可存取。',
    ],
  },
  MICROSOFT_URL_REQUIRED: {
    causes: ['Microsoft 預覽收到本機位元組、Blob URL，或非 HTTP(S) URL。'],
    summary: 'Microsoft Office Online 預覽需要公開可存取的 HTTP(S) DOCX URL。',
    title: 'Microsoft 預覽需要 URL',
    steps: [
      '將文件放在公開可存取的 HTTP(S) URL。',
      '本機 Blob 或私有檔案使用本機 docx-preview 渲染器。',
    ],
  },
  PLUGIN_DEPENDENCY_MISSING: {
    causes: ['整合明確回報相依套件缺失或版本不相容。'],
    summary: '這是公開的外掛相依錯誤碼，目前內建載入器不會自動擲出此碼。',
    title: '外掛相依套件缺失',
    steps: [
      '自訂整合回報此碼時，檢查該整合的相依套件驗證。',
      '安裝外掛要求的相依套件及相容版本。',
    ],
  },
  PLUGIN_LOAD_FAILED: {
    causes: ['來源類型不受支援、策略拒絕、匯入錯誤，或外掛匯出無效。'],
    summary: '平台環境、安全策略、模組匯入、模組匯出或外掛工廠導致載入失敗。',
    title: '外掛載入失敗',
    steps: [
      '環境問題使用對應的瀏覽器或 Node.js 載入器。',
      '匯入失敗時檢查進入點、相依套件和網路。',
      '匯出或工廠失敗時回傳有效外掛實例，並檢查原始 cause。',
      '安全策略拒絕時核對來源和權限，繼續阻止不可信程式碼。',
    ],
  },
  PLUGIN_NOT_REGISTERED: {
    causes: ['遺漏 .use()、未呼叫工廠，或執行階段名稱不一致。'],
    title: '外掛未註冊',
    steps: [
      '匯出前呼叫外掛工廠，將回傳的實例傳給 .use()。',
      '節點名稱必須符合 plugin.name，第三方工廠名稱不一定遵循內建命名方式。',
      'JSON DSL 同樣需要註冊外掛。',
    ],
    summary:
      'plugin 節點引用了未註冊的執行階段名稱，PDK 斷言也會用此碼回報無效外掛實例。',
  },
  PLUGIN_RENDER_FAILED: {
    causes: ['渲染器例外、選項無效，或 PDK 輸出類型不符。'],
    summary: '外掛 render 方法擲出錯誤，或 PDK 渲染斷言拒絕其輸出。',
    title: '外掛渲染失敗',
    steps: [
      '在本機檢查原始 cause，並核對外掛選項。',
      'render 回傳受支援的 docx 元素。',
      'PDK 斷言失敗時核對預期元素類型和數量。',
    ],
  },
  PLUGIN_VERSION_MISMATCH: {
    causes: ['發行版本組合不相容，或執行版本、版本範圍無效。'],
    summary: '資訊清單中的 docxKit 範圍不接受目前程式庫版本。',
    title: '外掛版本不相容',
    steps: [
      '比較訊息中的版本範圍和已安裝版本。',
      '選擇相容的外掛與 docx-kit 版本組合。',
    ],
  },
  PREVIEW_FETCH_FAILED: {
    causes: ['HTTP 回應失敗、網路錯誤，或跨來源請求遭到阻擋。'],
    summary: '未能成功從 URL 擷取 DOCX 檔案。',
    title: '預覽檔案擷取失敗',
    steps: [
      '檢查 HTTP 回應、網路、身分驗證與 CORS 設定。',
      '無法遠端擷取時傳入本機 DOCX 位元組。',
    ],
  },
  PREVIEW_INPUT_INVALID: {
    causes: ['傳入 null，或不支援的物件、字串類型。'],
    summary: '預覽輸入缺失，或類型不受支援。',
    title: '預覽輸入無效',
    steps: [
      '傳入 DOCX Blob、ArrayBuffer、Uint8Array 或 URL。',
      '生成完成後再開始預覽。',
    ],
  },
  PREVIEW_RENDER_FAILED: {
    causes: ['封存檔損壞、內容不受支援，或渲染器例外。'],
    summary: '本機 docx-preview 渲染器未能處理此文件。',
    title: '預覽渲染失敗',
    steps: [
      '確認輸入是有效的 DOCX ZIP 封存檔。',
      '檢查渲染器原始 cause，並嘗試用 Word 開啟同一檔案。',
    ],
  },
  STYLE_CIRCULAR_EXTENDS: {
    causes: ['樣式繼承自身，或多個樣式類別互相繼承形成閉環。'],
    summary: 'extends 鏈引用了正在解析的樣式類別。',
    title: '樣式繼承形成循環',
    steps: [
      '沿錯誤訊息中的繼承鏈檢查引用。',
      '刪除閉合循環的引用，共用父樣式仍可使用。',
    ],
  },
  STYLE_UNKNOWN_CLASS: {
    causes: ['名稱拼寫錯誤、父樣式缺失，或使用了錯誤的樣式表。'],
    summary: 'className 或 extends 引用的名稱與樣式表鍵不符。',
    title: '樣式類別不存在',
    steps: [
      '核對引用名稱和樣式表鍵，注意拼寫。',
      '呼叫 defineStyles 前補齊父樣式定義。',
    ],
  },
  TABLE_INVALID_COLUMNS: {
    causes: ['欄定義為空，或儲存格跨度超出網格、互相重疊。'],
    summary: '表格沒有欄，或儲存格跨度無效、重疊、超出邊界。',
    title: '表格網格無效',
    steps: [
      '至少定義一欄。',
      '將 rowSpan 和 colSpan 限制在網格範圍內，移除重疊的儲存格位置。',
    ],
  },
  UNKNOWN_NODE_TYPE: {
    summary: '目前上下文沒有支援此節點的編譯器。',
    title: '節點類型無法編譯',
    causes: [
      '節點類型拼寫錯誤、行內內容使用了區塊節點，或分節節點位置不正確。',
    ],
    steps: [
      '使用文件列出的 DSL 類型。',
      '自訂內容使用已註冊的 plugin 節點。',
      'sectionBreak 放在文件層級，不要巢狀放在子內容中。',
    ],
  },
} satisfies ErrorNotes
