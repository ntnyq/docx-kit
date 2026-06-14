import { defineConfig } from 'vitepress'
import { APP_TITLE } from './meta'
import {
  buildEnglishSidebar,
  buildSimplifiedSidebar,
  buildTraditionalSidebar,
  englishNav,
  navWithPrefix,
  simplifiedNav,
  traditionalNav,
} from './site'

export default defineConfig({
  lang: 'en-US',
  title: APP_TITLE,
  description:
    'CSS-like DOCX API Kit — Type-safe, plugin-extensible Word document generation',
  locales: {
    root: {
      label: 'English',
      lang: 'en-US',
      themeConfig: {
        nav: navWithPrefix('', englishNav),
        sidebar: buildEnglishSidebar(),
      },
    },
    'zh-CN': {
      label: '简体中文',
      lang: 'zh-CN',
      link: '/zh-CN/',
      title: APP_TITLE,
      description:
        'CSS 风格的 DOCX API 工具包，支持类型安全、插件扩展的 Word 文档生成',
      themeConfig: {
        nav: navWithPrefix('/zh-CN', simplifiedNav),
        sidebar: buildSimplifiedSidebar('/zh-CN'),
      },
    },
    'zh-TW': {
      label: '繁體中文',
      lang: 'zh-TW',
      link: '/zh-TW/',
      title: APP_TITLE,
      description:
        'CSS 風格的 DOCX API 工具包，支援型別安全、可擴充外掛的 Word 文件生成',
      themeConfig: {
        nav: navWithPrefix('/zh-TW', traditionalNav),
        sidebar: buildTraditionalSidebar('/zh-TW'),
      },
    },
  },

  markdown: {
    theme: {
      dark: 'github-dark',
      light: 'github-light',
    },
  },

  themeConfig: {
    nav: navWithPrefix('', englishNav),
    sidebar: buildEnglishSidebar(),

    // logo: '',
    footer: {
      copyright: 'Copyright © ntnyq',
      message: 'Released under the MIT License.',
    },
    search: {
      provider: 'local',
      options: {
        locales: {
          'zh-CN': {
            translations: {
              button: {
                buttonAriaLabel: '搜索文档',
                buttonText: '搜索',
              },
              modal: {
                displayDetails: '显示详情',
                noResultsText: '没有找到结果',
                resetButtonTitle: '清空搜索条件',
                footer: {
                  closeKeyAriaLabel: '关闭',
                  closeText: '关闭',
                  navigateText: '切换',
                  selectText: '选择',
                },
              },
            },
          },
          'zh-TW': {
            translations: {
              button: {
                buttonAriaLabel: '搜尋文件',
                buttonText: '搜尋',
              },
              modal: {
                displayDetails: '顯示詳細資訊',
                noResultsText: '找不到結果',
                resetButtonTitle: '清除搜尋條件',
                footer: {
                  closeKeyAriaLabel: '關閉',
                  closeText: '關閉',
                  navigateText: '切換',
                  selectText: '選擇',
                },
              },
            },
          },
        },
      },
    },

    socialLinks: [
      { icon: 'github', link: 'https://github.com/ntnyq/docx-kit' },
    ],
  },

  vite: {
    optimizeDeps: {
      include: ['monaco-editor'],
    },
    worker: {
      format: 'es',
    },
  },
})
