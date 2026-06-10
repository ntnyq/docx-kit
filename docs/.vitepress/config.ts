import { defineConfig } from 'vitepress'

export default defineConfig({
  lang: 'en-US',
  title: 'docx-kit',
  description:
    'CSS-like DOCX API Kit — Type-safe, plugin-extensible Word document generation',

  markdown: {
    theme: {
      dark: 'github-dark',
      light: 'github-light',
    },
  },

  themeConfig: {
    logo: '',
    footer: {
      copyright: 'Copyright © ntnyq',
      message: 'Released under the MIT License.',
    },

    nav: [
      { link: '/guide/getting-started', text: 'Guide' },
      { link: '/examples/basic-report', text: 'Examples' },
      { link: '/api/types', text: 'API' },
    ],

    search: {
      provider: 'local',
    },

    sidebar: {
      '/api/': [
        {
          items: [{ link: '/api/types', text: 'Types' }],
          text: 'API Reference',
        },
      ],
      '/examples/': [
        {
          text: 'Examples',
          items: [
            { link: '/examples/basic-report', text: 'Basic Report' },
            { link: '/examples/invoice', text: 'Invoice' },
            { link: '/examples/chart-report', text: 'Chart Report' },
          ],
        },
      ],
      '/guide/': [
        {
          text: 'Guide',
          items: [
            { link: '/guide/getting-started', text: 'Getting Started' },
            { link: '/guide/builder-api', text: 'Builder API' },
            { link: '/guide/styling', text: 'CSS-like Styling' },
            { link: '/guide/tables', text: 'Tables' },
            { link: '/guide/images', text: 'Images' },
            { link: '/guide/plugins', text: 'Plugins' },
            { link: '/guide/json-dsl', text: 'JSON DSL (renderDocx)' },
            { link: '/guide/platforms', text: 'Node.js & Browser' },
            { link: '/guide/errors', text: 'Error Handling' },
          ],
        },
      ],
    },

    socialLinks: [
      { icon: 'github', link: 'https://github.com/ntnyq/docx-kit' },
    ],
  },
})
