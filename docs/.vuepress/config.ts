import { defineUserConfig } from 'vuepress'
import { defaultTheme } from '@vuepress/theme-default'
import { viteBundler } from '@vuepress/bundler-vite'

export default defineUserConfig({
  lang: 'en-US',
  title: 'xlsx-strictdoc-sync',
  description:
    'Synchronize requirements between Microsoft Excel and StrictDoc (.sdoc) files, ' +
    'with support for bidirectional sync and per-field direction overrides.',

  base: '/xlsx-strictdoc-sync/',

  bundler: viteBundler(),

  theme: defaultTheme({
    logo: null,

    repo: 'the78mole/xlsx-strictdoc-sync',
    repoLabel: 'GitHub',
    docsDir: 'docs',

    editLink: true,
    editLinkText: 'Edit this page on GitHub',

    lastUpdated: true,
    contributors: false,

    navbar: [
      { text: 'Home', link: '/' },
      { text: 'Getting Started', link: '/getting-started' },
      { text: 'Configuration', link: '/configuration' },
      { text: 'Bidirectional Sync', link: '/bidirectional-sync' },
      { text: 'CLI Reference', link: '/cli-reference' },
      { text: 'Architecture', link: '/architecture' },
    ],

    sidebar: [
      { text: 'Home', link: '/' },
      { text: 'Getting Started', link: '/getting-started' },
      { text: 'Configuration', link: '/configuration' },
      { text: 'Bidirectional Sync', link: '/bidirectional-sync' },
      { text: 'CLI Reference', link: '/cli-reference' },
      { text: 'Architecture', link: '/architecture' },
    ],
  }),
})
