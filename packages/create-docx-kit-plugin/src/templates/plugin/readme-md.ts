/**
 * README template — generates a README.md for a plugin project.
 *
 * @module templates/plugin/readme-md
 */

/**
 * Render a README.md for a docx-kit plugin.
 *
 * @param pluginName - — npm package name (e.g. `docx-kit-plugin-chart`)
 * @param shortName - — Plugin name without prefix (e.g. `chart`)
 * @param description - — Plugin description
 * @param author - — Author name
 */
export function renderReadme(
  pluginName: string,
  shortName: string,
  description: string = 'A docx-kit plugin',
  author: string = '',
): string {
  return `# ${pluginName}

${description}

## Install

\`\`\`bash
npm install ${pluginName}
\`\`\`

## Usage

\`\`\`ts
import { createDocx } from 'docx-kit'
import { ${shortName}Plugin } from '${pluginName}'

const doc = createDocx()
  .use(${shortName}Plugin())
  .${shortName}({ text: 'Hello' })

doc.pack()
\`\`\`

## Plugin Manifest

This plugin declares a \`docx-kit.plugin.json\` manifest for compatibility
checking and metadata. See the [Plugin Authoring Guide](https://github.com/ntnyq/docx-kit/docs/ecosystem/creating-plugins.md)
for details.

## License

MIT${author ? ` — ${author}` : ''}
`
}
