/**
 * Compile a plugin node — delegates to the registered plugin's `render()`.
 *
 * @module compiler/nodes/compilePlugin
 */

import { createDocxKitError } from '@docxkit/errors'
import { resolveStyle } from '../../style/normalizeStyle'
import { dataUrlToUint8Array } from '../../utils/dataUrl'
import { compileNode } from '../compileNode'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  PluginNode,
  PluginRenderContext,
  StyleSheet,
} from '@docxkit/types'
import type { FileChild } from 'docx'
import type { CompilationSession } from '../numbers'

export async function compilePlugin<TStyles extends StyleSheet>(
  node: PluginNode<string, unknown, TStyles>,
  plugins: Map<string, DocxPlugin>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
): Promise<FileChild | FileChild[]> {
  const plugin = plugins.get(node.name)
  if (!plugin) {
    throw createDocxKitError('PLUGIN_NOT_REGISTERED', { pluginName: node.name })
  }

  const ctx: PluginRenderContext = {
    config: config as DocxKitConfig,
    style: resolveStyle({
      className: node.className,
      inline: node.style,
      styles: config.styles,
      theme: config.theme,
    }),
    utils: {
      image: {
        fromDataUrl: (dataUrl: string) => dataUrlToUint8Array(dataUrl),
        async fromBlob(blob: Blob) {
          return new Uint8Array(await blob.arrayBuffer())
        },
      },
    },
    compileNode(childNode: BlockNode) {
      return compileNode({
        config: config as DocxKitConfig,
        node: childNode,
        plugins,
        session,
      })
    },
  }

  try {
    return (await plugin.render(node.options, ctx)) as FileChild | FileChild[]
  } catch (error) {
    throw createDocxKitError(
      'PLUGIN_RENDER_FAILED',
      { pluginName: plugin.name },
      { cause: error },
    )
  }
}
