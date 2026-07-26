/**
 * Compile a plugin node — delegates to the registered plugin's `render()`.
 *
 * @module compiler/nodes/compilePlugin
 */

import { DocxKitError } from '@docxkit/types'
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
    throw new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      `Plugin not registered: ${node.name}`,
    )
  }

  const ctx: PluginRenderContext = {
    config: config as DocxKitConfig,
    utils: {
      image: {
        fromDataUrl: (dataUrl: string) => dataUrlToUint8Array(dataUrl),
        fromBlob: async (blob: Blob) =>
          new Uint8Array(await blob.arrayBuffer()),
      },
    },
    compileNode: (childNode: BlockNode) =>
      compileNode({
        config: config as DocxKitConfig,
        node: childNode,
        plugins,
        session,
      }),
  }

  try {
    return (await plugin.render(node.options, ctx)) as FileChild | FileChild[]
  } catch (error) {
    throw new DocxKitError(
      'PLUGIN_RENDER_FAILED',
      `Plugin render failed: ${plugin.name}`,
      error,
    )
  }
}
