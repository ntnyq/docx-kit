/**
 * Node compiler registry — open/closed dispatch for node compilation.
 *
 * Instead of a switch statement over `node.type`, each node type registers
 * its compiler function in a registry. This enables extensibility — plugins
 * and future versions can register new node types without modifying the
 * core dispatch logic.
 *
 * @module compiler/registry
 */

import { DocxKitError } from '@docxkit/types'
import type { BlockNode, DocxKitConfig, DocxPlugin } from '@docxkit/types'
import type { FileChild } from 'docx'
import type { CompilationSession } from './numbers'

/**
 * The full context available during node compilation.
 *
 * Created once per compilation and reused across all node compilers.
 */
export interface CompileNodeCtx {
  config: DocxKitConfig
  plugins: Map<string, DocxPlugin>
  session: CompilationSession
}

/**
 * A compiler function for a specific node type.
 *
 * Receives the node (can narrow to its specific type internally),
 * plus the full compilation context.
 */
export type NodeCompilerFn = (
  node: BlockNode,
  ctx: CompileNodeCtx,
) => Promise<FileChild | FileChild[]>

/**
 * Registry of node compilers keyed by `node.type` string.
 *
 * Satisfies the Open/Closed Principle: new node types can be registered
 * without modifying the core dispatch logic.
 */
export class NodeCompilerRegistry {
  private readonly _compilers = new Map<string, NodeCompilerFn>()

  /**
   * Compile a node by dispatching to the registered compiler.
   *
   * @param node - Block node to compile
   * @param ctx - Document configuration, plugins, and compilation session
   * @returns A promise that resolves to one or more compiled document children
   * @throws {DocxKitError} `UNKNOWN_NODE_TYPE` if no compiler is registered for the node
   * @throws Propagates errors thrown or rejected by the registered compiler
   */
  async compile(
    node: BlockNode,
    ctx: CompileNodeCtx,
  ): Promise<FileChild | FileChild[]> {
    const compiler = this._compilers.get(node.type)
    if (!compiler) {
      throw new DocxKitError(
        'UNKNOWN_NODE_TYPE',
        `Unknown node type: ${node.type}`,
      )
    }
    return compiler(node, ctx)
  }

  /**
   * Check if a node type has a registered compiler.
   *
   * @param nodeType - Node type discriminator to look up
   * @returns Whether a compiler is registered for the node type
   */
  has(nodeType: string): boolean {
    return this._compilers.has(nodeType)
  }

  /**
   * Register a compiler function for a given node type.
   *
   * @param nodeType - Node type discriminator to register
   * @param compiler - Compiler invoked for nodes of this type
   * @returns This registry instance for chaining
   */
  register(nodeType: string, compiler: NodeCompilerFn): this {
    this._compilers.set(nodeType, compiler)
    return this
  }
}
