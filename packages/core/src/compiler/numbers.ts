/**
 * Compilation session — scoped numbering state for a single compilation.
 *
 * Replaces the old module-level globals (`numberingConfigMap`, `numberingCounter`)
 * with a per-compilation session object, eliminating race conditions when
 * multiple documents are compiled concurrently.
 *
 * @module compiler/numbers
 */

import type { CommentNode, FootnoteNode } from '@docxkit/types'
import type { ILevelsOptions } from 'docx'

/**
 * Registered comment definition.
 */
export interface CommentDefinition {
  id: number
  node: CommentNode
}

/**
 * Registered footnote definition.
 */
export interface FootnoteDefinition {
  id: number
  node: FootnoteNode
}

/**
 * Numbering config entry shape — mirrors `INumberingOptions.config[number]`.
 */
export type NumberingConfigEntry = {
  levels: readonly ILevelsOptions[]
  reference: string
}

/**
 * Per-compilation session holding mutable numbering state.
 *
 * Create one instance per `compileDocument()` call and pass it through
 * the compilation context.
 */
export class CompilationSession {
  /**
   * Get the number of registered numbering configurations.
   *
   * @returns The number of registered numbering configurations
   */
  get size(): number {
    return this._map.size
  }
  private _commentCounter = 0
  /**
   * Document comments registered while compiling inline content.
   */
  private readonly _comments: CommentDefinition[] = []
  private _counter = 0

  private _footnoteCounter = 0

  /**
   * Document footnotes registered while compiling inline content.
   */
  private readonly _footnotes: FootnoteDefinition[] = []

  /**
   * Accumulated numbering configs keyed by reference string.
   */
  private readonly _map = new Map<string, NumberingConfigEntry>()

  /**
   * Return registered comment definitions.
   *
   * @returns The registered comment definitions in registration order
   */
  getComments(): readonly CommentDefinition[] {
    return this._comments
  }

  /**
   * Return registered footnote definitions.
   *
   * @returns The registered footnote definitions in registration order
   */
  getFootnotes(): readonly FootnoteDefinition[] {
    return this._footnotes
  }

  /**
   * Register a new numbering entry and return its unique reference.
   *
   * @param prefix - Prefix for the generated numbering reference
   * @param entry - Numbering levels and options without a reference
   * @returns A unique numbering reference for this compilation session
   */
  register(
    prefix: string,
    entry: Omit<NumberingConfigEntry, 'reference'>,
  ): string {
    const ref = `${prefix}-${++this._counter}`
    this._map.set(ref, { ...entry, reference: ref })
    return ref
  }

  /**
   * Register a document comment and return its range ID.
   *
   * @param node - Comment range and document-level body to register
   * @returns The zero-based ID assigned to the comment range
   */
  registerComment(node: CommentNode): number {
    const id = this._commentCounter++
    this._comments.push({ id, node })
    return id
  }

  /**
   * Register a footnote and return its reference ID.
   *
   * @param node - Footnote content to register
   * @returns The one-based ID assigned to the footnote reference
   */
  registerFootnote(node: FootnoteNode): number {
    const id = ++this._footnoteCounter
    this._footnotes.push({ id, node })
    return id
  }

  /**
   * Get the accumulated numbering configs for `Document({ numbering })`.
   *
   * @returns Registered numbering configurations, or `undefined` if none exist
   */
  toArray(): NumberingConfigEntry[] | undefined {
    return this._map.size > 0 ? Array.from(this._map.values()) : undefined
  }
}

// -----------------------------------------------------------------
// Backward-compatible module-level exports (deprecated).
//
// These exist only so existing tests that import numberingConfigMap /
// resetNumberingState directly continue to work. New code should use
// `CompilationSession` instead.
// -----------------------------------------------------------------

/**
 * @deprecated Use `CompilationSession` instead. Retained for test compatibility.
 */
export const numberingConfigMap = new Map<string, NumberingConfigEntry>()

/**
 * Clear the legacy numbering configuration map.
 *
 * @deprecated Use `CompilationSession` instead. Retained for test compatibility.
 */
export function resetNumberingState() {
  numberingConfigMap.clear()
}
