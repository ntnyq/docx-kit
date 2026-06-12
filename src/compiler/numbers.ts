/**
 * Compilation session — scoped numbering state for a single compilation.
 *
 * Replaces the old module-level globals (`numberingConfigMap`, `numberingCounter`)
 * with a per-compilation session object, eliminating race conditions when
 * multiple documents are compiled concurrently.
 *
 * @module compiler/numbers
 */

import type { ILevelsOptions } from 'docx'

/** Numbering config entry shape — mirrors `INumberingOptions.config[number]`. */
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
  /** Check if any numbering configs have been registered. */
  get size(): number {
    return this._map.size
  }
  private _counter = 0

  /** Accumulated numbering configs keyed by reference string. */
  private readonly _map = new Map<string, NumberingConfigEntry>()

  /** Register a new numbering entry and return its unique reference. */
  register(
    prefix: string,
    entry: Omit<NumberingConfigEntry, 'reference'>,
  ): string {
    const ref = `${prefix}-${++this._counter}`
    this._map.set(ref, { ...entry, reference: ref })
    return ref
  }

  /** Get the accumulated numbering configs for `Document({ numbering })`. */
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

/** @deprecated Use `CompilationSession` instead. Retained for test compatibility. */
export const numberingConfigMap = new Map<string, NumberingConfigEntry>()

/** @deprecated Use `CompilationSession` instead. Retained for test compatibility. */
export function resetNumberingState() {
  numberingConfigMap.clear()
}
