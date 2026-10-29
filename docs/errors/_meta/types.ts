import type { ErrorCode, ErrorParams } from '../../../packages/errors/src/index'

export interface ErrorExample {
  after: string
  before: string
  guide: string
}
export interface ErrorNote {
  causes: string[]
  steps: string[]
  summary: string
  title: string
}
export type ErrorNotes = Record<ErrorCode, ErrorNote>
export type ErrorSamples = { [Code in ErrorCode]: ErrorParams<Code> }
