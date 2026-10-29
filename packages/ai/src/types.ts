/**
 * AI template type definitions.
 *
 * Each template provides field definitions for its parameters,
 * a system prompt for LLMs, and a generator that produces
 * a DocxSchema from filled parameters.
 *
 * @module ai/types
 */

import type { DocxSchema } from '@docxkit/core'

/**
 * An AI template for generating documents.
 *
 * Templates provide a structured way for LLMs to generate
 * docx-kit documents. Each template defines:
 * - Field definitions for validation and JSON Schema conversion
 * - A system prompt that guides the LLM
 * - A generator that maps parameters → DocxSchema
 *
 * @template TParams — The template's parameter type
 */
export interface AiTemplate<TParams extends object = Record<string, unknown>> {
  /**
   * Human-readable description (shown to LLMs and in registries).
   */
  description: string
  /**
   * Unique template name (used as identifier).
   */
  name: string
  /**
   * Parameter field definitions. Use `toTemplateJsonSchema()` when a consumer
   * expects standard JSON Schema rather than per-property required markers.
   */
  schema: AiTemplateSchema
  /**
   * System prompt for LLMs generating this document type.
   */
  systemPrompt: string
  /**
   * Generate a DocxSchema from filled template parameters.
   *
   * @param params - Filled input parameters for this document template
   * @returns The document schema generated from the template parameters
   */
  generate: (params: TParams) => DocxSchema
}

export interface AiTemplateArraySchema extends AiTemplateSchemaBase {
  /**
   * Nested item schema.
   */
  items: AiTemplateSchema
  /**
   * Schema type.
   */
  type: 'array'
}

export interface AiTemplateBooleanSchema extends AiTemplateSchemaBase {
  /**
   * Schema type.
   */
  type: 'boolean'
}

/**
 * Metadata about a template (without the generator).
 *
 * Used for listing templates in MCP tools and registries.
 */
export interface AiTemplateInfo {
  /**
   * Template description.
   */
  description: string
  /**
   * Template name.
   */
  name: string
  /**
   * Parameter field definitions, convertible with `toTemplateJsonSchema()`.
   */
  schema: AiTemplateSchema
  /**
   * System prompt for LLMs.
   */
  systemPrompt: string
}

export interface AiTemplateNumberSchema extends AiTemplateSchemaBase {
  /**
   * Schema type.
   */
  type: 'number'
}

export interface AiTemplateObjectSchema extends AiTemplateSchemaBase {
  /**
   * Schema properties (name → type definition).
   */
  properties: Record<string, AiTemplateSchema>
  /**
   * Schema type.
   */
  type: 'object'
  /**
   * Schema title.
   */
  title?: string
}

/**
 * Field definitions used for AI template validation and JSON Schema conversion.
 */
export type AiTemplateSchema =
  | AiTemplateArraySchema
  | AiTemplateBooleanSchema
  | AiTemplateNumberSchema
  | AiTemplateObjectSchema
  | AiTemplateStringSchema

export interface AiTemplateStringSchema extends AiTemplateSchemaBase {
  /**
   * Schema type.
   */
  type: 'string'
  /**
   * Enum values for string fields.
   */
  enum?: string[]
}

/**
 * Options for `generateToolDefinitions()`.
 */
export interface GenerateToolDefinitionsOptions {
  /**
   * List of built-in plugin names to include in tool definitions.
   */
  plugins?: string[]
  /**
   * Template name to include in tool definitions.
   */
  template?: string
}

/**
 * Shared metadata for template parameter field definitions.
 *
 * Used by LLM function calling and MCP tool input validation.
 */
interface AiTemplateSchemaBase {
  /**
   * Property description (helps LLMs understand the field).
   */
  description?: string
  /**
   * Whether this property is required.
   */
  required?: boolean
}
