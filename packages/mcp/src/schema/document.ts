import { STYLE_RULE_SCHEMA, UNIT_VALUE_SCHEMA } from './styles'

const nonEmptyString = { minLength: 1, type: 'string' }

/**
 * Serializable page options shared by the resource and runtime validator.
 */
export const PAGE_SCHEMA = {
  type: 'object',
  properties: {
    footerDistance: UNIT_VALUE_SCHEMA,
    gutter: UNIT_VALUE_SCHEMA,
    headerDistance: UNIT_VALUE_SCHEMA,
    margin: STYLE_RULE_SCHEMA.properties.margin,
    orientation: { enum: ['landscape', 'portrait'], type: 'string' },
    borders: {
      type: 'object',
      properties: {
        bottom: STYLE_RULE_SCHEMA.properties.border,
        display: { enum: ['allPages', 'firstPage', 'notFirstPage'] },
        left: STYLE_RULE_SCHEMA.properties.border,
        offsetFrom: { enum: ['page', 'text'] },
        right: STYLE_RULE_SCHEMA.properties.border,
        top: STYLE_RULE_SCHEMA.properties.border,
        zOrder: { enum: ['back', 'front'] },
      },
    },
    pageNumber: {
      type: 'object',
      properties: {
        start: { minimum: 0, type: 'integer' },
        format: {
          enum: [
            'decimal',
            'lowerLetter',
            'lowerRoman',
            'upperLetter',
            'upperRoman',
          ],
        },
      },
    },
    size: {
      anyOf: [
        { enum: ['A3', 'A4', 'Legal', 'Letter'] },
        {
          properties: { height: UNIT_VALUE_SCHEMA, width: UNIT_VALUE_SCHEMA },
          required: ['height', 'width'],
          type: 'object',
        },
      ],
    },
  },
}

/**
 * External sources remain declarative; inline plugin instances are for
 * in-process callers and are not made executable by JSON validation.
 */
export const PLUGIN_SOURCES_SCHEMA = {
  type: 'array',
  items: {
    oneOf: [
      ...Object.entries({ local: 'path', npm: 'package', url: 'url' }).map(
        ([type, field]) => ({
          properties: { [field]: nonEmptyString, type: { const: type } },
          required: ['type', field],
          type: 'object',
        }),
      ),
      {
        required: ['plugin', 'type'],
        type: 'object',
        properties: {
          type: { const: 'inline' },
          plugin: {
            properties: { name: nonEmptyString },
            required: ['name', 'render'],
            type: 'object',
          },
        },
      },
    ],
  },
}

/**
 * Validate these options independently of recursive content, whose depth is
 * bounded by the node validator before traversing in-process cyclic objects.
 */
export const DOCUMENT_OPTIONS_SCHEMA = {
  properties: { page: PAGE_SCHEMA, plugins: PLUGIN_SOURCES_SCHEMA },
  type: 'object',
}
