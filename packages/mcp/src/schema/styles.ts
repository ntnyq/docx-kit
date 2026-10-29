import type { DocxStyleRule } from '@docxkit/core'

const numberPattern = '-?(?:\\d+(?:\\.\\d+)?|\\.\\d+)(?:[eE][+-]?\\d+)?'
const tokenPattern = '\\$(?:colors|fonts|fontSize|spacing)\\.[^\\s]+'
const lengthPattern = `(?:${numberPattern}(?:%|cm|in|mm|pt|px)|${tokenPattern})`
const shorthandPartPattern = `(?:${numberPattern}|${lengthPattern})`

/**
 * Shared style contracts for the MCP resource and runtime validation.
 */
export const UNIT_VALUE_SCHEMA = {
  anyOf: [
    { type: 'number' },
    { pattern: `^${lengthPattern}$`, type: 'string' },
  ],
}

const string = { type: 'string' }
const boolean = { type: 'boolean' }
const number = { type: 'number' }
const margin = {
  anyOf: [
    UNIT_VALUE_SCHEMA,
    {
      pattern: `^${shorthandPartPattern}(?:\\s+${shorthandPartPattern})(?:(?:\\s+${shorthandPartPattern}){2})?$`,
      type: 'string',
    },
  ],
}
const border = {
  additionalProperties: false,
  type: 'object',
  properties: {
    color: string,
    style: { enum: ['dashed', 'dotted', 'double', 'none', 'single'] },
    width: UNIT_VALUE_SCHEMA,
  },
}

const properties = {
  allCaps: boolean,
  backgroundColor: string,
  bold: boolean,
  border,
  borderBottom: border,
  borderLeft: border,
  borderRight: border,
  borderTop: border,
  characterSpacing: number,
  color: string,
  docx: { additionalProperties: true, type: 'object' },
  doubleStrike: boolean,
  emboss: boolean,
  fontFamily: string,
  fontFamilyEastAsia: string,
  fontSize: UNIT_VALUE_SCHEMA,
  fontStyle: { enum: ['italic', 'normal'] },
  height: UNIT_VALUE_SCHEMA,
  imprint: boolean,
  italic: boolean,
  keepLines: boolean,
  keepNext: boolean,
  letterSpacing: UNIT_VALUE_SCHEMA,
  lineHeight: UNIT_VALUE_SCHEMA,
  margin,
  marginBottom: UNIT_VALUE_SCHEMA,
  marginLeft: UNIT_VALUE_SCHEMA,
  marginRight: UNIT_VALUE_SCHEMA,
  marginTop: UNIT_VALUE_SCHEMA,
  outlineLevel: number,
  pageBreakBefore: boolean,
  rightToLeft: boolean,
  smallCaps: boolean,
  strike: boolean,
  subScript: boolean,
  superScript: boolean,
  textAlign: { enum: ['center', 'justify', 'left', 'right'] },
  textIndent: UNIT_VALUE_SCHEMA,
  underline: { anyOf: [boolean, { enum: ['double', 'single'] }] },
  verticalAlign: { enum: ['bottom', 'middle', 'top'] },
  widowControl: boolean,
  width: UNIT_VALUE_SCHEMA,
  fontWeight: {
    enum: ['bold', 'normal', 100, 200, 300, 400, 500, 600, 700, 800, 900],
  },
  highlight: {
    enum: [
      'black',
      'blue',
      'cyan',
      'darkBlue',
      'darkCyan',
      'darkGray',
      'darkGreen',
      'darkMagenta',
      'darkRed',
      'darkYellow',
      'green',
      'lightGray',
      'magenta',
      'none',
      'red',
      'white',
      'yellow',
    ],
  },
  tabStops: {
    type: 'array',
    items: {
      additionalProperties: false,
      required: ['position', 'type'],
      type: 'object',
      properties: {
        leader: { enum: ['dot', 'hyphen', 'middleDot', 'none', 'underscore'] },
        position: UNIT_VALUE_SCHEMA,
        type: {
          enum: [
            'bar',
            'center',
            'clear',
            'decimal',
            'end',
            'left',
            'num',
            'right',
            'start',
          ],
        },
      },
    },
  },
} satisfies Record<keyof DocxStyleRule, Record<string, unknown>>

export const STYLE_RULE_SCHEMA = {
  additionalProperties: true,
  description: 'CSS-like style properties for a node',
  type: 'object',
  properties: {
    ...properties,
    extends: {
      anyOf: [string, { items: string, type: 'array' }],
    },
  },
}
