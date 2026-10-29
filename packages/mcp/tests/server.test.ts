import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { createDocxKitError, createPluginLoader } from '@docxkit/core'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterEach, describe, expect, expectTypeOf, it, vi } from 'vitest'
import { createDocxKitServer } from '../src'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'

afterEach(() => vi.restoreAllMocks())

describe('MCP server integration', () => {
  it('connects and renders an advertised built-in plugin through the protocol', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-server-'))
    const server = await createDocxKitServer({ outputDirectory: directory })
    expectTypeOf(server).toEqualTypeOf<McpServer>()
    const client = new Client({ name: 'docx-kit-test', version: '1.0.0' })
    const [clientTransport, serverTransport] =
      InMemoryTransport.createLinkedPair()
    try {
      await server.connect(serverTransport)
      await client.connect(clientTransport)
      const listed = await client.callTool({
        arguments: { filter: 'badge' },
        name: 'list_plugins',
      })
      expect(listed.content).toEqual([
        expect.objectContaining({ text: expect.stringContaining('badge') }),
      ])
      const help = await client.callTool({
        arguments: { pluginName: 'badge' },
        name: 'get_plugin_help',
      })
      expect(help.content).toEqual([
        expect.objectContaining({
          text: expect.stringContaining('usageExample'),
        }),
      ])
      const result = await client.callTool({
        name: 'create_document',
        arguments: {
          outputPath: 'badge.docx',
          schema: {
            content: [
              {
                name: 'badge',
                options: { color: 'info', text: 'New' },
                type: 'plugin',
              },
            ],
          },
        },
      })
      expect(result.isError).not.toBe(true)
      const bytes = await readFile(path.join(directory, 'badge.docx'))
      expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b])
    } finally {
      await client.close()
      await server.close()
      await rm(directory, { force: true, recursive: true })
    }
  })
  it('returns diagnostic fields over the protocol and preserves validation details', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-errors-'))
    const server = await createDocxKitServer({ outputDirectory: directory })
    const client = new Client({ name: 'docx-kit-test', version: '1.0.0' })
    const [clientTransport, serverTransport] =
      InMemoryTransport.createLinkedPair()
    try {
      await server.connect(serverTransport)
      await client.connect(clientTransport)
      const result = await client.callTool({
        name: 'create_document',
        arguments: {
          outputPath: 'invalid.docx',
          schema: {
            styles: { body: { fontSize: 12 } },
            content: [
              { className: 'missing', text: 'Hello', type: 'paragraph' },
            ],
          },
        },
      })
      expect(result.isError).toBe(true)
      expect(result.content).toEqual([
        expect.objectContaining({
          text: expect.stringContaining('DOCXKIT_STYLE_UNKNOWN_CLASS'),
          type: 'text',
        }),
      ])
      const content = result.content
      if (!Array.isArray(content) || content[0]?.type !== 'text') {
        throw new Error('Expected text error content')
      }
      const diagnostic = JSON.parse(content[0].text)
      expect(diagnostic).toMatchObject({
        code: 'STYLE_UNKNOWN_CLASS',
        diagnosticCode: 'DOCXKIT_STYLE_UNKNOWN_CLASS',
        docs: 'https://docx-kit.ntnyq.dev/errors/docxkit_style_unknown_class/',
        message: 'Style class not found: "missing"',
      })
      expect(diagnostic.fix).toBeTruthy()
      expect(diagnostic).not.toHaveProperty('cause')
      expect(diagnostic).not.toHaveProperty('stack')
      const invalid = await client.callTool({
        arguments: { outputPath: 'invalid.docx', schema: {} },
        name: 'create_document',
      })
      expect(invalid.isError).toBe(true)
      expect(invalid.content).toEqual([
        expect.objectContaining({
          text: expect.stringContaining('Schema validation failed:'),
        }),
      ])
      expect(invalid.content).toEqual([
        expect.objectContaining({ text: expect.stringContaining('content:') }),
      ])
    } finally {
      await client.close()
      await server.close()
      await rm(directory, { force: true, recursive: true })
    }
  })
  it.each([
    {
      error: Object.create(null),
      expected: { message: 'Unknown error' },
      name: 'unconvertible exception',
    },
    {
      name: 'credential URL',
      error: createDocxKitError('PLUGIN_LOAD_FAILED', {
        reason: 'import',
        detail:
          'Failed to import https://demo:EXAMPLE_PASSWORD@example.invalid/plugin.js?token=EXAMPLE_TOKEN#EXAMPLE_FRAGMENT',
      }),
      expected: {
        code: 'PLUGIN_LOAD_FAILED',
        diagnosticCode: 'DOCXKIT_PLUGIN_LOAD_FAILED',
        docs: 'https://docx-kit.ntnyq.dev/errors/docxkit_plugin_load_failed/',
        message: 'Failed to import https://example.invalid/plugin.js',
      },
    },
    {
      name: 'nested redirect URL with a trailing access token',
      error: createDocxKitError('PLUGIN_LOAD_FAILED', {
        reason: 'import',
        detail:
          'Failed to import https://example.invalid/plugin.js?redirect=https://callback.invalid/path&access_token=EXAMPLE_TOKEN',
      }),
      expected: {
        code: 'PLUGIN_LOAD_FAILED',
        diagnosticCode: 'DOCXKIT_PLUGIN_LOAD_FAILED',
        docs: 'https://docx-kit.ntnyq.dev/errors/docxkit_plugin_load_failed/',
        message: 'Failed to import https://example.invalid/plugin.js',
      },
    },
  ])(
    'returns a stable safe error response for $name',
    async ({ error, expected }) => {
      const directory = await mkdtemp(
        path.join(tmpdir(), 'docx-kit-mcp-serialize-'),
      )
      const pluginLoader = createPluginLoader()
      vi.spyOn(pluginLoader, 'load').mockRejectedValue(error)
      const server = await createDocxKitServer({
        outputDirectory: directory,
        pluginLoader,
      })
      const client = new Client({ name: 'docx-kit-test', version: '1.0.0' })
      const [clientTransport, serverTransport] =
        InMemoryTransport.createLinkedPair()
      try {
        await server.connect(serverTransport)
        await client.connect(clientTransport)
        const result = await client.callTool({
          name: 'create_document',
          arguments: {
            outputPath: 'invalid.docx',
            schema: {
              content: [{ text: 'Hello', type: 'paragraph' }],
              plugins: [{ package: 'example-plugin', type: 'npm' }],
            },
          },
        })
        expect(pluginLoader.load).toHaveBeenCalled()
        expect(result.isError).toBe(true)
        const content = result.content
        if (!Array.isArray(content) || content[0]?.type !== 'text') {
          throw new Error('Expected text error content')
        }
        expect(JSON.parse(content[0].text)).toMatchObject(expected)
        expect(content[0].text).not.toContain('EXAMPLE_')
        expect(content[0].text).not.toContain('Cannot convert object')
        expect(JSON.parse(content[0].text)).not.toHaveProperty('stack')
      } finally {
        await client.close()
        await server.close()
        await rm(directory, { force: true, recursive: true })
      }
    },
  )
})
