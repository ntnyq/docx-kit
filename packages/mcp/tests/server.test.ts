import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { describe, expect, expectTypeOf, it } from 'vitest'
import { createDocxKitServer } from '../src'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'

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
})
