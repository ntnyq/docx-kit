import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { badgePlugin } from '../../packages-plugins/badge/src'
import { calloutPlugin } from '../../packages-plugins/callout/src'
import { changelogPlugin } from '../../packages-plugins/changelog/src'
import { checklistPlugin } from '../../packages-plugins/checklist/src'
import { codeBlockPlugin } from '../../packages-plugins/code-block/src'
import { coverPagePlugin } from '../../packages-plugins/cover-page/src'
import { dataTablePlugin } from '../../packages-plugins/data-table/src'
import { dividerPlugin } from '../../packages-plugins/divider/src'
import { figurePlugin } from '../../packages-plugins/figure/src'
import { invoicePlugin } from '../../packages-plugins/invoice/src'
import { letterheadPlugin } from '../../packages-plugins/letterhead/src'
import { meetingMinutesPlugin } from '../../packages-plugins/meeting-minutes/src'
import { metricCardPlugin } from '../../packages-plugins/metric-card/src'
import { propertyTablePlugin } from '../../packages-plugins/property-table/src'
import { timelinePlugin } from '../../packages-plugins/timeline/src'
import { watermarkPlugin } from '../../packages-plugins/watermark/src'
import { createDocx } from '../../packages/core/src'

const image =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='
const config = {
  theme: {
    colors: {
      accent: '#0f0',
      border: '#abc',
      danger: '#f00',
      info: '#00f',
      muted: '#666',
      primary: '#f00',
      success: '#0f0',
      surface: '#eee',
    },
  },
}

describe('short hexadecimal colors across document and plugin boundaries', () => {
  it.each([
    [
      'badge',
      () =>
        createDocx(config).use(badgePlugin()).plugin('badge', {
          backgroundColor: '#eee',
          color: '#f00',
          text: 'Badge',
        }),
    ],
    [
      'callout',
      () =>
        createDocx(config).use(calloutPlugin()).plugin('callout', {
          content: 'Content',
          title: 'Title',
          type: 'info',
        }),
    ],
    [
      'changelog',
      () =>
        createDocx(config)
          .use(changelogPlugin())
          .plugin('changelog', {
            entries: [
              {
                changes: 'Change',
                date: '2026-10-30',
                type: 'added',
                version: '1.0.0',
              },
            ],
          }),
    ],
    [
      'checklist',
      () =>
        createDocx(config)
          .use(checklistPlugin())
          .plugin('checklist', {
            items: [{ group: 'Group', note: 'Note', text: 'Task' }],
            title: 'Tasks',
          }),
    ],
    [
      'codeBlock',
      () =>
        createDocx(config)
          .use(codeBlockPlugin())
          .plugin('codeBlock', { code: 'const x = 1' }),
    ],
    [
      'coverPage',
      () =>
        createDocx(config).use(coverPagePlugin()).plugin('coverPage', {
          author: 'Author',
          backgroundColor: '#eee',
          subtitle: 'Subtitle',
          title: 'Cover',
        }),
    ],
    [
      'dataTable',
      () =>
        createDocx(config)
          .use(dataTablePlugin())
          .plugin('dataTable', { data: [{ x: 1 }, { x: 2 }], striped: true }),
    ],
    [
      'divider',
      () =>
        createDocx(config)
          .use(dividerPlugin())
          .plugin('divider', { color: '#abc' }),
    ],
    [
      'figure',
      () =>
        createDocx(config).use(figurePlugin()).plugin('figure', {
          caption: 'Caption',
          data: image,
          source: 'Source',
        }),
    ],
    [
      'invoice',
      () =>
        createDocx(config)
          .use(invoicePlugin())
          .plugin('invoice', {
            date: '2026-10-30',
            from: { name: 'Issuer' },
            invoiceNumber: 'INV',
            items: [{ description: 'Item', quantity: 1, unitPrice: 1 }],
            to: { name: 'Client' },
          }),
    ],
    [
      'letterhead',
      () =>
        createDocx(config)
          .use(letterheadPlugin())
          .plugin('letterhead', { companyName: 'Company' }),
    ],
    [
      'meetingMinutes',
      () =>
        createDocx(config)
          .use(meetingMinutesPlugin())
          .plugin('meetingMinutes', {
            agenda: [{ discussion: 'Discussion', topic: 'Topic' }],
            attendees: ['Alice'],
            date: '2026-10-30',
            title: 'Meeting',
          }),
    ],
    [
      'metricCard',
      () =>
        createDocx(config)
          .use(metricCardPlugin())
          .plugin('metricCard', {
            items: [{ label: 'Metric', trend: 'up', value: 1 }],
          }),
    ],
    [
      'propertyTable',
      () =>
        createDocx(config)
          .use(propertyTablePlugin())
          .plugin('propertyTable', { items: [{ key: 'Key', value: 'Value' }] }),
    ],
    [
      'timeline',
      () =>
        createDocx(config)
          .use(timelinePlugin())
          .plugin('timeline', {
            accentColor: '#0f0',
            events: [{ date: '2026-10-30', title: 'Event' }],
          }),
    ],
    [
      'watermark',
      () =>
        createDocx(config)
          .use(watermarkPlugin())
          .plugin('watermark', { color: '#abc', text: 'Draft' }),
    ],
    [
      'table',
      () =>
        createDocx(config).table({
          columns: [{ key: 'x', title: 'X' }],
          data: [{ x: 1 }, { x: 2 }],
          striped: true,
        }),
    ],
  ] as const)('packs %s with valid native colors', async (_name, build) => {
    const zip = await JSZip.loadAsync(await build().toBuffer())
    const xml = await zip.file('word/document.xml')!.async('string')
    const colors = [
      ...xml.matchAll(/(?:<w:color w:val|w:(?:fill|color))="([^"]+)"/g),
    ].map(match => match[1])
    expect(colors.length).toBeGreaterThan(0)
    expect(colors.every(color => /^(?:[\da-f]{6}|auto)$/i.test(color))).toBe(
      true,
    )
  })
})
