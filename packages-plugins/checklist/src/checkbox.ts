import { createImageRun, dataUrlToUint8Array } from '@docxkit/core'
import type { DocxTheme } from '@docxkit/core'

// Grayscale SVGs and 4x indexed PNGs share foreground/background color masks.
const CHECKED_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><rect x="1" y="1" width="14" height="14" rx="3" fill="#000"/><path d="m4.5 8 2.25 2.25 4.75-4.75" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>'
const CHECKED_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAMAAACdt4HsAAAAq1BMVEVMaXEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8BAQGWlpY9PT2KiooCAgKampqTk5OJiYmPj4+RkZGUlJQj/zDdAAAALHRSTlMA2PFOHPK1/fP8btcdbB5rG7dt+xm49FCqnpnbSWmiIJGYb2q2FZtMTdmcGuWvG60AAAAJcEhZcwAACxMAAAsTAQCanBgAAAGHSURBVFiF7ZfbcoIwEIZXBUHUarXa8/ncblZoEfv+T9ZxbEOIISTbTqcX/tf5vs1ClgkAu6jZm8xac2zIvDUbDsCU8UkPHRMdXW7zhzF6JOzo/E3kwyNGiVY/8uMRo8oeAq/9bxK+KoJjfx7xQXl/PY6gF0hBwuERR1KwzxN0peCUJ2hJQcwTxFKAzPy9IM2y9AcCKoQQImMLaLHmhUiZAvrixZInoG+eKSDJ81qgki+QIaAPyb+/MQRUw7sKqI43C1LttKn9LwgbBdl64YqceJMg1ZfaeJNgqS228jaB2Cy385YWxAZo4I0PcaUgTbxRQArUxJvPAeUSqzs/dgFuG+r4uqNM5dZt+7fMAuVO9S3DRAuX+rZppNyhvnWcKXfgrd8DKvSx9BTg9ofBV+CQfySIeXxbCp55gv7vXbKGPMG5FAxYF81pedGEF47guuRhHPrz7QNFAGfeTURPKg+Q+P5wTKo8QMeri/ajzgPcX02dy19U+pcJRnf9xlMd97u3yvuDXQA+AWcSBCUGVGoTAAAAAElFTkSuQmCC'

const UNCHECKED_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><rect x="1.625" y="1.625" width="12.75" height="12.75" rx="2.375" fill="#fff" stroke="#000" stroke-width="1.25"/></svg>'
const UNCHECKED_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAMAAACdt4HsAAAAvVBMVEVMaXEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8AAAAVFRXz8/NSUlLx8fGBgYHy8vL9/f3w8PDOzs6Dg4MWFhaCgoL09PTNzc0UFBRUVFTQ0NBQUFD+/v7Pz89RUVG2a/HFAAAAKHRSTlMAa7VpG/vy/fMc8ddOtvRu/E1Q1rTYmRmqb5pJGm2cm2odbCDZkRWidAEVTwAAAAlwSFlzAAALEwAACxMBAJqcGAAAAUtJREFUWIXtl9duwzAMReV4yCNO2mZ07yU6duLETdL9/58VkAbk0QCW+tjqvlHQuSIpCKAYM6orsOJoAB0aRLEVsH1ypi4oip/c/+SPfNCQ12/zp1yHB+DD1vlcjwfgjRycMv+X4msjOrRZFUlZxUXN4IqW3hZCSVlK288q/vaA+E81XoicHNzqNu8of8XzKQeq4kYaxBiu1Xkh5kj0pMEDhu86BiskJtLAx3CmY7BAwpcGgNLhhSDEGIDpAZgeoMxbANMDMD1AmbcAf6oHnvaAkSERSoPJ70acSBr0MCy0h6yxNLAwTDJ1fvmKhC0NAhr001yVz79xv+tUo+o1NTVdKp7/Qdun9WHbo6Vkve28i9l2TvlDeFwf189pWtYRP6zzjA11PxyXTZ6xflmFosLW+ainR/VP13OjfinH7o06v17+aGzv//axf6sdsgbXjq5XMloAAAAASUVORK5CYII='

/**
 * Draw a static GitHub-style checkbox without depending on symbol fonts.
 */
export async function createGithubCheckbox(
  checked: boolean,
  label: string,
  fontSize: number,
  theme?: DocxTheme,
) {
  const colors = theme?.colors
  const foreground = checked
    ? normalizeColor(colors?.accent ?? colors?.primary, '#0969da')
    : normalizeColor(colors?.muted ?? colors?.border, '#8c959f')
  const background = checked
    ? '#ffffff'
    : normalizeColor(colors?.background, '#ffffff')
  const svg = (checked ? CHECKED_SVG : UNCHECKED_SVG).replaceAll(
    /#000|#fff/g,
    color => (color === '#000' ? foreground : background),
  )
  const fallback = await dataUrlToUint8Array(
    checked ? CHECKED_PNG : UNCHECKED_PNG,
  )

  return createImageRun({
    alt: `${checked ? 'Checked' : 'Unchecked'}: ${label}`,
    data: new TextEncoder().encode(svg),
    height: (fontSize * 4) / 3,
    // Inline drawings sit on the baseline; lower them to the body's visual center.
    run: { position: -Math.round(fontSize * 0.3), size: fontSize * 2 },
    type: 'svg',
    width: (fontSize * 4) / 3,
    fallback: {
      data: recolorPng(fallback, foreground, background),
      type: 'png',
    },
  })
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xff_ff_ff_ff
  for (const byte of bytes) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xed_b8_83_20 : 0)
    }
  }
  return (crc ^ 0xff_ff_ff_ff) >>> 0
}

function normalizeColor(color: string | undefined, fallback: string): string {
  const hex = color?.replace(/^#/, '')
  if (!hex || !/^(?:[\da-f]{3}|[\da-f]{6})$/i.test(hex)) {
    return fallback
  }
  return `#${hex.length === 3 ? [...hex].map(value => value.repeat(2)).join('') : hex}`
}

/**
 * Tint the indexed PNG palette without rasterizing SVGs or importing Node APIs.
 * Grayscale palette values interpolate between the two colors, preserving
 * antialiasing and the original transparency table.
 */
function recolorPng(
  bytes: Uint8Array,
  foreground: string,
  background: string,
): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let offset = 8; offset < bytes.length;) {
    const length = view.getUint32(offset)
    // PLTE is the PNG palette chunk; other chunks retain their original bytes.
    if (view.getUint32(offset + 4) === 0x50_4c_54_45) {
      for (let index = offset + 8; index < offset + 8 + length; index += 3) {
        const backgroundWeight = bytes[index] / 255
        for (let channel = 0; channel < 3; channel++) {
          const start = 1 + channel * 2
          const from = Number.parseInt(foreground.slice(start, start + 2), 16)
          const to = Number.parseInt(background.slice(start, start + 2), 16)
          bytes[index + channel] = Math.round(
            from + (to - from) * backgroundWeight,
          )
        }
      }
      view.setUint32(
        offset + 8 + length,
        crc32(bytes.subarray(offset + 4, offset + 8 + length)),
      )
      return bytes
    }
    offset += length + 12
  }
  throw new Error('Checkbox PNG template is missing its palette')
}
