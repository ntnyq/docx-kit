import { spawn } from 'node:child_process'
import { once } from 'node:events'
import {
  access,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, resolve } from 'node:path'
import process from 'node:process'
import { setTimeout as delay } from 'node:timers/promises'
import { createServer } from 'vite'

const repositoryRoot = resolve(import.meta.dirname, '../..')
const browserRoot = resolve(repositoryRoot, 'benchmarks/browser')
const outputPath = resolve(
  repositoryRoot,
  'output/performance/browser-latest.json',
)
const baselinePath = resolve(
  repositoryRoot,
  'benchmarks/baselines/browser.json',
)
const updateBaseline = process.argv.includes('--update-baseline')
const profileDirectory = await mkdtemp(
  resolve(tmpdir(), 'docx-kit-browser-benchmark-'),
)
const chromePath = await findChrome()
const server = await createServer({
  configFile: false,
  logLevel: 'error',
  root: browserRoot,
  server: {
    host: '127.0.0.1',
    port: 0,
  },
})

let chrome

try {
  await server.listen()
  const benchmarkUrl = server.resolvedUrls?.local[0]
  if (!benchmarkUrl) {
    throw new Error('Vite did not expose a local benchmark URL')
  }

  chrome = spawn(
    chromePath,
    [
      '--disable-background-networking',
      '--disable-extensions',
      '--disable-gpu',
      '--headless=new',
      '--no-default-browser-check',
      '--no-first-run',
      '--remote-debugging-port=0',
      `--user-data-dir=${profileDirectory}`,
      benchmarkUrl,
    ],
    { stdio: ['ignore', 'ignore', 'pipe'] },
  )

  const portFile = resolve(profileDirectory, 'DevToolsActivePort')
  await waitForFile(portFile)
  const [port] = (await readFile(portFile, 'utf8')).trim().split('\n')
  const target = await waitForTarget(Number(port), benchmarkUrl)
  const report = await readBenchmarkReport(target.webSocketDebuggerUrl)

  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`)
  if (updateBaseline) {
    await mkdir(dirname(baselinePath), { recursive: true })
    await writeFile(baselinePath, `${JSON.stringify(report, null, 2)}\n`)
  }

  console.table(
    report.scenarios.map(scenario => ({
      duration: `${scenario.durationMedianMs.toFixed(1)} ms`,
      heapDelta: formatBytes(scenario.heapDeltaMaxBytes),
      output: formatBytes(scenario.outputMedianBytes),
      scenario: scenario.scenario,
    })),
  )
  console.log(`Report: ${outputPath}`)
  if (updateBaseline) {
    console.log(`Baseline updated: ${baselinePath}`)
  }
} finally {
  await stopChrome(chrome)
  await server.close()
  await rm(profileDirectory, {
    force: true,
    maxRetries: 5,
    recursive: true,
    retryDelay: 100,
  })
}

function connectDevTools(webSocketUrl) {
  return new Promise((resolveConnection, rejectConnection) => {
    const socket = new WebSocket(webSocketUrl)
    const pending = new Map()
    let nextId = 1

    socket.addEventListener('error', rejectConnection, { once: true })
    socket.addEventListener(
      'open',
      () => {
        socket.addEventListener('message', event => {
          const message = JSON.parse(event.data)
          const request = pending.get(message.id)
          if (!request) {
            return
          }
          pending.delete(message.id)
          if (message.error) {
            request.reject(new Error(message.error.message))
          } else {
            request.resolve(message.result)
          }
        })

        resolveConnection({
          close: () => socket.close(),
          send(method, params = {}) {
            return new Promise((resolveRequest, rejectRequest) => {
              const id = nextId
              nextId += 1
              pending.set(id, {
                reject: rejectRequest,
                resolve: resolveRequest,
              })
              socket.send(JSON.stringify({ id, method, params }))
            })
          },
        })
      },
      { once: true },
    )
  })
}

async function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
  ].filter(Boolean)

  for (const candidate of candidates) {
    try {
      await access(candidate)
      return candidate
    } catch {
      // Try the next known executable.
    }
  }

  throw new Error(
    'Chrome or Chromium was not found. Set CHROME_PATH to run the browser benchmark.',
  )
}

function formatBytes(bytes) {
  const units = ['B', 'KiB', 'MiB', 'GiB']
  let value = bytes
  let unitIndex = 0
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024
    unitIndex += 1
  }
  return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`
}

async function readBenchmarkReport(webSocketUrl) {
  const connection = await connectDevTools(webSocketUrl)
  const timeoutAt = Date.now() + 120_000

  try {
    await connection.send('Runtime.enable')

    while (Date.now() < timeoutAt) {
      const evaluation = await connection.send('Runtime.evaluate', {
        returnByValue: true,
        expression:
          "document.querySelector('#benchmark-result')?.dataset.report ?? ''",
      })
      const encodedReport = evaluation.result?.value
      if (encodedReport) {
        return JSON.parse(decodeURIComponent(encodedReport))
      }
      await new Promise(resolveWait => setTimeout(resolveWait, 100))
    }
  } finally {
    connection.close()
  }

  throw new Error('Browser benchmark did not complete within 120 seconds')
}

async function stopChrome(childProcess) {
  if (!childProcess || childProcess.exitCode != null) {
    return
  }

  const exited = once(childProcess, 'exit')
  childProcess.kill('SIGTERM')
  await Promise.race([exited, delay(5_000)])

  if (childProcess.exitCode == null) {
    const forcedExit = once(childProcess, 'exit')
    childProcess.kill('SIGKILL')
    await forcedExit
  }
}

async function waitForFile(path) {
  const timeoutAt = Date.now() + 15_000
  while (Date.now() < timeoutAt) {
    try {
      await access(path)
      return
    } catch {
      await new Promise(resolveWait => setTimeout(resolveWait, 50))
    }
  }
  throw new Error(`Timed out waiting for ${path}`)
}

async function waitForTarget(port, benchmarkUrl) {
  const timeoutAt = Date.now() + 15_000
  while (Date.now() < timeoutAt) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`)
      const targets = await response.json()
      const target = targets.find(
        item => item.type === 'page' && item.url.startsWith(benchmarkUrl),
      )
      if (target) {
        return target
      }
    } catch {
      // Chrome may still be starting.
    }
    await new Promise(resolveWait => setTimeout(resolveWait, 50))
  }
  throw new Error('Timed out waiting for the Chrome benchmark page')
}
