# Performance baselines

These benchmarks track generation time, memory use, and DOCX output size for
large or structurally expensive documents. They are informational trend
signals, not pull-request gates: CPU scheduling, Node.js, browser, font, and
operating-system differences make strict cross-machine thresholds unreliable.

## Scenarios

| Scenario           | Coverage                                        |
| ------------------ | ----------------------------------------------- |
| `paragraphs-100`   | Small-document control                          |
| `paragraphs-1000`  | Medium paragraph-heavy document                 |
| `paragraphs-10000` | Large paragraph-heavy document                  |
| `table-1000-rows`  | 1,000 rows × 5 columns                          |
| `images-100`       | Image metadata, media registration, and packing |
| `plugins-500`      | 500 Callout nodes, style adaptation, DOCX packing |

## Commands

Build packages before running the benchmarks:

```sh
pnpm run build
pnpm run bench:node
pnpm run bench:browser
```

The Node benchmark runs each scenario in an isolated process three times and
reports median duration, maximum peak RSS, and median output size. Set
`DOCX_KIT_BENCH_RUNS=1` for a quick local smoke run. Update the committed
machine baseline only after an intentional performance review:

```sh
pnpm run bench:node:update
pnpm run bench:browser:update
```

The browser benchmark also runs each scenario three times, starts a temporary
Vite server, and controls a local headless Chrome/Chromium instance through the
DevTools protocol. Set `CHROME_PATH` when the browser is not installed in a
standard location. Its results are written to
`output/performance/browser-latest.json`; Chromium heap delta is included when
the browser exposes `performance.memory`.

When reviewing a change, compare like-for-like hardware and runtime versions.
Investigate sustained regressions across repeated runs before adding a budget
or making the benchmark blocking in CI. The non-blocking `Performance`
workflow runs weekly and on manual dispatch, then retains its JSON reports as a
workflow artifact for 30 days.

Compare the generated work as well as the runtime. The original `plugins-500`
baseline used paragraph-based Callouts; the current plugin renders a single-cell
table for each Callout. That adds table, row, cell, border, and margin XML, so its
total duration is not a direct measure of plugin dispatch overhead. Do not remove
layout features or replace a baseline just to eliminate that difference.

For diagnosis, replay historical code with the same installed dependencies and
alternate old/new runs in isolated processes. Separate builder construction,
`toDocument()` compilation, and DOCX packing, then compare uncompressed document
and style XML. Matching XML helps distinguish implementation overhead from a
change in the amount of generated work. Recorded environment metadata alone
cannot attribute a timing difference to Node.js, dependencies, or the OS.
