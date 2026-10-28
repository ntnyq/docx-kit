# 範例：學術預設論文

使用 `academicPreset` 產生完整學術文件，採用 Times New Roman、雙倍行距與兩端對齊，適合正式論文。

## 完整程式碼

```ts
import {
  createDocx,
  academicPreset,
  pageNumberPlugin,
  codeBlockPlugin,
} from 'docx-kit/node'

const doc = createDocx({
  ...academicPreset.config,
  metadata: {
    title: 'A Novel Approach to Distributed Cache Coherence',
    creator: 'Jane Researcher',
    subject: 'Computer Science — Distributed Systems',
    keywords: ['distributed systems', 'cache coherence', 'consensus'],
  },
  page: {
    size: 'A4',
    margin: '1in',
  },
})

doc
  // ── Title ──
  .h1('A Novel Approach to Distributed Cache Coherence')
  .p('Jane Researcher', { style: { textAlign: 'center', fontStyle: 'italic' } })
  .p('Department of Computer Science', {
    style: { textAlign: 'center', color: '#555' },
  })
  .pageBreak()

  // ── Abstract ──
  .h1('Abstract')
  .p(
    'This paper proposes a novel approach to maintaining cache coherence in distributed systems. We introduce the Coherence-Vector Protocol (CVP), which combines vector clocks with a gossip-based invalidation scheme to achieve eventual consistency with bounded staleness. Our evaluation shows that CVP reduces coordination overhead by 40% compared to traditional lease-based approaches, while maintaining strong consistency guarantees for critical sections.',
  )

  .h1('1. Introduction')
  .p(
    'Distributed cache coherence is a fundamental challenge in modern cloud-native systems. As applications scale horizontally across data centers, maintaining consistency across thousands of cache replicas becomes increasingly complex. Traditional approaches, such as lease-based invalidation [1] and write-through protocols [2], suffer from high coordination overhead and limited scalability.',
  )

  .p(
    'In this paper, we propose the Coherence-Vector Protocol (CVP), a novel approach that combines vector clocks with gossip-based invalidation to achieve eventual consistency with bounded staleness. Our key contributions are:',
  )

  .numberedList([
    'A formal model of cache coherence in distributed systems',
    'The CVP protocol design and algorithm',
    'An evaluation showing 40% reduction in coordination overhead',
    'A reference implementation in Go',
  ])

  // ── Related work ──
  .h1('2. Related Work')
  .p(
    'Lease-based cache invalidation, originally proposed by [1], has been widely adopted in commercial systems. However, lease renewal requires a round-trip to a coordinator, which becomes a bottleneck in geo-distributed deployments.',
  )

  .p(
    'Dynamo [3] uses a gossip-based protocol for replication but does not address cache coherence directly. Our work extends the gossip approach to cache invalidation while maintaining strong consistency for critical sections.',
  )

  // ── Method ──
  .h1('3. The Coherence-Vector Protocol')
  .p(
    'CVP combines three key ideas: (1) per-replica vector clocks tracking the last-known write timestamp, (2) periodic gossip exchanges of invalidation hints, and (3) lease-free read coordination through bounded staleness windows.',
  )

  .h2('3.1 System Model')
  .p(
    'We assume a system of N replicas, each holding a complete or partial copy of the keyspace. Writes are serialized through a coordinator, while reads can be served by any replica within the staleness window W.',
  )

  .h2('3.2 Protocol Details')
  .p('The core protocol is shown in Algorithm 1.')

  // Code block via plugin
  .use(codeBlockPlugin())
  .plugin('codeBlock', {
    language: 'typescript',
    code: `// CVP: Coherence-Vector Protocol
interface VectorClock {
  replicaId: string
  timestamp: number
}

async function handleWrite(key: string, value: unknown) {
  const clock = await coordinator.nextTimestamp()
  await broadcast({ type: 'write', key, value, clock })
}

async function handleRead(key: string): Promise<unknown> {
  const replica = await selectReplica()
  return replica.get(key)  // Bounded staleness: returns value with clock >= W
}`,
    showLineNumbers: true,
    caption: 'Algorithm 1: CVP core operations',
  })

  // ── Evaluation ──
  .h1('4. Evaluation')
  .p(
    'We evaluated CVP on a 16-node cluster with mixed read/write workloads. Results show a 40% reduction in coordination overhead compared to lease-based approaches, with no measurable increase in read latency.',
  )

  .h2('4.1 Throughput')
  .p(
    'Under 70/30 read/write workload, CVP sustained 12,400 ops/sec compared to 8,800 ops/sec for the baseline — a 41% improvement.',
  )

  .h2('4.2 Latency')
  .p(
    'Read tail latency (p99) at 8.2ms was within 5% of the baseline, while write tail latency improved by 18% due to reduced coordination.',
  )

  // ── Conclusion ──
  .h1('5. Conclusion')
  .p(
    'We have presented CVP, a novel protocol for distributed cache coherence that achieves strong consistency with bounded staleness. Our evaluation demonstrates significant performance improvements over lease-based approaches, making CVP a practical choice for geo-distributed systems.',
  )

  .h2('References')
  .numberedList([
    'G. DeCandia et al., "Dynamo: Amazon\'s Highly Available Key-value Store," SOSP 2007.',
    'J. Gray and L. Lamport, "Consensus on Transaction Commit," ACM TODS, 2006.',
    'A. Lakshman and P. Malik, "Cassandra: A Decentralized Structured Storage System," LADIS 2010.',
  ])

  // ── Footer with page numbers ──
  .section({
    footer: {
      default: {
        children: [
          {
            type: 'plugin',
            name: 'pageNumber',
            options: { showTotal: true, format: 'Page %current% of %total%' },
          },
        ],
      },
    },
  })
  .use(pageNumberPlugin())

await doc.save('thesis.docx')
```

## 預設提供的能力

`academicPreset` 自動提供：

- 全文使用 **Times New Roman**
- 正文采用**雙倍行距**（2.0×）
- 段落**兩端對齊**
- **24pt** 首行縮進（約 2 個字符）
- **h1** 居中、加粗、16pt
- **h4** 加粗並傾斜
- 圖片**居中**，垂直外邊距為 12pt

這些設置符合常見的大學論文排版要求。

## 變體

### 自訂引用樣式

```ts
import { createDocx, academicPreset, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...academicPreset.config,
  styles: {
    ...academicPreset.config.styles,
    citation: {
      fontSize: 10,
      lineHeight: 1.0,
      textIndent: '-24pt', // hanging indent
      marginLeft: '24pt',
    },
  },
})

doc.p('[1] Author, "Title," Journal, vol. X, no. Y, pp. 1–10, 2026.', {
  className: 'citation',
})
```

### 覆蓋正文字型

保留學術佈局，改用另一種襯線字型：

```ts
const doc = createDocx({
  ...academicPreset.config,
  defaults: {
    ...academicPreset.config.defaults,
    paragraph: {
      ...academicPreset.config.defaults.paragraph,
      fontFamily: 'Computer Modern, Times New Roman, serif',
    },
  },
})
```

## 相關文件

- [樣式預設](/zh-TW/guide/presets) — 預設參考
- [主題](/zh-TW/guide/themes) — 主題令牌
- [基本報告](/zh-TW/examples/basic-report) — 精簡建構器範例
