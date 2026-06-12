# Property Table

Renders key-value pairs as a styled 2-column table. Ideal for configuration summaries, parameter docs, and spec sheets.

## Import

```ts
import {
  propertyTablePlugin,
  type PropertyTableOptions,
  type PropertyItem,
} from 'docx-kit'
```

## Options

| Option    | Type             | Default      | Description                         |
| --------- | ---------------- | ------------ | ----------------------------------- |
| `items`   | `PropertyItem[]` | _(required)_ | Key-value items to display          |
| `keyBold` | `boolean`        | `true`       | Whether the key column text is bold |
| `striped` | `boolean`        | `true`       | Alternate row background shading    |

### `PropertyItem`

| Field   | Type     | Description                                    |
| ------- | -------- | ---------------------------------------------- |
| `key`   | `string` | Property name (right-aligned, gray background) |
| `value` | `string` | Property value (left-aligned)                  |

## Examples

### Basic Property Table

```ts
import { createDocx, propertyTablePlugin } from 'docx-kit'

const doc = createDocx()
  .use(propertyTablePlugin)
  .h1('Project Overview')
  .plugin('propertyTable', {
    items: [
      { key: '项目名称', value: 'XX Management System' },
      { key: '技术栈', value: 'React + Node.js + PostgreSQL' },
      { key: '团队规模', value: '12 people' },
      { key: '开发周期', value: '6 months' },
      { key: '代码仓库', value: 'github.com/company/project' },
    ],
  })
  .save('project-overview.docx')
```

### Without Bold Keys

```ts
const doc = createDocx()
  .use(propertyTablePlugin)
  .h1('API Endpoint')
  .plugin('propertyTable', {
    items: [
      { key: 'Method', value: 'POST' },
      { key: 'URL', value: '/api/v1/users' },
      { key: 'Content-Type', value: 'application/json' },
      { key: 'Authorization', value: 'Bearer <token>' },
      { key: 'Rate Limit', value: '1000 req/min' },
    ],
    keyBold: false,
  })
  .save('api-doc.docx')
```

### No Striped Rows

```ts
const doc = createDocx()
  .use(propertyTablePlugin)
  .h1('System Info')
  .plugin('propertyTable', {
    items: [
      { key: 'OS', value: 'Ubuntu 24.04 LTS' },
      { key: 'Kernel', value: '6.8.0' },
      { key: 'Node.js', value: 'v22.0.0' },
      { key: 'Memory', value: '32 GB' },
      { key: 'CPU', value: 'AMD EPYC 16 cores' },
    ],
    striped: false,
  })
  .save('system-info.docx')
```

### Configuration Reference

```ts
const doc = createDocx()
  .use(propertyTablePlugin)
  .h1('Environment Variables')

  .h2('Database')
  .plugin('propertyTable', {
    items: [
      { key: 'DB_HOST', value: 'localhost' },
      { key: 'DB_PORT', value: '5432' },
      { key: 'DB_NAME', value: 'production' },
      { key: 'DB_POOL_SIZE', value: '20' },
    ],
  })

  .h2('Redis')
  .plugin('propertyTable', {
    items: [
      { key: 'REDIS_HOST', value: 'redis-cluster.internal' },
      { key: 'REDIS_PORT', value: '6379' },
      { key: 'REDIS_DB', value: '0' },
      { key: 'REDIS_TTL', value: '3600s' },
    ],
  })

  .save('config.docx')
```

### Deployment Spec Sheet

```ts
const doc = createDocx()
  .use(propertyTablePlugin)
  .h1('Deployment Specification')
  .plugin('propertyTable', {
    items: [
      { key: 'Environment', value: 'Production' },
      { key: 'Region', value: 'ap-southeast-1' },
      { key: 'Instance Type', value: 'c7g.xlarge' },
      { key: 'Auto Scaling', value: 'Enabled (2–10 instances)' },
      { key: 'Load Balancer', value: 'ALB with HTTPS' },
      { key: 'Database', value: 'RDS PostgreSQL Multi-AZ' },
      { key: 'Monitoring', value: 'CloudWatch + Grafana' },
      { key: 'Backup', value: 'Daily snapshots, 30-day retention' },
    ],
  })
  .save('deploy-spec.docx')
```
