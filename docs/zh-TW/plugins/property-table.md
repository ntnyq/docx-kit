# 屬性表（Property Table）

將鍵值對呈現為帶樣式的兩欄表格，適合設定摘要、參數文件與規格表。

## 匯入

```ts
import {
  propertyTablePlugin,
  type PropertyTableOptions,
  type PropertyItem,
} from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --------- | ---------------- | ------------ | ----------------------------------- |
| `items` | `PropertyItem[]` | _必填_ | 要顯示的鍵值條目 |
| `keyBold` | `boolean` | `true` | 是否加粗鍵列文字 |
| `striped` | `boolean` | `true` | 交替行背景底紋 |

### `PropertyItem`

| 欄位 | 型別 | 說明 |
| ------- | -------- | ---------------------------------------------- |
| `key` | `string` | 屬性名稱，右對齊、灰色背景 |
| `value` | `string` | 屬性值，左對齊 |

## 範例

### 基本屬性表

```ts
import { createDocx, propertyTablePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(propertyTablePlugin())
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

### 不加粗鍵名

```ts
const doc = createDocx()
  .use(propertyTablePlugin())
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

### 不使用斑馬紋

```ts
const doc = createDocx()
  .use(propertyTablePlugin())
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

### 設定參考

```ts
const doc = createDocx()
  .use(propertyTablePlugin())
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

### 部署規格表

```ts
const doc = createDocx()
  .use(propertyTablePlugin())
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
