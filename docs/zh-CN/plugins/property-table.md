# 属性表（Property Table）

将键值对渲染为带样式的两列表格，适合配置摘要、参数文档与规格表。

## 导入

```ts
import {
  propertyTablePlugin,
  type PropertyTableOptions,
  type PropertyItem,
} from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --------- | ---------------- | ------------ | ----------------------------------- |
| `items` | `PropertyItem[]` | _必填_ | 要显示的键值条目 |
| `keyBold` | `boolean` | `true` | 是否加粗键列文字 |
| `striped` | `boolean` | `true` | 交替行背景底纹 |

### `PropertyItem`

| 字段 | 类型 | 说明 |
| ------- | -------- | ---------------------------------------------- |
| `key` | `string` | 属性名称，右对齐、灰色背景 |
| `value` | `string` | 属性值，左对齐 |

## 示例

### 基础属性表

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

### 不加粗键名

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

### 不使用斑马纹

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

### 配置参考

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

### 部署规格表

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
