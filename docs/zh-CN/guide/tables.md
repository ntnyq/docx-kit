# 表格

创建支持类型化列、自定义单元格渲染器、表头与样式的结构化数据表格。

## 基础表格

```ts
import { createDocx } from 'docx-kit'

const doc = createDocx()

doc.table({
  columns: [
    { key: 'name', title: 'Name' },
    { key: 'age', title: 'Age', align: 'center' },
    { key: 'city', title: 'City' },
  ],
  data: [
    { name: 'Alice', age: 28, city: 'New York' },
    { name: 'Bob', age: 35, city: 'London' },
    { name: 'Charlie', age: 42, city: 'Tokyo' },
  ],
})
```

## 类型化表格

使用 TypeScript 泛型确保列键类型安全：

```ts
interface Employee {
  name: string
  department: string
  salary: number
  hireDate: string
}

doc.table<Employee>({
  columns: [
    { key: 'name', title: 'Name' },
    { key: 'department', title: 'Department' },
    { key: 'salary', title: 'Salary', align: 'right' },
    { key: 'hireDate', title: 'Hire Date', align: 'center' },
  ],
  data: [
    { name: 'Alice', department: 'Engineering', salary: 95000, hireDate: '2023-03-15' },
    { name: 'Bob', department: 'Marketing', salary: 82000, hireDate: '2022-11-01' },
  ],
})
```

## 列宽

用百分比或绝对值设置列宽：

```ts
doc.table({
  columns: [
    { key: 'id', title: '#', width: '10%' },
    { key: 'name', title: 'Product Name', width: '40%' },
    { key: 'price', title: 'Price', width: '25%', align: 'right' },
    { key: 'stock', title: 'Stock', width: '25%', align: 'center' },
  ],
  data: [
    { id: 1, name: 'Widget', price: 19.99, stock: 150 },
    { id: 2, name: 'Gadget', price: 49.99, stock: 80 },
  ],
})
```

## 自定义单元格渲染器

通过 `render()` 设置格式化输出、行内样式或计算值：

```ts
interface Sale {
  product: string
  amount: number
  status: 'paid' | 'pending' | 'cancelled'
}

doc.table<Sale>({
  columns: [
    { key: 'product', title: 'Product' },
    {
      key: 'amount',
      title: 'Amount',
      align: 'right',
      render: (val) => `¥${(val as number).toLocaleString()}`,
    },
    {
      key: 'status',
      title: 'Status',
      align: 'center',
      render: (val) => {
        const color = val === 'paid' ? '#22c55e' : val === 'pending' ? '#f59e0b' : '#ef4444'
        return [
          {
            type: 'text',
            text: val as string,
            style: { color, fontWeight: 'bold' },
          },
        ]
      },
    },
  ],
  data: [
    { product: 'Widget', amount: 1500, status: 'paid' },
    { product: 'Gadget', amount: 2300, status: 'pending' },
    { product: 'Doodad', amount: 800, status: 'cancelled' },
  ],
})
```

## 表格样式

### 带边框表格

```ts
doc.table({
  columns: [{ key: 'name', title: 'Name' }],
  data: [{ name: 'Item' }],
  bordered: true,
})
```

### 斑马纹行

```ts
doc.table({
  columns: [{ key: 'name', title: 'Name' }, { key: 'value', title: 'Value' }],
  data: [
    { name: 'A', value: 1 },
    { name: 'B', value: 2 },
    { name: 'C', value: 3 },
  ],
  striped: true,
})
```

### 表头单元格样式

```ts
doc.table({
  columns: [
    { key: 'name', title: 'Name' },
    { key: 'value', title: 'Value', align: 'right' },
  ],
  data: data,
  headerCellStyle: {
    fontWeight: 'bold',
    backgroundColor: '#1e293b',
    color: '#ffffff',
    fontSize: 11,
  },
  cellStyle: {
    fontSize: 10,
    verticalAlign: 'middle',
  },
})
```

### 隐藏表头

```ts
doc.table({
  columns: [{ key: 'label', title: 'Label' }, { key: 'value', title: 'Value' }],
  data: [{ label: 'Total', value: '$1,234' }],
  header: false,  // no header row
  bordered: true,
})
```

## 完整表格示例

```ts
import { createDocx, defineStyles } from 'docx-kit/node'

const styles = defineStyles({
  tableHeader: {
    fontWeight: 'bold',
    backgroundColor: '#2563eb',
    color: '#ffffff',
    fontSize: 10,
  },
  tableCell: {
    fontSize: 10,
  },
})

interface FinancialRow {
  quarter: string
  revenue: number
  costs: number
  profit: number
  margin: string
}

const doc = createDocx({ styles })

doc
  .h1('Financial Summary')
  .table<FinancialRow>({
    columns: [
      { key: 'quarter', title: 'Quarter' },
      {
        key: 'revenue',
        title: 'Revenue',
        align: 'right',
        render: (v) => `¥${(v as number).toLocaleString()}`,
      },
      {
        key: 'costs',
        title: 'Costs',
        align: 'right',
        render: (v) => `¥${(v as number).toLocaleString()}`,
      },
      {
        key: 'profit',
        title: 'Profit',
        align: 'right',
        render: (v) => {
          const n = v as number
          return [
            {
              type: 'text',
              text: `¥${n.toLocaleString()}`,
              style: { color: n >= 0 ? '#22c55e' : '#ef4444', fontWeight: 'bold' },
            },
          ]
        },
      },
      { key: 'margin', title: 'Margin', align: 'center' },
    ],
    data: [
      { quarter: 'Q1', revenue: 500000, costs: 350000, profit: 150000, margin: '30%' },
      { quarter: 'Q2', revenue: 620000, costs: 410000, profit: 210000, margin: '34%' },
      { quarter: 'Q3', revenue: 580000, costs: 430000, profit: 150000, margin: '26%' },
      { quarter: 'Q4', revenue: 750000, costs: 480000, profit: 270000, margin: '36%' },
    ],
    bordered: true,
    striped: true,
    headerCellStyle: {
      fontWeight: 'bold',
      backgroundColor: '#1e293b',
      color: '#ffffff',
    },
    cellStyle: {
      fontSize: 10,
      verticalAlign: 'middle',
    },
  })
  .save('financials.docx')
```

## 表格选项参考

| 选项 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `columns` | `TableColumn[]` | _必填_ | 列定义：key、title、width、align、render |
| `data` | `TData[]` | _必填_ | 与列键匹配的行数据 |
| `bordered` | `boolean` | `false` | 显示表格边框 |
| `striped` | `boolean` | `false` | 交替行底纹 |
| `header` | `boolean` | `true` | 显示表头行 |
| `headerCellStyle` | `DocxStyleRule` | — | 表头单元格样式 |
| `cellStyle` | `DocxStyleRule` | — | 数据单元格默认样式 |
