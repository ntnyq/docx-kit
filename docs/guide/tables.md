# Tables

Create structured data tables with typed columns, custom cell renderers, headers, and styling.

## Basic Table

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

## Typed Tables

Use TypeScript generics for type-safe column keys:

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

## Column Widths

Set column widths as percentages or absolute values:

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

## Custom Cell Renderers

Use `render()` for formatted output, inline styles, or computed values:

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

## Table Styling

### Bordered Tables

```ts
doc.table({
  columns: [{ key: 'name', title: 'Name' }],
  data: [{ name: 'Item' }],
  bordered: true,
})
```

### Striped Rows

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

### Header Cell Styling

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

### Hide Header

```ts
doc.table({
  columns: [{ key: 'label', title: 'Label' }, { key: 'value', title: 'Value' }],
  data: [{ label: 'Total', value: '$1,234' }],
  header: false,  // no header row
  bordered: true,
})
```

## Full Table Example

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

## Table Options Reference

| Option | Type | Default | Description |
|---|---|---|---|
| `columns` | `TableColumn[]` | _(required)_ | Column definitions with key, title, width, align, render |
| `data` | `TData[]` | _(required)_ | Row data matching column keys |
| `bordered` | `boolean` | `false` | Show table borders |
| `striped` | `boolean` | `false` | Alternate row shading |
| `header` | `boolean` | `true` | Show the header row |
| `headerCellStyle` | `DocxStyleRule` | — | Style for header cells |
| `cellStyle` | `DocxStyleRule` | — | Default style for data cells |
