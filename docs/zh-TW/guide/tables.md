# 表格

建立支援型別化欄、自訂儲存格呈現器、表頭與樣式的結構化資料表格。

## 基本表格

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

## 型別化表格

使用 TypeScript 泛型確保欄鍵型別安全：

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

## 欄寬

用百分比或絕對值設置欄寬：

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

## 自訂儲存格呈現器

透過 `render()` 設置格式化輸出、行內樣式或計算值：

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

## 表格樣式

### 帶邊框表格

```ts
doc.table({
  columns: [{ key: 'name', title: 'Name' }],
  data: [{ name: 'Item' }],
  bordered: true,
})
```

### 斑馬紋行

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

### 表頭儲存格樣式

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

### 隱藏表頭

```ts
doc.table({
  columns: [{ key: 'label', title: 'Label' }, { key: 'value', title: 'Value' }],
  data: [{ label: 'Total', value: '$1,234' }],
  header: false,  // no header row
  bordered: true,
})
```

## 完整表格範例

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

## 表格選項參考

| 選項 | 型別 | 預設值 | 說明 |
|---|---|---|---|
| `columns` | `TableColumn[]` | _必填_ | 欄定義：key、title、width、align、render |
| `data` | `TData[]` | _必填_ | 與欄鍵匹配的行資料 |
| `bordered` | `boolean` | `false` | 顯示錶格邊框 |
| `striped` | `boolean` | `false` | 交替行底紋 |
| `header` | `boolean` | `true` | 顯示錶頭行 |
| `headerCellStyle` | `DocxStyleRule` | — | 表頭儲存格樣式 |
| `cellStyle` | `DocxStyleRule` | — | 資料儲存格預設樣式 |
