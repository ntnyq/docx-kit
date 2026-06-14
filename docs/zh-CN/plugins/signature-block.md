# Signature Block

生成合同、审批单常见的签署区块。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `parties` | `SignatureParty[]` | 签署方数组 |
| `columns` | `number` | 每行列数，默认 2 |

`SignatureParty` 包含 `label`、`name?`、`date?`。

## 用法

```ts
createDocx().use(signatureBlockPlugin()).plugin('signatureBlock', {
  parties: [{ label: '甲方（盖章）' }, { label: '乙方（盖章）' }],
})
```
