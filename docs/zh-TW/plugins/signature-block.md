# Signature Block

生成合約、簽核單常見的簽署區塊。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `parties` | `SignatureParty[]` | 簽署方陣列 |
| `columns` | `number` | 每列欄數，預設 2 |

`SignatureParty` 包含 `label`、`name?`、`date?`。

## 用法

```ts
createDocx().use(signatureBlockPlugin()).plugin('signatureBlock', {
  parties: [{ label: '甲方（蓋章）' }, { label: '乙方（蓋章）' }],
})
```
