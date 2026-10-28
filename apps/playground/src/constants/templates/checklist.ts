import { unindent } from '@ntnyq/utils'

/**
 * Complete checklist example shared by documentation and the playground.
 */
export const CHECKLIST_CODE = unindent(`
  // #region checklist-example
  import { checklistPlugin, createDocx, useTheme } from 'docx-kit'

  const doc = createDocx({
    theme: useTheme('ocean'),
    page: { size: { width: '170mm', height: '205mm' }, margin: '15mm' },
  })
    .use(checklistPlugin())
    // #region checklist-block
    .add({
      type: 'plugin',
      name: 'checklist',
      style: { fontSize: 12 },
      options: {
        title: '版本发布检查清单',
        items: [
          {
            group: '发布前',
            text: '完成代码审核与合并',
            checked: true,
            owner: '林晨',
            note: '关键改动已复核，合并请求检查全部通过。',
          },
          {
            group: '发布前',
            text: '完成核心流程验收',
            owner: '周宁',
            note: '覆盖登录、文档生成和导出，并附上测试报告。',
          },
          {
            group: '发布前',
            text: '确认发布配置、回滚方案与监控告警均已准备就绪',
            owner: 'Alex',
            note: '发布前完成配置核对，并记录本次部署版本。',
          },
          {
            group: '发布后',
            text: '检查生产环境运行状态',
            checked: true,
            owner: '林晨',
            note: '观察错误率与接口响应时间，确认没有新增告警。',
          },
          {
            group: '发布后',
            text: '归档发布记录与验收材料',
            owner: '周宁',
            note: '补充变更说明，保存测试报告与发布截图。',
          },
        ],
      },
    })
    // #endregion checklist-block

  doc.toBlob()
  // #endregion checklist-example
`)

const checklistParts = CHECKLIST_CODE.split(
  /\/\/ #(?:end)?region checklist-block\n?/u,
)
if (checklistParts.length !== 3) {
  throw new Error('The shared checklist example is missing its block region')
}

/**
 * Reuse the same checklist node in the all-plugins document.
 */
export const CHECKLIST_BLOCK_CODE = unindent(checklistParts[1])
