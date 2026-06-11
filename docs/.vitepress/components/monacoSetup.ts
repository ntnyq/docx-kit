/**
 * Monaco Editor setup module.
 * This file MUST only be imported dynamically in `onMounted` to
 * avoid SSR issues — Monaco references `window` and imports CSS.
 */
import * as monaco from 'monaco-editor'
import editorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker'
import cssWorker from 'monaco-editor/esm/vs/language/css/css.worker?worker'
import htmlWorker from 'monaco-editor/esm/vs/language/html/html.worker?worker'
import jsonWorker from 'monaco-editor/esm/vs/language/json/json.worker?worker'
import tsWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker'

const globalSelf: any = self

globalSelf.MonacoEnvironment = {
  getWorker(_: any, label: string) {
    if (label === 'typescript' || label === 'javascript') {
      // eslint-disable-next-line new-cap
      return new tsWorker()
    }
    if (label === 'json') {
      // eslint-disable-next-line new-cap
      return new jsonWorker()
    }
    if (label === 'css' || label === 'scss' || label === 'less') {
      // eslint-disable-next-line new-cap
      return new cssWorker()
    }
    if (label === 'html' || label === 'handlebars' || label === 'razor') {
      // eslint-disable-next-line new-cap
      return new htmlWorker()
    }
    // eslint-disable-next-line new-cap
    return new editorWorker()
  },
}

export { monaco }
