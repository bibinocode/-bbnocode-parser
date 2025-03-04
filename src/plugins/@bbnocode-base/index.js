import { ENUM_TO_PIXEL, parsePageSetting } from './core/index.js'

export default {
  name: "@bbnocode/parse-base",
  priority: 0,
  process: async (data) => {
    const ctx = {
      ...data.ctx,
      ENUM_TO_PIXEL
    }

    const pageInfo = parsePageSetting(data.document)

    // 文档主体、页眉、页脚的内容元素。
    const documentElements = []
    const headerElements = []
    const footerElements = []
    const bodyElements = []
    const lisMap = []

    const content = {
      ...data.content,
      page: pageInfo
    }
    return {
      document: data.document,
      ctx,
      content
    }
  }
}