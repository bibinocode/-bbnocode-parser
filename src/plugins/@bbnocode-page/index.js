/**
 * @author bbnocode@outlook.com
 * @description 解析docx页面信息
 * @version 1.0.0
 * @since 2025-03-04
 */
import { parseSections } from './core/index.js'

export default {
  name: "@bbnocode/parse-page",
  priority: 1,
  process: async (data) => {
    const { ctx, zip } = data
    const sectionsXml = await ctx.zipToXml(zip, 'word/document.xml')
    await parseSections(data, sectionsXml)
    return data
  }
}


