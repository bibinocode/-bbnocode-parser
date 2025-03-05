/**
 * @author bbnocode@outlook.com
 * @description 解析docx页面信息
 * @version 1.0.0
 * @since 2025-03-04
 */
import { fontMapper, parseHeader, parseSectPr, parserFooter } from './core/index.js'

export default {
  name: "@bbnocode/parse-page",
  priority: 1,
  process: async (data) => {
    const { ctx, zip } = data
    const $ = await ctx.parseXML(zip, 'word/document.xml')
    const page = await parseSectPr(ctx, $)

    // 走页眉
    if(page.header_rules.length > 0 ){
      await parseHeader(zip,ctx,page.header_rules)
    }

    // 走页脚
    if(page.footer_rules.length > 0){
      await parserFooter(zip,ctx,page.footer_rules)
    }


    return {
      ...data,
      ctx:{
        ...data.ctx,
      },
      page:{
        ...page,
        fontMapper
      }
    }
  }
}


