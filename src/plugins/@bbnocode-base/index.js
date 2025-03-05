/**
 * @author bbnocode@outlook.com
 * @description 插件对外暴露一些方法
 * @version 1.0.0
 * @since 2025-03-04
 */
import {
  asColor, cm2Px, dxa2Px, emu2Px,
  parseXML,
  pt2px, toPx,
  zipToXml
} from './core/index.js'

export default {
  name: "@bbnocode/parse-base",
  priority: 0,
  process: async (data) => {
    console.log("base-data")
    return {
      ...data,
      ctx: {
        ...data.ctx,
        asColor, cm2Px, dxa2Px, emu2Px, pt2px, toPx,
        zipToXml,
        parseXML
      }
    }
  }
}