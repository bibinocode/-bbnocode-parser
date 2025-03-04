/**
 * @author bbnocode@outlook.com
 * @description 插件对外暴露一些方法
 * @version 1.0.0
 * @since 2025-03-04
 */
import * as ctx from './core/index.js'

export default {
  name: "@bbnocode/parse-base",
  priority: 0,
  process: async (data) => {
    return {
      ...data,
      ctx: {
        ...data.ctx,
        ...ctx
      }
    }
  }
}