import { ENUM_TO_PIXEL, parsePageSetting } from './core/index.js'

export default {
  name:"@bbnocode/parse-base",
  priority:0,
  process:async(data)=>{
    const pageInfo = parsePageSetting(data.document)
    const ctx = {
      ...data.ctx,
      ENUM_TO_PIXEL
    }
    const content = {
      ...data.content,
      pageInfo
    }
    return {
      document:data.document,
      ctx,
      content
    }
  }
}