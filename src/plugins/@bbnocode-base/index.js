import { ENUM_TO_PIXEL } from './core/index.js'

export default {
  name:"@bbnocode/parse-base",
  priority:0,
  process:async(data)=>{
    console.log("加载基础插件")
    const ctx = {
      ...data.ctx,
      ENUM_TO_PIXEL
    }
    return {
      ...data,
      ctx
    }
  }
}