import fs from 'fs'
import { parse, registerPlugin } from './src/index.js'

registerPlugin({
  name:"@bbnocode/parse-",
  process:async(data)=>{
    data.content.text = "我的信息"
    return data
  }
})

// 读取文件
const docxData = fs.readFileSync('./docs/niulan.docx')

const result = await parse(docxData)
console.log("解析结果",result)