import fs from 'fs'
import { parse } from './src/index.js'

// 读取文件
const docxData = fs.readFileSync('./docs/niulan.docx')
const result = await parse(docxData)
console.log("解析结果",result)