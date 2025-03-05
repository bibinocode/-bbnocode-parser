import { load } from 'cheerio';
import Color from 'color';
import xml2js from 'xml2js';

const RE_LENGTH_UNIT = /^([a-zA-Z]+)$/
const RGB = /([a-fA-F0-9]{2}?){3}?/;

/**
 * pt -> px
 * 1 pt = 1/72 inch
 */
function pt2px(pt) {
  return pt * 96 / 72
}

/**
 * emu -> px
 * 12700 emu = 1 inch
 */
function emu2Px(emu) {
  return pt2px(emu / 12700)
}

/**
 * dxa -> px
 * 1 dxa = 1/20 pt
 */
function dxa2Px(dxa) {
  return pt2px(dxa / 20.0)
}

/**
 * cm -> px
 * 1 cm = 96/2.54 px
 */
function cm2Px(cm) {
  return pt2px(parseInt(cm) * 28.3464567)
}

/**
 * toPx
 */
function toPx(length) {
  const value = parseInt(length)
  const units = String(length).match(RE_LENGTH_UNIT)[1]

  switch (units) {
    case 'cm': return cm2Px(value) // 厘米
    case 'mm': return cm2Px(value / 10) // 毫米
    case 'in': return pt2px(value * 72) // 英寸
    case 'emu': return emu2Px(value) // 缇
    case 'dxa': return dxa2Px(value) // 点
    case 'pt': return pt2px(value) // 磅
    case 'ft': return pt2px(value * 864) // 英尺
    default: return value
  }
}

/**
 * asColor
 * 将颜色转换为十六进制
 * @param {String} v 颜色值
 * @param {ColorTransform} transform 颜色变换
 * 
 * @typedef {Object} ColorTransform
 * @property {Number} lumMod 亮度调节
 * @property {Number} lumOff 亮度偏移
 * @property {Number} tint 透明度
 * @property {Number} shade 阴影
 */
function asColor(v, transform) {
  if (!v || v.length == 0 || v == 'auto') {
    return '#000000'
  }

  v = v.split(' ')[0]
  const rgb = v.charAt(0) == '#' ? v : (RGB.test(v) ? '#' + v : v)
  if (transform) {
    const { lumMod, lumOff, tint, shade } = transform
    if (lumMod || lumOff || tint) {
      let color = Color(rgb)

      if (tint != undefined) {
        color = color.lighten(1 - tint)
      }

      if (lumMod != undefined) {
        color = color.lighten(lumMod)
      }

      if (lumOff != undefined) {
        color = color.darken(lumOff)
      }

      if (shade != undefined) {
        color = color
          .red(color.red() * (1 + shade))
          .green(color.green() * (1 + shade))
          .blue(color.blue() * (1 + shade))
      }

      return `${color.hex()}`.replace(/^0x/, "#")
    }
  }
  return rgb
}



/**
 * 提供从zip文件夹中获取到xml节点
 * @param {JSZip} zip
 * @param {String} file 需要读取的文件名
 */
async function zipToXml(zip, file) {
  try {
    const zipFile = zip.file(file)
    if (!zipFile) {
      console.error('可用文件列表:', Object.keys(zip.files))
      throw new Error(`文件 ${file} 不存在`)
    }

    const xml = await zipFile.async('text')
    const parser = new xml2js.Parser()
    return parser.parseStringPromise(xml)
  } catch (error) {
    throw error
  }
}

/**
 * 通过cheerio解析xml
 */
async function parseXML(zip,file){
  try {
    const zipFile = zip.file(file)
    if(!zipFile){
      console.error('可用文件列表:', Object.keys(zip.files))
      throw new Error(`文件 ${file} 不存在`)
    }
    const xml = await zipFile.async('text')
    return load(xml,{
      xml:true,
      xmlMode:true
    })
  } catch (error) {
    throw error
  }
}

export {
  asColor, cm2Px, dxa2Px, emu2Px, parseXML, pt2px, toPx,
  zipToXml
};

