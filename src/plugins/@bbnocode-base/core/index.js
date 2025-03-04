/**
 * ENUM 转 像素
 * @param {String|Number} value
 * @returns {Number}
 */
function ENUM_TO_PIXEL(value){
  return Math.ceil(parseInt(value) / 20  * 1.33445)
}


/**
 * 解析页面设置信息
 */
function parsePageSetting(docx){
  // 获取 pgSz 元素
  const sectPr = docx.officeDocument.content("w\\:sectPr")
  const pgSz = sectPr.children("w\\:pgSz").get(0)
  const pgMar = sectPr.children("w\\:pgMar").get(0)
  
  const pageWidth = ENUM_TO_PIXEL(pgSz.attribs["w:w"])
  const pageHeight = ENUM_TO_PIXEL(pgSz.attribs["w:h"])
  const pageMarginTop = ENUM_TO_PIXEL(pgMar.attribs["w:top"])
  const pageMarginBottom = ENUM_TO_PIXEL(pgMar.attribs["w:bottom"])
  const pageMarginLeft = ENUM_TO_PIXEL(pgMar.attribs["w:left"])
  const pageMarginRight = ENUM_TO_PIXEL(pgMar.attribs["w:right"])
  const pageHeader = ENUM_TO_PIXEL(pgMar.attribs["w:header"])
  const pageGutter = ENUM_TO_PIXEL(pgMar.attribs["w:gutter"])

  return {
    width: pageWidth,
    height: pageHeight,
    marginTop: pageMarginTop,
    marginBottom: pageMarginBottom,
    marginLeft: pageMarginLeft,
    marginRight: pageMarginRight,
    header: pageHeader,
    gutter: pageGutter
  }
}


export {
  ENUM_TO_PIXEL,
  parsePageSetting
}

