/**
 * ENUM 转 像素
 * @param {String|Number} value
 * @returns {Number}
 */
function ENUM_TO_PIXEL(value) {
  return Math.ceil(parseInt(value) / 20 * 1.33445)
}


/**
 * 根据宽度高度推断纸张方向
 * @param {Number} width
 * @param {NUmber} height
 */
function inferPaperDirection(width, height) {
  const paperDirection = width > height ? 'horizontal' : 'vertical'
  return paperDirection
}


/**
 * 解析纸张方向
 */
async function parserPaperDirection(docx) {
  const documentXml = await docx.officeDocument.content('w\\:pgSize');
  console.log(documentXml)
}



/**
 * 解析页面设置信息
 */
function parsePageSetting(docx) {
  parserPaperDirection(docx)
  // 获取 pgSz 元素
  const sectPr = docx.officeDocument.content("w\\:sectPr")
  const pgSz = sectPr.children("w\\:pgSz").get(0)
  const pgMar = sectPr.children("w\\:pgMar").get(0)

  let pageWidth = ENUM_TO_PIXEL(pgSz.attribs["w:w"])
  let pageHeight = ENUM_TO_PIXEL(pgSz.attribs["w:h"])
  const pageMarginTop = ENUM_TO_PIXEL(pgMar.attribs["w:top"])
  const pageMarginBottom = ENUM_TO_PIXEL(pgMar.attribs["w:bottom"])
  const pageMarginLeft = ENUM_TO_PIXEL(pgMar.attribs["w:left"])
  const pageMarginRight = ENUM_TO_PIXEL(pgMar.attribs["w:right"])
  const pageHeader = ENUM_TO_PIXEL(pgMar.attribs["w:header"])
  const pageGutter = ENUM_TO_PIXEL(pgMar.attribs["w:gutter"])

  const paperDirection = inferPaperDirection(pageWidth, pageHeight)


  return {
    width: pageWidth,
    height: pageHeight,
    marginTop: pageMarginTop,
    marginBottom: pageMarginBottom,
    marginLeft: pageMarginLeft,
    marginRight: pageMarginRight,
    header: pageHeader,
    gutter: pageGutter,
    paperDirection
  }
}

async function parserDocument(docx) {

}


export {
  ENUM_TO_PIXEL,
  parsePageSetting
};

