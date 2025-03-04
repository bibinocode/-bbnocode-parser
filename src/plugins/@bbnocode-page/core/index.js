/**
 * 解析节属性
 * @param {Object} sectionsXml word/sect1.xml
 */
async function parseSections(ctx, sectionsXml) {
  const wBody = sectionsXml['w:document']['w:body'][0]
  const wSectPr = wBody['w:sectPr'][0]
  const wPgSz = wSectPr['w:pgSz'][0]
  const wPgMar = wSectPr['w:pgMar'][0]
}


export {
  parseSections
}
