/**
 * 字体映射
 */
const fontMapper = {
  "LiSu":"隶书",
  "SimSun":"宋体",
  "Microsoft Yahei":"微软雅黑",
  "SimHei":"黑体",
  "KaiTi":"楷体",
  "NSimSun":"新宋体",
  "STXingkai":"华文行楷",
  "STFangsong":"华文仿宋",
  "FangSong":"仿宋" ,
  "YouYuan":"幼圆",
  "STSong":"华文宋体",
  "STZhongsong":"华文中宋",
  "SimSun":"等线",
  "SimSun":"等线 Light",
  "STHupo":"华文琥珀",
  "STLiti":"华文隶书",
  "STXinwei":"华文新魏",
  "STCaiyun":"华文彩云",
  "FZYaoti":"方正姚体",
  "FZShuTi":"方正舒体",
  "STXihei":"华文细黑",
  "simsun-extB":"宋体扩展",
  "FangSong_GB2312":"仿宋_GB2312",
  "SimSun":"新細明體",
}





/**
 * 解析 w:sectP 中的信息
 * @param {Object} sectionsXml word/sect1.xml
 */
async function parseSectPr(ctx, $) {
  const pgSz = $('w\\:pgSz')[0].attribs
  const pgMar = $('w\\:pgMar')[0].attribs
  // 尝试查找页眉引用
  const headerReference = $('w\\:headerReference') || []
  const header_rules = []

  // 查找页脚引用
  const footerReference = $('w\\:footerReference') || []
  const footer_rules = []

  // 如果header存在,则去查找rule对应的id
  if(headerReference.length > 0){
    for(const headerRef of headerReference){
      const ruleId = headerRef.attribs['r:id']
      header_rules.push(ruleId)
    }
  }

  if(footerReference.length > 0){
    for(const footerRef of footerReference){
      const ruleId = footerRef.attribs['r:id']
      footer_rules.push(ruleId)
    }
  }

  const width = Math.ceil(ctx.dxa2Px(parseInt(pgSz['w:w'])))
  const height = Math.ceil(ctx.dxa2Px(parseInt(pgSz['w:h'])))
  const paperDirection = width > height ? 'horizontal' : 'vertical'
  return {
    size: {
      width: paperDirection === 'vertical' ? width : height,
      height: paperDirection === 'vertical' ? height : width
    },
    margin: {
      top: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:top']))),
      right: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:right']))),
      bottom: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:bottom']))),
      left: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:left']))),
      header: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:header']))),
      footer: Math.ceil(ctx.dxa2Px(parseInt(pgMar['w:footer'])))
    },
    paperDirection,
    header_rules,
    footer_rules
  }
}



/**
 * 从rules中用r:id提取对应的文件索引
 */
async function getRulesFile (zip,ctx,Rid){
  const $ = await ctx.parseXML(zip, 'word/_rels/document.xml.rels')
  // 根据Id="rId12" 提取 Target
  const target = $('Relationships').find(`Relationship[Id="${Rid}"]`).attr('Target')
  // 当前索引位置是 word/ 如果target是 ../ 开头则需要处理路径
  return target.startsWith('../') ? `${zip.name.split('/').slice(0,-1).join('/')}/${target.slice(2)}` : `word/${target}`
}



/**
 * 解析页眉
 */
async function parseHeader(zip, ctx, rules) {
  const headers = []
  for (const rule of rules) {
    const file = await getRulesFile(zip, ctx, rule)
    const $ = await ctx.parseXML(zip, file)
    
    const headerContent = []
    $('w\\:p').each((_, paragraph) => {
      // 过滤掉所有在 mc:AlternateContent 内的内容
      const $paragraph = $(paragraph)
      if ($paragraph.find('mc\\:AlternateContent').length > 0) {
        const textOutsideAlt = $paragraph.clone()
          .find('mc\\:AlternateContent')
          .remove()
          .end()
          .find('w\\:t')
          .text()
          .trim()

        if (textOutsideAlt) {
          const rPr = $paragraph.find('w\\:rPr').first()
          const style = {
            bold: rPr.find('w\\:b').length > 0,
            italic: rPr.find('w\\:i').length > 0,
            underline: rPr.find('w\\:u').length > 0,
            strikeout: rPr.find('w\\:strike').length > 0,
            size: parseInt(rPr.find('w\\:sz').attr('w:val') || '24') / 2,
            font: '微软雅黑',
            rowFlex: 'left',
            dashArray: []
          }

          const rFonts = rPr.find('w\\:rFonts')
          if (rFonts.length) {
            const fontName = rFonts.attr('w:eastAsiaTheme') || rFonts.attr('w:ascii')
            if (fontName) {
              style.font = fontName
            }
          }

          headerContent.push({
            ...style,
            value: textOutsideAlt
          })
        }
      }
    })

    headers.push(headerContent)
  }
  
  return headers
}


/**
 * 解析页脚
 */
async function parserFooter(zip, ctx, rules) {
  const footers = []
  for (const rule of rules) {
    const file = await getRulesFile(zip, ctx, rule)
    const $ = await ctx.parseXML(zip, file)
    
    const footerContent = []
    
    $('w\\:p').each((_, paragraph) => {
      // 过滤掉所有在 mc:AlternateContent 内的内容
      const $paragraph = $(paragraph)
      if ($paragraph.find('mc\\:AlternateContent').length > 0) {
        const textOutsideAlt = $paragraph.clone()
          .find('mc\\:AlternateContent')
          .remove()
          .end()
          .find('w\\:t')
          .text()
          .trim()

        if (textOutsideAlt) {
          const rPr = $paragraph.find('w\\:rPr').first()
          const style = {
            bold: rPr.find('w\\:b').length > 0,
            italic: rPr.find('w\\:i').length > 0,
            underline: rPr.find('w\\:u').length > 0,
            strikeout: rPr.find('w\\:strike').length > 0,
            size: parseInt(rPr.find('w\\:sz').attr('w:val') || '24') / 2,
            font: '微软雅黑',
            rowFlex: 'left',
            dashArray: []
          }

          const rFonts = rPr.find('w\\:rFonts')
          if (rFonts.length) {
            const fontName = rFonts.attr('w:eastAsiaTheme') || rFonts.attr('w:ascii')
            if (fontName) {
              style.font = fontName
            }
          }

          footerContent.push({
            ...style,
            value: textOutsideAlt
          })
        }
      } else {
        // 处理普通文本
        $paragraph.find('w\\:r').each((_, run) => {
          const rPr = $(run).find('w\\:rPr')
          const text = $(run).find('w\\:t').text()
          
          if (text) {
            const style = {
              bold: rPr.find('w\\:b').length > 0,
              italic: rPr.find('w\\:i').length > 0,
              underline: rPr.find('w\\:u').length > 0,
              strikeout: rPr.find('w\\:strike').length > 0,
              size: parseInt(rPr.find('w\\:sz').attr('w:val') || '24') / 2,
              font: '微软雅黑',
              rowFlex: 'left',
              dashArray: []
            }

            const rFonts = rPr.find('w\\:rFonts')
            if (rFonts.length) {
              const fontName = rFonts.attr('w:eastAsiaTheme') || rFonts.attr('w:ascii')
              if (fontName) {
                style.font = fontName
              }
            }

            footerContent.push({
              ...style,
              value: text
            })
          }
        })
      }
    })

    footers.push(footerContent)
  }
  
  return footers
}



export {
  fontMapper, parseHeader, parserFooter, parseSectPr
}




