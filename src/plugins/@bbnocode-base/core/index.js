
/**
 * ENUM 转 像素
 * @param {String|Number} value
 * @returns {Number}
 */
function ENUM_TO_PIXEL(value){
  return Math.ceil(parseFloat(value) / 20 * 1.33445)
}

export {
  ENUM_TO_PIXEL
}
