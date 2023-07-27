// 定义一个函数来判断字符是否为中文
export function isChinese(char: string) {
  const re = /[^\u4E00-\u9FA5]/;
  return !re.test(char);
}
// 定义函数判断是否是数字
export function isNumber(char: string) {
  const re = /\D/;
  return !re.test(char);
}
