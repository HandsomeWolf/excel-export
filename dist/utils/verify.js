"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.isNumber = exports.isChinese = void 0;
// 定义一个函数来判断字符是否为中文
function isChinese(char) {
    const re = /[^\u4E00-\u9FA5]/;
    return !re.test(char);
}
exports.isChinese = isChinese;
// 定义函数判断是否是数字
function isNumber(char) {
    const re = /\D/;
    return !re.test(char);
}
exports.isNumber = isNumber;
