'use strict';
var iconv = require('iconv-lite');

//iconv-lite supported encodings
// -  https://github.com/ashtuchkin/iconv-lite/wiki/Supported-Encodings
exports.encodingFromCodePage = {
  1200: 'utf16le',
  10000: 'macroman',
  10006: 'macgreek', // guess
  10007: 'maccyrillic', // guess
  10029: 'maccenteuro', // guess
  10079: 'maciceland', // guess
  10081: 'macturkish', // guess
  32768: 'macroman',
  32769: 'cp1252',
};

/*
// some more guessing, for Indic scripts
// codePage 57000 range:
// 2 Devanagari [0]
// 3 Bengali [1]
// 4 Tamil [5]
// 5 Telegu [6]
// 6 Assamese [1] c.f. Bengali
// 7 Oriya [4]
// 8 Kannada [7]
// 9 Malayalam [8]
// 10 Gujarati [3]
// 11 Gurmukhi [2]
*/

exports.decode = function (data, pos, len, encoding) {
  var buf = pos == 0 && data.length == len ? data : data.slice(pos, pos + len);
  var ret = iconv.decode(buf, encoding);
  return ret;
};

exports.isEncoding = function (encoding) {
  return iconv.encodingExists(encoding);
};
