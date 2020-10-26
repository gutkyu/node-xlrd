'use strict';

// in node.js v8 ~ v10, uitl.TextDecoder ,but just can use TextDecoder in v12
var TextDecoder =
  TextDecoder === null || TextDecoder === undefined
    ? require('util').TextDecoder
    : TextDecoder;

//TextDecoder supported encodings
exports.encodingFromCodePage = {
  1200: 'utf16le',
  10000: 'x-mac-roman',
  //  10006: 'macgreek', // guess
  10007: 'x-mac-cyrillic', // guess
  //  10029: 'maccenteuro', // guess
  //  10079: 'maciceland', // guess
  //  10081: 'macturkish', // guess
  32768: 'x-mac-roman',
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

var codecs = {};
codecs['utf16le'] = codecs['utf-16le'] = new TextDecoder('utf-16le', {
  fatal: true,
});

//node.js v8, v12's TextDecoder don't support encoding 'latin1'
//but v15 support
try {
  codecs['latin1'] = new TextDecoder('latin1', {fatal: true});
} catch (error) {
  if (Buffer !== null && Buffer !== undefined)
    codecs['latin1'] = {
      decode: function (data) {
        return Buffer.from(data).toString('latin1');
      },
    };
}
if (!codecs['latin1']) throw new Error('unsupported encoding : "latin1"');

exports.decode = function (data, pos, len, encoding) {
  var buf = pos == 0 && data.length == len ? data : data.slice(pos, pos + len);
  return codecs[encoding].decode(buf);
};
