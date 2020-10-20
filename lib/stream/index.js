'use strict';

var SATStream = require('./satStream');
var FileStream = require('./fileStream');
var ByteStream = require('./byteStream');
var StreamReader = require('./streamReader');

exports.create = function (source, param1, param2, param3) {
  if (
    source instanceof SATStream ||
    source instanceof ByteStream ||
    source instanceof FileStream
  ) {
    return new SATStream(source, param1, param2, param3); //basedStream, sat, basedOffset, length
  } else if (source && source.buffer && source.buffer instanceof ArrayBuffer) {
    return new ByteStream(source); //bytes
  } else if (typeof source == 'number') {
    return new FileStream(source, param1); //fd, size
  } else {
    throw new Error('Error : source is unknown Stream type');
  }
};

exports.StreamReader = StreamReader;
