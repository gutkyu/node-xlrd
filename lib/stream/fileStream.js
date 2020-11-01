'use strict';
var fs = require('fs');

var READ_SIZE = 1024 * 128;

module.exports = function FileStream(fd, defaultBufferSize) {
  var self = this;
  var bufSize =
    defaultBufferSize && defaultBufferSize > READ_SIZE
      ? defaultBufferSize
      : READ_SIZE;
  //todo : ArrayBuffer
  var buf = new Uint8Array(bufSize);
  var dv = new DataView(buf.buffer);
  var bufStart = 0; // cache first index, inclusive
  var bufEnd = 0; // cache last index, exclusive

  function hit(offset, size) {
    return offset >= bufStart && offset + size <= bufEnd;
  }
  function preload(offset) {
    var tryAccumulatedLen = offset + bufSize;
    var readSize =
      tryAccumulatedLen > streamLength ? streamLength - offset : bufSize;
    fs.readSync(fd, buf, 0, readSize, offset);
    bufStart = offset;
    bufEnd = bufStart + readSize;
  }
  Object.defineProperty(self, 'length', {
    get: function () {
      return streamLength;
    },
  });
  self.read8 = function (buffer, offset) {
    if (!hit(offset, 1)) preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf];
  };
  self.read16 = function (buffer, offset) {
    if (!hit(offset, 2)) preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf];
  };
  self.read32 = function (buffer, offset) {
    if (!hit(offset, 4)) preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf++];
    buffer[2] = buf[offsetBuf++];
    buffer[3] = buf[offsetBuf];
  };
  self.readUInt8 = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return buf[offset - bufStart];
  };
  self.readUInt16LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return dv.getUint16(offset - bufStart, true);
  };
  self.readUInt32LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return dv.getUint32(offset - bufStart, true);
  };
  self.readInt16LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return dv.getInt16(offset - bufStart, true);
  };
  self.readInt32LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return dv.getInt32(offset - bufStart, true);
  };
  //end : exclusive position
  self.slice = function (start, end) {
    var readSize = end - start;
    if (!hit(start, readSize)) preload(start);
    return buf.slice(start - bufStart, end - bufStart);
  };
  self.copy = function (dest, destOffset, streamOffset, readSize) {
    if (hit(streamOffset, readSize)) {
      var destU8 =
        dest instanceof Uint8Array
          ? dest
          : new Uint8Array(dest.buffer, destOffset, readSize);
      var srcU8 = new Uint8Array(buf.buffer, streamOffset - bufStart, readSize);
      destU8.set(srcU8, destOffset);
      return;
    }
    fs.readSync(fd, dest, destOffset, readSize, streamOffset);
  };

  var stats = fs.fstatSync(fd);
  if (!stats.size) throw new Error('File is 0 bytes');
  var streamLength = stats.size;
};
