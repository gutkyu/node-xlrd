"use strict";
var READ_SIZE = 1024 * 128;

//sat : 순차적으로 나열한 sector, [[sector 시작 위치, sector 끝 위치), ...]
//startBySAT : sat를 순차적으로 정렬했을때 복사 시작 위치
function loadBySAT(stream, sat, buffer, bufferOffset, startBySAT, readSize) {
  var satIdx = -1;
  var streamOffset = -1;
  for (var i = 0, sum = 0; i < sat.length; i++) {
    sec = sat[i];
    sum += sec[1] - sec[0];
    if (startBySAT >= sum) {
      streamOffset = sec[0] + (startBySAT - sum);
      break;
    }
    satIdx++;
  }
  if (satIdx < 0 || streamOffset < 0)
    throw new Error("loadBySAT Error : not found");
  var remain = readSize;
  while (remain > 0) {
    readSize = Math.min(sat[satIdx][1] - streamOffset, remain);
    stream.copy(buffer, bufferOffset, streamOffset, readSize);
    remain -= readSize;
    bufferOffset += readSize;
    satIdx++;
    streamOffset = sat[satIdx][0];
  }
}
module.exports = function SATStream(basedStream, sat, basedOffset, length) {
  var self = this;
  var bufSize = READ_SIZE;
  //todo : ArrayBuffer
  var buf = new Buffer(bufSize);
  var bufStart = 0; // cache first index, inclusive
  var bufEnd = 0; // cache last index, exclusive
  if (!length) throw new Error("stream is 0 bytes");
  var streamLength = length;

  function hit(offset, size) {
    return offset >= bufStart && offset + size <= bufEnd;
  }
  function preload(offset) {
    var tryAccumulatedLen = offset + bufSize;
    var readSize =
      tryAccumulatedLen > streamLength ? streamLength - offset : bufSize;
    loadBySAT(basedStream, sat, buf, 0, basedOffset + offset, readSize);
    bufStart = offset;
    bufEnd = bufStart + readSize;
  }
  Object.defineProperty(self, "length", {
    get: function () {
      return streamLength;
    },
  });
  self.read8 = function (buffer, offset) {
    if (!hit(offset, 1)) preload(offset, 1);
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
    return buf.readUInt8(offset - bufStart);
  };
  self.readUInt16LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return buf.readUInt16LE(offset - bufStart);
  };
  self.readUInt32LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return buf.readUInt32LE(offset - bufStart);
  };
  self.readInt16LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return buf.readInt16LE(offset - bufStart);
  };
  self.readInt32LE = function (offset) {
    if (!hit(offset, 4)) preload(offset);
    return buf.readInt32LE(offset - bufStart);
  };
  //end : exclusive position
  self.slice = function (start, end) {
    var readSize = end - start;
    if (!hit(start, readSize)) preload(start);
    return buf.slice(start - bufStart, end - bufStart);
  };
  self.copy = function (buffer, bufferOffset, streamOffset, readSize) {
    if (hit(streamOffset, readSize)) {
      buf.copy(
        buffer,
        bufferOffset,
        streamOffset - bufStart,
        streamOffset - bufStart + readSize
      );
      return;
    }
    loadBySAT(
      basedStream,
      sat,
      buffer,
      bufferOffset,
      basedOffset + streamOffset,
      readSize
    );
  };
};
