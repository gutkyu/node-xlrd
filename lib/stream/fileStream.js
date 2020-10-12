'use strict'
var fs = require('fs');

var READ_SIZE = 1024 * 128;

module.exports = function FileStream(fd, defaultBufferSize){
  var self = this;
  var bufSize = defaultBufferSize && defaultBufferSize > READ_SIZE ? defaultBufferSize : READ_SIZE; 
  //todo : ArrayBuffer
  var buf = new Buffer(bufSize);
  var bufStart = 0; // cache first index, inclusive
  var bufEnd = 0; // cache last index, exclusive

  var stats = fs.fstatSync(fd);
  if(!stats.size) throw new Error('File is 0 bytes');
  var streamLength = stats.size;

  function hit(offset, size){
    return offset >= bufStart && (offset+size <= bufEnd);
  }
  function preload(offset){
    var tryAccumulatedLen = offset + bufSize;
    var readSize =tryAccumulatedLen > streamLength?streamLength-offset:bufSize;
    fs.readSync(fd,buf,0, readSize, offset);
    bufStart = offset;
    bufEnd = bufStart + readSize;
  }
  Object.defineProperty(self,'length', {get:function(){
    return streamLength;
  }});
  self.read8 = function(buffer, offset){
    if(!hit(offset,1)) 
      preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf];
  }
  self.read16 = function(buffer, offset){
    if(!hit(offset,2))
      preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf];
  }
  self.read32 = function(buffer, offset){
    if(!hit(offset,4))
      preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf++];
    buffer[2] = buf[offsetBuf++];
    buffer[3] = buf[offsetBuf];
  }
  self.readUInt8 = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt8(offset - bufStart)
  }
  self.readUInt16LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt16LE(offset - bufStart);
  }
  self.readUInt32LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt32LE(offset - bufStart);
  }
  self.readInt16LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readInt16LE(offset - bufStart);
  }
  self.readInt32LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readInt32LE(offset - bufStart);
  }
  //end : exclusive position
  self.slice = function(start, end){
    var readSize = end - start;
    if(!hit(start,readSize))
      preload(start);
    return buf.slice(start - bufStart, end - bufStart);
  }
  self.copy = function(buffer, bufferOffset, streamOffset, readSize){
    if(hit(streamOffset,readSize)){
      buf.copy(buffer, bufferOffset, streamOffset - bufStart, streamOffset - bufStart + readSize);
      return;
    }
    fs.readSync(fd,buffer,bufferOffset, readSize, streamOffset);
  }
}
//bytes : UInt8Array
function ByteStream(bytes){
  
}

function SATStream(basedStream, sat, basedOffset, length){
  var self = this;
  var bufSize = READ_SIZE; 
  //todo : ArrayBuffer
  var buf = new Buffer(bufSize);
  var bufStart = 0; // cache first index, inclusive
  var bufEnd = 0; // cache last index, exclusive
  if(!length) throw new Error('stream is 0 bytes');
  var streamLength = length;

  function hit(offset, size){
    return offset >= bufStart && offset+size <= bufEnd;
  }
  function preload(offset){
    var tryAccumulatedLen = offset + bufSize;
    var readSize =tryAccumulatedLen > streamLength?streamLength-offset:bufSize;
    basedStream.copy(buf, 0, basedOffset + offset, readSize);
    bufStart = offset;
    bufEnd = bufStart + readSize;
  }
  Object.defineProperty(self,'length', {get:function(){
    return streamLength;
  }});
  self.read8 = function(buffer, offset){
    if(!hit(offset,1)) 
      preload(offset,1);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf];
  }
  self.read16 = function(buffer, offset){
    if(!hit(offset,2))
      preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf];
  }
  self.read32 = function(buffer, offset){
    if(!hit(offset,4))
      preload(offset);
    var offsetBuf = offset - bufStart;
    buffer[0] = buf[offsetBuf++];
    buffer[1] = buf[offsetBuf++];
    buffer[2] = buf[offsetBuf++];
    buffer[3] = buf[offsetBuf];
  }
  self.readUInt8 = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt8(offset - bufStart)
  }
  self.readUInt16LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt16LE(offset - bufStart);
  }
  self.readUInt32LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readUInt32LE(offset - bufStart);
  }
  self.readInt16LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readInt16LE(offset - bufStart);
  }
  self.readInt32LE = function(offset){
    if(!hit(offset,4))
      preload(offset);
    return buf.readInt32LE(offset - bufStart);
  }
  //end : exclusive position
  self.slice = function(start, end){
    var readSize = end - start;
    if(!hit(start,readSize))
      preload(start);
    return buf.slice(start - bufStart, end - bufStart);
  }
  self.copy = function(buffer, bufferOffset, streamOffset, readSize){
    if(hit(streamOffset,readSize)){
      buf.copy(buffer, bufferOffset, streamOffset - bufStart, streamOffset - bufStart + readSize);
      return;
    }
    basedStream.copy(buffer, bufferOffset, basedOffset + streamOffset, readSize);
  }
}