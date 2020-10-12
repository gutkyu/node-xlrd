'use strict'
var fs = require('fs');
var util = require('util');

//basedPostion : default based stream position
module.exports = function StreamReader(stream, endian){
  var self = this;
  var base = 0;
  var len = stream.length;
  var cur = base;

  if(endian === null || endian === undefined){
    endian = 'little';
  }
  var readUInt16, readInt16, readUInt32, readInt32;
  if(endian == 'little'){
    readInt16 = stream.readInt16LE;
    readUInt16 = stream.readUInt16LE;
    readInt32 = stream.readInt32LE;
    readUInt32 = stream.readUInt32LE;
  }else{
    throw new Error('StreamReader Error : big endian not yet');
  }
  self.mov = function(offset){
    var mCur = base + offset;
    if(mCur >= len) throw new Error(util.format('StreamReader Error : offset %d is out of boundary %d', mCur, len ));
    cur = mCur;
  }
  //consume bytes
  self.null = function(count){
    var mCur = cur + count;
    if(mCur >= len) throw new Error(util.format('StreamReader Error : offset %d is out of boundary %d', mCur, len ));
    cur = mCur;
  }
  Object.defineProperty(self,'cur',{get:function(){
    return cur;
  }});
  Object.defineProperty(self,'u8',{get:function(){
    if(cur >= len) throw new Error(util.format('StreamReader Error : u8 current position %d is out of boundary %d', cur, len ));
    return stream.readUInt8(cur++);
  }});
  Object.defineProperty(self,'u16',{get:function(){
    if(cur >= len - 1) throw new Error(util.format('StreamReader Error : u16 current position %d is out of boundary %d', cur, len ));
    var ret = readUInt16(cur);
    cur += 2;
    return ret;
  }});
  Object.defineProperty(self,'i16',{get:function(){
    if(cur >= len - 1) throw new Error(util.format('StreamReader Error : i16 position %d is out of boundary %d', cur, len ));		
    var ret = readInt16(cur);
    cur += 2;
    return ret;
  }});
  Object.defineProperty(self,'u32',{get:function(){
    if(cur >= len - 3) throw new Error(util.format('StreamReader Error : u32 position %d is out of boundary %d', cur, len ));
    var ret = readUInt32(cur);
    cur += 4;
    return ret;
  }});
  Object.defineProperty(self,'i32',{get:function(){
    if(cur >= len - 3) throw new Error(util.format('StreamReader Error : i32 position %d is out of boundary %d', cur, len ));
    var ret = readInt32(cur);
    cur += 4;
    return ret;
  }});
  self.str = function(encoding, readSize){
    if(readSize === null || readSize === undefined) readSize = len - cur;
    if(cur >= len - readSize + 1) throw new Error(util.format('StreamReader Error : i32 position %d is out of boundary %d', cur, len )); 
    var ret = stream.slice(cur,cur +readSize).toString(encoding);
    cur += readSize;
    return ret;
  }
  //readSize : byte count 
  self.i32s = function(readSize){
    if(readSize) {
      if((readSize & 0x03))
        throw new Error(util.format('StreamReader Error : i32s readSize %d is not a multiple of 4', readSize ));
      if( cur > len - readSize )
        throw new Error(util.format('StreamReader Error : i32s position %d is out of boundary %d', cur, len ));
    }
    var ret = [];
    var i = cur;
    var bEnd = (readSize === null || readSize === undefined) ? len : cur + readSize;
    for(; i < bEnd; i += 4) ret.push(readInt32(i));
    cur = i;
    return ret;
  }
  self.buf = function(readSize){
    if(readSize === null || readSize === undefined) readSize = len - cur;
    if(cur >= len - readSize + 1) throw new Error(util.format('StreamReader Error : i32 position %d is out of boundary %d', cur, len )); 
    
    var ret = new Buffer(readSize);
    stream.copy(ret,0,cur,readSize);
    cur += readSize;
    return ret;
    /*
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,_start+start);
    return ret;
    */
  }
  //start, end :  relative to base position
  self.dump = function(start,end){
    if(start === null || start === undefined) start = base;
    if(end === null || end === undefined) end = len;
    if(start && start >= end) throw new Error(util.format('StreamReader Error : dump start %d is greater than end %d', start, end));		
    if(end && end > self.length) throw new Error(util.format('StreamReader Error : dump end %d is out of boundary %d', end, len));		
    
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,base+start);
    return  ret;
    /*
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,_start+start);
    return ret;
    */
  }
}
