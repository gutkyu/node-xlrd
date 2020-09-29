'use strict'
var fs = require('fs');
var util = require('util');

var READ_SIZE = 1024 * 128;

//create SimpleBuffer Object which is readonly
exports.create = function(fd, size){
	if(Buffer.isBuffer(fd)){
		return new SimpleBuffer;
	}else{
		if(size === null || size === undefined)  size = 8;
		var fb = new FileBuffer(fd, size);
		return fb;
	}	
}

function SimpleBuffer(buf, size){
  var self = this;
  if(size === null || size === undefined)  size = buf.length;
  var len = 0;
  var _start = 0;
  self.setBound = function(start,end){
    _start = start;
    return len = end - start;
  };
  
  self.ui8= function(position){
    if(position >= len) throw new Error(util.format('SimpleBuffer Error : readUInt8 position %d is out of boundary %d', position, len ));
    return buf.readUInt8(_start + position);
  }
  self.ui16le= function(position){
    if(position >= len +1) throw new Error(util.format('SimpleBuffer Error : readUInt16LE position %d is out of boundary %d', position, len ));
    return buf.readUInt16LE(_start + position);
  }
  self.i16le= function(position){
    if(position >= len +1) throw new Error(util.format('SimpleBuffer Error : readInt16LE position %d is out of boundary %d', position, len ));		
    return buf.readInt16LE(_start + position);
  }
  self.i32le= function(position){
    if(position >= len+3) throw new Error(util.format('SimpleBuffer Error :  readInt32LE position %d is out of boundary %d', position, len ));
    return buf.readInt32LE(_start + position);
  }
  self.toStr = function(encoding, start, end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('SimpleBuffer Error : toString start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('SimpleBuffer Error : toString end %d is out of boundary %d', end, len));		
    
    return buf.toString(encoding,_start+start,_start+end);
  }
  self.toInt32LEs = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('SimpleBuffer Error : toInt32LEs start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('SimpleBuffer Error : toInt32LEs end %d is out of boundary %d', end, len));		
    
    var ret = [];
    for(var i = _start+start; i < _start+end; i += 4) ret.push(buf.readInt32LE(i));
    return ret;
  }
  
  self.toBuf = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('SimpleBuffer Error : toBuf start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('SimpleBuffer Error : toBuf end %d is out of boundary %d', end, len));		
    
    var length = end - start;
    var ret = new Buffer(length);
    buf.copy(ret,0,_start + start, _start + end);
    return  ret;
  }
  self.toSlicedBuffer = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('SimpleBuffer Error : toSlicedBuffer start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('SimpleBuffer Error : toSlicedBuffer end %d is out of boundary %d', end, len));		
    
    return buf.slice(_start + start, _start + end);
  }
}

function FileBuffer(fd, defaultBufferSize){
  var self = this;
  var buf = new Buffer(defaultBufferSize);
  var len = 0;

  //access directly
  var b4 = new Buffer(4);
  //position is a absolute from the beginning of the file 
  self.readUInt8 = function(position){
    fs.readSync(fd,b4,0,1,position);
    return b4.getUInt8(0);
  }
  
  self.readUInt16LE = function(position){
    fs.readSync(fd,b4,0,2,position);
    return b4.readUInt16LE(0);
  }
  
  self.readInt16LE = function(position){
    fs.readSync(fd,b4,0,2,position);
    return b4.getInt16LE(0);
  }
  
  self.readInt32LE = function(position){
    fs.readSync(fd,b4,0,4,position);
    return b4.getInt32LE(0);
  }

  self.slice = function(start, end){
    var length = end - start;
    var buf = new Buffer(length);
    if(length) fs.readSync(fd,buf,0,length,start);
    return buf;
  }

  var fileSize = null;
  Object.defineProperty(self,'length',{get:function(){
    if(fileSize === null || fileSize === undefined){
      var stats = fs.fstatSync(fd);
      if(!stats.size) throw new Error('File size is 0 bytes');
      fileSize = stats.size;
    }
    return fileSize;
  }});
  
  //use buffer
  self.setBound = function(start,end){
    len = end -start;
    if(buf.length < len) buf = Buffer.concat([buf,new Buffer(len-buf.length)] ,len);
    fs.readSync(fd,buf,0,len,start);
    return len;
  }
  
  self.ui8= function(position){
    if(position >= len) throw new Error(util.format('FileBuffer Error : readUInt8 position %d is out of boundary %d', position, len ));
    return buf.readUInt8(position);
  }
  self.ui16le= function(position){
    if(position >= len +1) throw new Error(util.format('FileBuffer Error : readUInt16LE position %d is out of boundary %d', position, len ));
    return buf.readUInt16LE(position);
  }
  self.i16le= function(position){
    if(position >= len+1) throw new Error(util.format('FileBuffer Error : readInt16LE position %d is out of boundary %d', position, len ));		
    return buf.readInt16LE(position);
  }
  self.i32le= function(position){
    if(position >= len+3) throw new Error(util.format('FileBuffer Error : readInt32LE position %d is out of boundary %d', position, len ));
    return buf.readInt32LE(position);
  }
  self.toStr = function(encoding, start, end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('FileBuffer Error : toString start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('FileBuffer Error : toString end %d is out of boundary %d', end, len));		
    
    return buf.toString(encoding,start,end);
  }
  self.toInt32LEs = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('FileBuffer Error : toInt32LEs start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('FileBuffer Error : toInt32LEs end %d is out of boundary %d', end, len));		
    
    var ret = [];
    for(var i = start; i < end; i += 4) ret.push(buf.readInt32LE(i));
    return ret;
  }
  self.toBuf = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('FileBuffer Error : toBuf start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('FileBuffer Error : toBuf end %d is out of boundary %d', end, len));		
    
    var length = end - start;
    var ret = new Buffer(length);
    buf.copy(ret,0,start,end);
    return  ret;
    /*
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,_start+start);
    return ret;
    */
  }
  //start, end : absolute position
  self.dump = function(start,end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = self.length;
    if(start && start >= end) throw new Error(util.format('FileBuffer Error : dump start %d is out of boundary %d', start, len));		
    if(end && end > self.length) throw new Error(util.format('FileBuffer Error : dump end %d is out of boundary %d', end, self.length));		
    
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,start);
    return  ret;
    /*
    var length = end - start;
    var ret = new Buffer(length);
    fs.readSync(fd,ret,0,length,_start+start);
    return ret;
    */
  }
  self.toSlicedBuffer = function(start, end){
    if(start === null || start === undefined) start = 0;
    if(end === null || end === undefined) end = len;
    if(start && start >= len) throw new Error(util.format('FileBuffer Error : toSlicedBuffer start %d is out of boundary %d', start, len));		
    if(end && end > len) throw new Error(util.format('FileBuffer Error : toSlicedBuffer end %d is out of boundary %d', end, len));		
    
    return buf.slice(start,end);
  }
   
}

