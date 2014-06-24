'use strict'
var fs = require('fs'),
	util = require('util'),
	book = require('./workbook'),
	comm = require('./common'),
	assert = require('assert');

//open(fileName, function(err, workbook))
exports.open = function(fileName, options, callback){
	var peeksz = 4;
	var buf = new Buffer(4);
	fs.open(fileName, 'r',function(err,fd){
		if(typeof options == 'function'){
			callback = options;
			options = null;
		}
		if(err) {callback(err); return;}
		fs.read(fd,buf,0,peeksz,0,function(err,bytesRead,buffer){
			if(err) {callback(err); return;}
			
			var opts = options || {};
			var bk = null;
			if(buffer[0] == 'P' && buffer[1] == 'K' && buffer[2] == 0x03 && buffer[3] == 0x04){
				//xlsx
				callback(new Error('.xlsx not yet implemented'));
				//cleanUp
				fs.close(fd);
				return;
			}else{
				//xls
				bk = book.openWorkbook(fd,opts);
			}
			var cleanUp = bk.cleanUp;
			var cleanUpCount = 0;
			if(opts.onDemand){
				bk.cleanUp = function(){
					if(cleanUpCount) return;
					cleanUpCount++;
					cleanUp();
					fs.close(fd);
				};
				try{
					callback(null,bk);
				}catch(e){
					bk.cleanUp();
					throw e;
				}
			}else{
				bk.cleanUp = function(){
					if(cleanUpCount) return;
					cleanUpCount++;
					cleanUp();
				};
				fs.close(fd,function(err){ 
					try{
						if(err) {callback(err); return;}
						callback(null,bk);
					}catch(e){
						throw e;
					}finally{
						bk.cleanUp();
					}
				});
			}
		});
		
	});
	
}

// === helper functions

//convert zero-based column index to the corresponding column name
var aToz="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
exports.toColumnName = function(colx){
    assert(colx >= 0,'colx >= 0');
    var name = '';
    while (1){
        var quot = Math.floor(colx/26), rem = colx % 26;
        name = aToz[rem] + name;
        if (!quot)
            return name;
        colx = quot - 1;
	}
}

exports.toBiffVersionString = function(ver){
	return comm.biff_text_from_num[ver];
}