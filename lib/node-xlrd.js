'use strict'
var fs = require('fs'),
	util = require('util'),
	book = require('./workbook'),
	comm = require('./common'),
	assert = require('assert');

//open(fileName, function(err, workbook))
exports.open = function(fileName, options, callback){
	var peekSz = 4;
	var buf = new Buffer(4);
	fs.open(fileName, 'r',function(err,fd){
		if(typeof options == 'function'){
			callback = options;
			options = null;
		}
		if(err) {callback(err); return;}
		fs.read(fd,buf,0,peekSz,0,function(err,bytesRead,buffer){
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

var helper = exports.common = {}
//convert zero-based column index to the corresponding column name
var aToz="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
helper.toColumnName = function(columnIndex){
    assert(columnIndex >= 0,'columnIndex >= 0');
    var name = '';
    while (1){
        var quot = Math.floor(columnIndex/26), rem = columnIndex % 26;
        name = aToz[rem] + name;
        if (!quot)
            return name;
        columnIndex = quot - 1;
	}
}

helper.toBiffVersionString = function(ver){
	return comm.biffCodeStringMap[ver];
}

helper.toCountryName = function(countryCode){
    return comm.windowsCountryIdList[countryCode];
} 
