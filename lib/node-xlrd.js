'use strict'
var fs = require('fs'),
	util = require('util'),
	book = require('./workbook'),
	comm = require('./common'),
	assert = require('assert');

//open(filename, function(err, workbook))
exports.open = function(filename, callback){
	var peeksz = 4;
	var buf = new Buffer(4);
	fs.open(filename, 'r',function(err,fd){
		if(err) {callback(err); return;}
		fs.read(fd,buf,0,peeksz,0,function(err,bytesRead,buffer){
			if(err) {callback(err); return;}
			var opts = {};
			var bk = null;
			if(buffer[0] == 'P' && buffer[1] == 'K' && buffer[2] == 0x03 && buffer[3] == 0x04){
				//xlsx
				callback(new Error('not yet'));
			}else{
				//xls
				bk = book.openWorkbook(fd,opts);
			}
			fs.close(fd,function(err){ 
				if(err) {callback(err); return;}
				callback(null,bk);
			});
		});
		
	});
	
}

// === helper functions
var a_z="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
exports.colname = function(colx){
    assert(colx >= 0,'colx >= 0');
    var name = '';
    while (1){
        var quot = Math.floor(colx/26), rem = colx % 26;
        name = a_z[rem] + name;
        if (!quot)
            return name;
        colx = quot - 1;
	}
}

exports.toBiffVersionText = function(ver){
	return comm.biff_text_from_num[ver];
}