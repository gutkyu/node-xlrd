'use strict';
var fs = require('fs');
var util = require('util');
var book = require('./workbook');
var comm = require('./common');
var assert = require('assert');
require('./utils/expandedUint8Array');

//open(fileName, function(err, workbook))
exports.open = function (fileName, options, callback) {
  checkParams(options, callback, function (err, opts, cb) {
    if (err) return cb(err);
    readFile(fileName, function (err, fd, head4bytes) {
      if (err) return cb(err);
      checkXLType(head4bytes, function (err) {
        if (err) return cb(err);
        var bk = book.openWorkbook(fd, opts);
        cb(null, bk);
        finalize(fd, bk, cb);
      });
    });
  });
};
exports.create = function (fileName, options) {
  var opts = checkParams(options);
  var xlFdHead4 = readFileSync(fileName);
  checkXLType(xlFdHead4.head4);
  return book.openWorkbook(xlFdHead4.fd, opts);
};
function checkParams(options, callback, next) {
  if (typeof options == 'function') {
    callback = options;
    options = null;
  }
  var opts = options || {}; //todo: default 값으로 병합
  opts.eventMode =
    opts.eventMode === null || opts.eventMode === undefined
      ? false
      : opts.eventMode;
  opts.raggedRows =
    opts.raggedRows === null || opts.raggedRows === undefined
      ? true
      : opts.raggedRows; //raggedRows default : true
  if (opts.eventMode && !opts.raggedRows) {
    var err = new Error(
      'if eventMode == true, options.raggedRows should be true.'
    );
    if (next) return next(err);
    else throw err;
  }
  if (next) next(null, opts, callback);
  else return opts;
}
function readFile(fileName, next) {
  var peekSz = 4;
  var buf = new Uint8Array(peekSz);
  fs.open(fileName, 'r', function (err, fd) {
    if (err) {
      return next(err);
    }
    fs.read(fd, buf, 0, peekSz, 0, function (err, bytesRead, buffer) {
      if (err) {
        return next(err);
      }
      next(null, fd, buffer);
    });
  });
}
function readFileSync(fileName) {
  var peekSz = 4;
  var buf = new Uint8Array(peekSz);
  var fd = fs.openSync(fileName, 'r');
  fs.readSync(fd, buf, 0, peekSz, 0);
  return {fd: fd, head4: buf};
}
function checkXLType(buffer, next) {
  if (
    buffer[0] == 'P' &&
    buffer[1] == 'K' &&
    buffer[2] == 0x03 &&
    buffer[3] == 0x04
  ) {
    //xlsx
    var err = new Error('.xlsx not yet implemented');
    if (next) callback(err);
    else throw err;
  }
  if (next) next();
}
function finalize(fd, bk, next) {
  if (!next) fs.closeSync(fd);
  fs.close(fd, function (err) {
    if (err) {
      console.log('file close error : ' + err);
      bk._release();
      return;
    }
    bk._release();
  });
}

// === helper functions
var helper = (exports.common = {});
//convert zero-based column index to the corresponding column name
var aToz = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
helper.toColumnName = function (columnIndex) {
  assert(columnIndex >= 0, 'columnIndex >= 0');
  var name = '';
  while (true) {
    var quot = Math.floor(columnIndex / 26),
      rem = columnIndex % 26;
    name = aToz[rem] + name;
    if (!quot) return name;
    columnIndex = quot - 1;
  }
};

helper.toBiffVersionString = function (ver) {
  return comm.biffCodeStringMap[ver];
};

helper.toCountryName = function (countryCode) {
  return comm.windowsCountryIdList[countryCode];
};
