'use strict';
var stringDecode = require('./StringEncoding').decode;

if (Uint8Array === undefined) throw new Error('Uint8Array not found');

Uint8Array.concat = function (arrayList, totalSize) {
  if (totalSize === undefined)
    totalSize = arrayList.reduce(function (acc, cur) {
      return acc + cur.length;
    }, 0);
  var ret = new Uint8Array(totalSize);
  var offset = 0;
  arrayList.forEach(function (x) {
    ret.set(x, offset);
    offset += x.length;
  });
  return ret;
};

Uint8Array.prototype.equals = function (array) {
  var self = this;
  if (self.length !== array.length) return false;
  return self.every(function (val, idx) {
    return val === array[idx];
  });
};
Uint8Array.prototype.getU8 = function (offset) {
  var self = this;
  if (!self._dv) self._dv = new DataView(self.buffer);
  return self._dv.getUint8(offset);
};
Uint8Array.prototype.getU16L = function (offset) {
  var self = this;
  if (!self._dv) self._dv = new DataView(self.buffer);
  return self._dv.getUint16(offset, true);
};
Uint8Array.prototype.getU32L = function (offset) {
  var self = this;
  if (!self._dv) self._dv = new DataView(self.buffer);
  return self._dv.getUint32(offset, true);
};
Uint8Array.prototype.getI16L = function (offset) {
  var self = this;
  if (!self._dv) self._dv = new DataView(self.buffer);
  return self._dv.getInt16(offset, true);
};
Uint8Array.prototype.getI32L = function (offset) {
  var self = this;
  if (!self._dv) self._dv = new DataView(self.buffer);
  return self._dv.getInt32(offset, true);
};
Uint8Array.prototype.getStr = function (encoding, start, end) {
  var self = this;
  if (start === null || start === undefined) start = 0;
  if (!Number.isInteger(start))
    throw new Error('Uint8Array getStr error : start is not integer');
  if (end === null || end === undefined) end = self.length;
  if (!Number.isInteger(end))
    throw new Error('Uint8Array getStr error : end is not integer');
  return stringDecode(self, start, end - start, encoding);
};

Uint8Array.prototype.getLEReader = function () {
  var self = this;
  return new Uint8ArrayReader(self, true);
};

function Uint8ArrayReader(typedArray, isLittleEndian) {
  var self = this;
  self._isLittleEndian =
    isLittleEndian === null || isLittleEndian === undefined
      ? true
      : isLittleEndian;
  if (typedArray instanceof Uint8Array) {
    if (!typedArray.buffer._dv)
      typedArray.buffer._dv = new DataView(typedArray.buffer);
    self._dv = typedArray.buffer._dv;
  } else throw new Error('Uint8ArrayReader Error : Unsupported Type');
  self._cur = 0;
}

//src is Uint8Array
Uint8ArrayReader.prototype.str = function (encoding) {};

Uint8ArrayReader.prototype.mov = function (offset) {
  var self = this;
  self._cur = offset;
  return self;
};

Object.defineProperty(Uint8ArrayReader.prototype, 'u8', {
  get: function () {
    var self = this;
    return self._dv.getUint8(self._cur++);
  },
});
Object.defineProperty(Uint8ArrayReader.prototype, 'u16', {
  get: function () {
    var self = this;
    var ret = self._dv.getUint16(self._cur, self._isLittleEndian);
    self._cur += 2;
    return ret;
  },
});
Object.defineProperty(Uint8ArrayReader.prototype, 'u32', {
  get: function () {
    var self = this;
    var ret = self._dv.getUint32(self._cur, self._isLittleEndian);
    self._cur += 4;
    return ret;
  },
});
Object.defineProperty(Uint8ArrayReader.prototype, 'i16', {
  get: function () {
    var self = this;
    var ret = self._dv.getInt16(self._cur, self._isLittleEndian);
    self._cur += 2;
    return ret;
  },
});
Object.defineProperty(Uint8ArrayReader.prototype, 'i32', {
  get: function () {
    var self = this;
    var ret = self._dv.getInt32(self._cur, self._isLittleEndian);
    self._cur += 4;
    return ret;
  },
});
Object.defineProperty(Uint8ArrayReader.prototype, 'dbl', {
  get: function () {
    var self = this;
    var ret = self._dv.getFloat64(self._cur, self._isLittleEndian);
    self._cur += 4;
    return ret;
  },
});
