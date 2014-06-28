'use strict'
var fs = require('fs'),
    util = require('util'),
    assert = require('assert');

var SIGNATURE = exports.SIGNATURE = "D0CF11E0A1B11AE1".toLowerCase();

var EOCSID = -2,
	FREESID = -1,
	SATSID = -3,
	MSATSID = -4,
	EVILSID = -5;

function likeBuffer(fd, size){
	
	if(Buffer.isBuffer(fd)){
		var buf = fd;
		if(size === null || size === undefined)  size = buf.length;
		var len = 0;
		var _start = 0;
		buf.pull = function(start,end){
			_start = start;
			return len = end - start;
		};
		buf.getUInt8= function(position){
			if(position >= len) throw new Error('likeBuffer Error : position %d of readUInt8 beyond buffer length %d', position, len );
			return buf.readUInt8(_start + position);
		}
		buf.getUInt16LE= function(position){
			if(position >= len +1) throw new Error('likeBuffer Error : position %d of readUInt16LE beyond buffer length %d', position, len );
			return buf.readUInt16LE(_start + position);
		}
		buf.getInt16LE= function(position){
			if(position >= len +1) throw new Error('likeBuffer Error : position %d of readInt16LE beyond buffer length %d', position, len );		
			return buf.readInt16LE(_start + position);
		}
		buf.getInt32LE= function(position){
			if(position >= len+3) throw new Error('likeBuffer Error : position %d of readInt32LE beyond buffer length %d', position, len );
			return buf.readInt32LE(_start + position);
		}
		buf.toStr = function(encoding, start, end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error('likeBuffer Error : start %d of toString beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('likeBuffer Error : end %d of toString beyond buffer length %d', end, len);		
			
			return buf.toString(encoding,_start+start,_start+end);
		}
		buf.toInt32LEs = function(start,end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error('likeBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('likeBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var ret = [];
			for(var i = _start+start; i < _start+end; i += 4) ret.push(buf.readInt32LE(i));
			return ret;
		}
		buf.toBuf = function(start,end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error('likeBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('likeBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var length = end - start;
			var ret = new Buffer(length);
			buf.copy(ret,0,_start + start, _start + end);
			return  ret;
		}
		return buf;
	}else{
		if(size === null || size === undefined)  size = 8;
		var fb = new FileBuffer(fd, size);
		return fb;
	}
	
	function FileBuffer(fd, defaultBufferSize){
		var self = this;
		var buf = new Buffer(defaultBufferSize);
		var len = 0;
		
		//access directly 
		var b4 = new Buffer(4);
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
		self.pull = function(start,end){
			len = end -start;
			if(buf.length < len) buf = Buffer.concat([buf,new Buffer(len-buf.length)] ,len);
			return fs.readSync(fd,buf,0,len,start);
		}
		self.getUInt8= function(position){
			if(position >= len) throw new Error('FileBuffer Error : position %d of readUInt8 beyond buffer length %d', position, len );
			return buf.readUInt8(position);
		}
		self.getUInt16LE= function(position){
			if(position >= len +1) throw new Error('FileBuffer Error : position %d of readUInt16LE beyond buffer length %d', position, len );
			return buf.readUInt16LE(position);
		}
		self.getInt16LE= function(position){
			if(position >= len+1) throw new Error('FileBuffer Error : position %d of readInt16LE beyond buffer length %d', position, len );		
			return buf.readInt16LE(position);
		}
		self.getInt32LE= function(position){
			if(position >= len+3) throw new Error('FileBuffer Error : position %d of readInt32LE beyond buffer length %d', position, len );
			return buf.readInt32LE(position);
		}
		self.toStr = function(encoding, start, end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error(util.format('FileBuffer Error : start %d of toString beyond buffer length %d', start, len));		
			if(end && end > len) throw new Error(util.format('FileBuffer Error : end %d of toString beyond buffer length %d', end, len));		
			
			return buf.toString(encoding,start,end);
		}
		self.toInt32LEs = function(start,end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error('FileBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('FileBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var ret = [];
			for(var i = start; i < end; i += 4) ret.push(buf.readInt32LE(i));
			return ret;
		}
		self.toBuf = function(start,end){
			if(start === null || start === undefined) start = 0;
			if(end === null || end === undefined) end = len;
			if(start && start >= len) throw new Error('FileBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('FileBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
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
		
	}
}

//parameter dent is instance Of Buffer
function DirNode(DID, dent){
	var self = this;
	// dent is the 128-byte directory entry
	self.DID = DID;

	var cBufSize = dent.readUInt16LE(64);
	self.etype = dent.readUInt8(66);	
	self.colour = dent.readUInt8(67);
	self.leftDID = dent.readInt32LE(68);
	self.rightDID = dent.readInt32LE(72);
	self.rootDID = dent.readInt32LE(76);
	
	self.firstSID = dent.readInt32LE(116);
	self.totalSize = dent.readInt32LE(120);
	if(cBufSize == 0) self.name = '';//UNICODE_LITERAL('')
	else self.name = dent.toString('utf16le',0,cBufSize-2);// unicode(dent[0:cBufSize-2], 'utf_16_le'); // omit the trailing U+0000
	self.children = []; //# filled in later
	self.parent = -1;// # indicates orphan; fixed up later
	//self.tsinfo = unpack('<IIII', dent[100:116])
	//if DEBUG:
	//	self.dump(DEBUG)

	/*
    def dump(self, DEBUG=1):
        fprintf(
            self.logfile,
            "DID=%d name=%r etype=%d DIDs(left=%d right=%d root=%d parent=%d kids=%r) firstSID=%d totalSize=%d\n",
            self.DID, self.name, self.etype, self.leftDID,
            self.rightDID, self.rootDID, self.parent, self.children, self.firstSID, self.totalSize
            )
        if DEBUG == 2:
            # cre_lo, cre_hi, mod_lo, mod_hi = tsinfo
            print("timestamp info", self.tsinfo, file=self.logfile)
	*/
}
function buildFamilyTree(dirList, parentDID, childDID){
    if (childDID < 0) return;
    buildFamilyTree(dirList, parentDID, dirList[childDID].leftDID);
    dirList[parentDID].children.push(childDID);
    dirList[childDID].parent = parentDID;
    buildFamilyTree(dirList, parentDID, dirList[childDID].rightDID);
    if( dirList[childDID].etype == 1) // storage
        buildFamilyTree(dirList, childDID, dirList[childDID].rootDID);
}
		
exports.CompDoc = function (fd, totalSize){
	var self = this;
	var wbuf = likeBuffer(fd,436);
	wbuf.pull(0,8);
	var sig =wbuf.toStr('hex',0,8);
	
	if( sig != SIGNATURE) throw new Error('Not an OLE2 compound document');
	wbuf.pull(28,30);
	if(wbuf.toStr('hex',0,2) != 'feff')
		throw new Error('Expected "little-endian" marker, found %r' + wbuf.toStr('hex',0,2));
	wbuf.pull(24,28);
	var revision = wbuf.getUInt16LE(0),
		version = wbuf.getUInt16LE(2);
	wbuf.pull(30,34);
	var ssz = wbuf.getUInt16LE(0),
		sssz = wbuf.getUInt16LE(2);
	if (ssz > 20){ // allows for 2**20 bytes) i.e. 1MB
        console.log("WARNING: sector size (2**%d) is preposterous; assuming 512 and continuing ...",ssz);
        ssz = 9;
	}
	if (sssz > ssz){
        console.log("WARNING: short stream sector size (2**%d) is preposterous; assuming 64 and continuing ...",sssz);
        sssz = 6;
	}
	var secSize =  self.secSize =1 << ssz;
    var shortSecSize = 1 << sssz;
    if (self.secSize != 512 || shortSecSize != 64)
            console.log("@@@@ secSize=%d short_sec_size=%d",self.secSize, shortSecSize);
    wbuf.pull(44,76);
    var satTotalSecs = wbuf.getInt32LE(0);
	var dirFirstSecSid = wbuf.getInt32LE(4);
	var _unused = wbuf.getInt32LE(8);
	var minSizeStdStream = wbuf.getInt32LE(12);
    var ssatFirstSecSid = wbuf.getInt32LE(16);
	var ssatTotalSecs = wbuf.getInt32LE(20);
    var msatxFirstSecSid = wbuf.getInt32LE(24);
	var msatxTotalSecs = wbuf.getInt32LE(28);
        
    var memDataLen = totalSize - 512;
	var memDataSecs = Math.floor(memDataLen / secSize) ,
		leftOver =  memDataLen % secSize;
	if (leftOver){
		//#### raise CompDocError("Not a whole number of sectors")
		memDataSecs += 1;
		console.log("WARNING *** file size (%d) not 512 + multiple of sector size (%d)",totalSize, secSize);
	}
	self.memDataSecs = memDataSecs; // use for checking later
	self.memDataLen = memDataLen;
	var seen = self.seen = new Array(memDataSecs);
	for(var i =0; i < memDataSecs; i++ ) seen[i] = 0;
	/*
	if DEBUG:
		print('sec sizes', ssz, sssz, secSize, shortSecSize, file=logfile)
		print("mem data: %d bytes == %d sectors" % (memDataLen, memDataSecs), file=logfile)
		print("satTotalSecs=%d, dir_first_sec_sid=%d, min_size_std_stream=%d" \
			% (satTotalSecs, dirFirstSecSid, minSizeStdStream,), file=logfile)
		print("ssatFirstSecSid=%d, ssatTotalSecs=%d" % (ssatFirstSecSid, ssatTotalSecs,), file=logfile)
		print("msatxFirstSecSid=%d, msatxTotalSecs=%d" % (msatxFirstSecSid, msatxTotalSecs,), file=logfile)
	*/
	var nent = Math.floor(secSize/4) ;// number of SID entries in a sector
	var truncWarned = 0;
	//
	// === build the MSAT ===
	//
	wbuf.pull(76,512);
	var MSAT = wbuf.toInt32LEs(0,436);
	
	var satSectorsReqd = Math.floor((memDataSecs + nent - 1)/ nent);
	var expectedMsatxSecs = Math.max(0, Math.floor((satSectorsReqd - 109 + nent - 2)/ (nent - 1)));
	var actualMsatxSecs = 0;
	
	if( msatxTotalSecs == 0  &&  [EOCSID, FREESID, 0].indexOf(msatxFirstSecSid) >=0)
		// Strictly, if there is no MSAT extension, then msatxFirstSecSid
		// should be set to EOCSID ... FREESID and 0 have been met in the wild.
		;//pass # Presuming no extension
	else{
		sid = msatxFirstSecSid;
		
		while( EOCSID != sid && FREESID != sid){
			// Above should be only EOCSID according to MS & OOo docs
			// but Excel doesn't complain about FREESID. Zero is a valid
			// sector number, not a sentinel.
			
			//if (DEBUG > 1)
			//	console.log('MSATX: sid=%d (0x%08X)',sid, sid)
			if (sid >= memDataSecs){
				msg = util.format("MSAT extension: accessing sector %d but only %d in file",sid, memDataSecs);
				//if DEBUG > 1:
				//	print(msg, file=logfile)
				//	break
				throw new  Error(msg);
			}
			else if (sid < 0)
				throw new Error( util.format("MSAT extension: invalid sector id: %d", sid));
			if (seen[sid])
				throw new Error( util.format("MSAT corruption: seen[%d] == %d" % (sid, seen[sid])));
			seen[sid] = 1;
			actualMsatxSecs += 1;
			//if DEBUG and actualMsatxSecs > expectedMsatxSecs:
			//	print("[1]===>>>", memDataSecs, nent, satSectorsReqd, expectedMsatxSecs, actualMsatxSecs, file=logfile)
			
			var offset = 512 + secSize * sid;
			wbuf.pull(offset,offset+secSize);
			var buf = wbuf.toInt32LEs();
			sid = buf.pop();
			MSAT = MSAT.concat(buf);
		}
	}

	//if DEBUG and actualMsatxSecs != expectedMsatxSecs:
	//	print("[2]===>>>", memDataSecs, nent, satSectorsReqd, expectedMsatxSecs, actualMsatxSecs, file=logfile)
	//if DEBUG:
	//	print("MSAT: len =", len(MSAT), file=logfile)
	//	dump_list(MSAT, 10, logfile)
	//
	// === build the SAT ===
	//
	var sat = [];
	var actualSatSecs = 0;
	var dumpAgain = 0;
	for(var msidx =0; msidx < MSAT.length; msidx++){
		var msid = MSAT[msidx];
		if (msid == FREESID || msid ==EOCSID)
			// Specification: the MSAT array may be padded with trailing FREESID entries.
			// Toleration: a FREESID or EOCSID entry anywhere in the MSAT array will be ignored.
			continue;
		if (msid >= memDataSecs){
			if (!truncWarned){
				console.log("WARNING *** File is truncated, or OLE2 MSAT is corrupt!!");
				console.log(util.format("INFO: Trying to access sector %d but only %d available",
				msid, memDataSecs));
				truncWarned = 1;
			}
			MSAT[msidx] = EVILSID;
			dumpAgain = 1;
			continue;
		}else if (msid < -2)
			console.log(util.format("MSAT: invalid sector id: %d", msid));
		if (seen[msid])
			console.log(util.format("MSAT extension corruption: seen[%d] == %d",msid, seen[msid]));
		seen[msid] = 2;
		actualSatSecs += 1;
		//if DEBUG and actualSatSecs > satSectorsReqd:
		//	print("[3]===>>>", memDataSecs, nent, satSectorsReqd, expectedMsatxSecs, actualMsatxSecs, actual_SAT_sectors, msid, file=logfile)

		var offset = 512 + secSize * msid;
		wbuf.pull(offset,offset+secSize);
		sat = sat.concat(wbuf.toInt32LEs());
	}
	/*
	if DEBUG:
		print("SAT: len =", len(sat), file=logfile)
		dump_list(sat, 10, logfile)
		# print >> logfile, "SAT ",
		# for i, s in enumerate(sat):
			# print >> logfile, "entry: %4d offset: %6d, next entry: %4d" % (i, 512 + secSize * i, s)
			# print >> logfile, "%d:%d " % (i, s),
		print(file=logfile)
	if DEBUG and dumpAgain:
		print("MSAT: len =", len(MSAT), file=logfile)
		dump_list(MSAT, 10, logfile)
		for satx in xrange(memDataSecs, len(sat)):
			sat[satx] = EVILSID
		print("SAT: len =", len(sat), file=logfile)
		dump_list(sat, 10, logfile)
	*/
	        
	//
	// === build the directory ===
	//
	var dbytes = getStream(self,wbuf, 512, sat, self.secSize, dirFirstSecSid,null,"directory",3);
	var dirList = [];
	for(var did =0, pos = 0; pos < dbytes.length; pos += 128){
		dirList.push(new DirNode(did++, dbytes.slice(pos,pos+128)));
	}
    self.dirList = dirList;
	buildFamilyTree(dirList, 0, dirList[0].rootDID); // and stand well back ...
	/*
	if DEBUG:
		for d in dirList:
			d.dump(DEBUG)
	*/
	
	//
	// === get the SSCS ===
	//
	var sscsDir = dirList[0];
	assert(sscsDir.etype == 5 ,'sscsDir.etype == 5 '); // root entry
	if(sscsDir.firstSID < 0 || sscsDir.totalSize == 0){
		self.SSCS = new Buffer(0);
	}else{
		self.SSCS = getStream(self,wbuf,512, sat, secSize, sscsDir.firstSID,sscsDir.totalSize, "SSCS", 4);
	}
	// if DEBUG: print >> logfile, "SSCS", repr(self.SSCS)
	
	//
	// === build the SSAT ===
	//
	var ssat = [];
	var sectors =[];
	
	if(ssatTotalSecs > 0 && sscsDir.totalSize == 0)
		console.log("WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero");
	if (sscsDir.totalSize > 0){
		var sid = ssatFirstSecSid;
		var nsecs = ssatTotalSecs;
		while( sid >= 0 && nsecs > 0){
			if(seen[sid])
				throw new  Error(util.format("SSAT corruption: seen[%d] == %d" ,sid, seen[sid]));
			seen[sid] = 5;
			nsecs -= 1;
			var startPos = 512 + sid * secSize;
			wbuf.pull(startPos,startPos+secSize);
			ssat = ssat.concat(wbuf.toInt32LEs());
			
			sid = sat[sid];
		}
		//if DEBUG: print("SSAT last sid %d; remaining sectors %d" % (sid, nsecs), file=logfile)
		assert(nsecs == 0 && sid == EOCSID, 'nsecs == 0 && sid == EOCSID');
	}
	/*
	if DEBUG:
		print("SSAT", file=logfile)
		dump_list(ssat, 10, logfile)
	if DEBUG:
		print("seen", file=logfile)
		dump_list(seen, 20, logfile)
	*/
	
	
	/*
	* Interrogate the compound document's directory; return Buffer if found, otherwise
	* return null.
	* @param qname Name of the desired stream e.g. u'Workbook'. Should be in Unicode or convertible thereto.
	*/
	self.getNamedStream = function (qname){
		var d = dirSearch(self,qname.split("/"));
		if( d === null ) return null;
		if( d.totalSize >= minSizeStdStream)
			return getStream(self, wbuf, 512, sat, self.secSize, d.firstSID,d.totalSize, qname, d.DID+6);
		else
			return getStream(self, likeBuffer(self.SSCS), 0, ssat, shortSecSize, d.firstSID,
				d.totalSize, name=qname + " (from SSCS)", null)
	}
	/*
	* Interrogate the compound document's directory.
	* If the named stream is not found, (null, 0, 0) will be returned.
	* If the named stream is found and is contiguous within the original byte sequence ("mem")
	* used when the document was opened,
	* then (mem, offset_to_start_of_stream, length_of_stream) is returned.
	* Otherwise a new string is built from the fragments and (new_string, 0, length_of_stream) is returned.
	* @param qname Name of the desired stream e.g. u'Workbook'. Should be in Unicode or convertible thereto.
	*/
	self.locateNamedStream = function(qname){
		var d = dirSearch(self,qname.split("/"));
		if (d === null) return [null, 0, 0];
		if (d.totalSize > self.memDataLen)
			throw new  Error(util.format("%r stream length (%d bytes) > file data size (%d bytes)",
				qname, d.totalSize, self.memDataLen));
		if (d.totalSize >= minSizeStdStream){
			/*
			if self.DEBUG:
				print("\nseen", file=self.logfile)
				dump_list(self.seen, 20, self.logfile)
			*/
			return locateStream(self, wbuf, 512, sat, self.secSize, d.firstSID, 
				d.totalSize, qname, d.DID+6);
		}else{
			return [
				getStream(self, likeBuffer(self.SSCS), 0, ssat, shortSecSize, d.firstSID,d.totalSize, qname + " (from SSCS)", null),
				0,
				d.totalSize
			];
		}
	}
}

//return buffer;
function getStream(compDoc, wbuf, base, sat, secSize, startSid, size, name, seenId){
	// print >> self.logfile, "getStream", base, secSize, startSid, size
	var sectors = [];
	var s = startSid;
	var totalSz = 0;var i=0;
	name = name ||'';
	if (size === null || size === undefined){
		// nothing to check against
		while( s >= 0){
			if(seenId !== null && seenId !== undefined){
				if (compDoc.seen[s])
					throw new  Error(util.format("%s corruption: seen[%d] == %d" ,name, s, compDoc.seen[s]));
				compDoc.seen[s] = seenId;
			}
			var startPos = base + s * secSize;

			wbuf.pull(startPos,startPos+secSize);
			sectors.push(wbuf.toBuf());

		    totalSz += secSize;
		    if(sat.length <= s) throw new Error(util.format("OLE2 stream %r: sector allocation table invalid entry (%d)",name, s));
		    s = sat[s];
		}
		assert(s == EOCSID, 's == EOCSID');
	}else{
		var todo = size;
		while(s >= 0){
			if(seenId !== null && seenId !== undefined){
				if (compDoc.seen[s])
					throw new  Error(util.format("%s corruption: seen[%d] == %d" ,name, s, compDoc.seen[s]));
				compDoc.seen[s] = seenId;
			}
			var startPos = base + s * secSize;
			var grab = secSize;
			if (grab > todo) grab = todo;
			todo -= grab;
			wbuf.pull(startPos,startPos+grab);
			sectors.push(wbuf.toBuf());
		    totalSz += grab;
		    if(sat.length <= s) throw new Error(util.format("OLE2 stream %r: sector allocation table invalid entry (%d)",name, s));
		    s = sat[s];
		}
		assert(s == EOCSID, 's == EOCSID');
		if (todo != 0)
			console.log(util.format("WARNING *** OLE2 stream %r: expected size %d, actual size %d\n",name, size, size - todo));
	}
	return Buffer.concat(sectors,totalSz);
}

function dirSearch(compDoc, path, storageDID){
	// Return matching DirNode instance, or None
	if(storageDID === null || storageDID === undefined) storageDID = 0;  
	var head = path[0];
	var tail = path.slice(1);
	var dl = compDoc.dirList;
	var length = dl[storageDID].children.length;
	for(var i =0; i < length; i++){
		var child = dl[storageDID].children[i];
		if (dl[child].name.toLowerCase() == head.toLowerCase()){
			var et = dl[child].etype;
			if (et == 2) return dl[child];
			if (et == 1) {
				if (! tail) throw new Error("Requested component is a 'storage'");
				return dirSearch(compDoc, tail, child);
			}
			//dl[child].dump(1);
			throw new Error("Requested stream is not a 'user stream'");
		}
	}
	return null;
}
	


function locateStream(compDoc, wbuf, base, sat, secSize, startSid, expectedStreamSize, qname, seenId){
        // print >> self.logfile, "locateStream", base, secSize, startSid, expectedStreamSize
	var s = startSid;
	if (s < 0) throw new  Error(util.format("locateStream: startSid (%d) is -ve", startSid));
	var p = -99 ; // dummy previous SID
	var startPos = -9999;
	var endPos = -8888;
	var slices = [];
	var totalFound = 0;
	var foundLimit = Math.floor((expectedStreamSize + secSize - 1)/ secSize);
	var totalSz =0;
	while (s >= 0){
		if(compDoc.seen[s]){
			//console.log(util.format("locateStream(%s): seen" , qname); dump_list(self.seen, 20, self.logfile)
			throw new  Error(util.format("%s corruption: seen[%d] == %d",qname, s, compDoc.seen[s]));
		}
		compDoc.seen[s] = seenId;
		totalFound += 1;
		if (totalFound > foundLimit){
			throw new  Error(util.format(
				"%s: size exceeds expected %d bytes; corrupt?"
				, qname, foundLimit * secSize)
				); // Note: expected size rounded up to higher sector
		}
		if( s == p+1)
			// contiguous sectors
			endPos += secSize;
		else{
			// start new slice
			if( p >= 0){
				// not first time
				var len = endPos - startPos;
				slices.push([startPos, endPos]);
				totalSz += len;
			}
			startPos = base + s * secSize;
			endPos = startPos + secSize;
		}
		p = s;
		s = sat[s];
	}
	assert(s == EOCSID,'s == EOCSID');
	assert(totalFound == foundLimit,'totalFound == foundLimit');
	// print >> self.logfile, "locateStream(%s): seen" % qname; dump_list(self.seen, 20, self.logfile)
	if (!slices.length){
		// The stream is contiguous ... just what we like!
		wbuf.pull(startPos,startPos+expectedStreamSize);
		wbuf.hasData = true;
		return [wbuf, startPos, expectedStreamSize];
	}
	slices.push([startPos, endPos]);
	// print >> self.logfile, "+++>>> %d fragments" % len(slices)
	var ret = [];
	slices.forEach(function(x){
		wbuf.pull(x[0], x[1]);
		ret.push(wbuf.toBuf());
	});
	return [Buffer.concat(ret,totalSz), 0, expectedStreamSize];
}