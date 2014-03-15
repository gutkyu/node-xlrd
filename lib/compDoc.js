var fs = require('fs'),
    util = require('util'),
    assert = require('assert');

exports.SIGNATURE = SIGNATURE = "D0CF11E0A1B11AE1".toLowerCase();

var EOCSID = -2,
	FREESID = -1,
	SATSID = -3,
	MSATSID = -4,
	EVILSID = -5;

function likeBuffer(fd, size){
	
	if(Buffer.isBuffer(fd)){
		var buf = fd;
		if(size == null || size == undefined)  size = buf.length;
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
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error('likeBuffer Error : start %d of toString beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('likeBuffer Error : end %d of toString beyond buffer length %d', end, len);		
			
			return buf.toString(encoding,_start+start,_start+end);
		}
		buf.toInt32LEs = function(start,end){
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error('likeBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('likeBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var ret = [];
			for(var i = _start+start; i < _start+end; i += 4) ret.push(buf.readInt32LE(i));
			return ret;
		}
		this.toBuf = function(start,end){
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error('FileBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('FileBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var length = end - start;
			var ret = new Buffer(length);
			fs.readSync(fd,ret,0,length,_start+start);
			return ret;
		}
		return buf;
	}else{
		if(size == null || size == undefined)  size = 8;
		var fb = new FileBuffer(fd, size);
		return fb;
	}
	
	function FileBuffer(fd, defaultBufferSize){
		var buf = new Buffer(defaultBufferSize);
		var len = 0;
		
		//access directly 
		var b4 = new Buffer(4);
		this.readUInt8 = function(position){
			fs.readSync(fd,b4,0,1,position);
			return b4.getUInt8(0);
		}
		
		this.readUInt16LE = function(position){
			fs.readSync(fd,b4,0,2,position);
			return b4.readUInt16LE(0);
		}
		
		this.readInt16LE = function(position){
			fs.readSync(fd,b4,0,2,position);
			return b4.getInt16LE(0);
		}
		
		this.readInt32LE = function(position){
			fs.readSync(fd,b4,0,4,position);
			return b4.getInt32LE(0);
		}

		this.slice = function(start, end){
			var length = end - start;
			var buf = new Buffer(length);
			if(length) fs.readSync(fd,buf,0,length,start);
			return buf;
		}
		var fileSize = null;
		Object.defineProperty(this,'length',{get:function(){
			
			if(fileSize == null || fileSize == undefined){
				var stats = fs.fstatSync(fd);
				if(!stats.size) throw new Error('File size is 0 bytes');
				fileSize = stats.size;
			}
		return fileSize;}});
		
		//use buffer
		this.pull = function(start,end){
			len = end -start;
			if(buf.length < len) buf = Buffer.concat([buf,new Buffer(len-buf.length)] ,len);
			return fs.readSync(fd,buf,0,len,start);
		}
		this.getUInt8= function(position){
			if(position >= len) throw new Error('FileBuffer Error : position %d of readUInt8 beyond buffer length %d', position, len );
			return buf.readUInt8(position);
		}
		this.getUInt16LE= function(position){
			if(position >= len +1) throw new Error('FileBuffer Error : position %d of readUInt16LE beyond buffer length %d', position, len );
			return buf.readUInt16LE(position);
		}
		this.getInt16LE= function(position){
			if(position >= len+1) throw new Error('FileBuffer Error : position %d of readInt16LE beyond buffer length %d', position, len );		
			return buf.readInt16LE(position);
		}
		this.getInt32LE= function(position){
			if(position >= len+3) throw new Error('FileBuffer Error : position %d of readInt32LE beyond buffer length %d', position, len );
			return buf.readInt32LE(position);
		}
		this.toStr = function(encoding, start, end){
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error(util.format('FileBuffer Error : start %d of toString beyond buffer length %d', start, len));		
			if(end && end > len) throw new Error(util.format('FileBuffer Error : end %d of toString beyond buffer length %d', end, len));		
			
			return buf.toString(encoding,start,end);
		}
		this.toInt32LEs = function(start,end){
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error('FileBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('FileBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var ret = [];
			for(var i = start; i < end; i += 4) ret.push(buf.readInt32LE(i));
			return ret;
		}
		this.toBuf = function(start,end){
			if(start == null || start == undefined) start = 0;
			if(end == null || end == undefined) end = len;
			if(start && start >= len) throw new Error('FileBuffer Error : start %d of toInt32LEs beyond buffer length %d', start, len);		
			if(end && end > len) throw new Error('FileBuffer Error : end %d of toInt32LEs beyond buffer length %d', end, len);		
			
			var length = end - start;
			return buf.slice(start,end);
		}
		
	}
}

//parameter dent is instance Of Buffer
function DirNode(DID, dent){
	// dent is the 128-byte directory entry
	this.DID = DID;
	
	var cbufsize = dent.readUInt16LE(64);
	this.etype = dent.readUInt8(66);
	this.colour = dent.readUInt8(67);
	this.left_DID = dent.readInt32LE(68);
	this.right_DID = dent.readInt32LE(72);
	this.root_DID = dent.readInt32LE(76);
	
	this.first_SID = dent.readInt32LE(116);
	this.tot_size = dent.readInt32LE(120);
	if(cbufsize == 0) this.name = '';//UNICODE_LITERAL('')
	else this.name = dent.toString('utf16le',0,cbufsize-2);// unicode(dent[0:cbufsize-2], 'utf_16_le'); // omit the trailing U+0000
	this.children = []; //# filled in later
	this.parent = -1;// # indicates orphan; fixed up later
	//this.tsinfo = unpack('<IIII', dent[100:116])
	//if DEBUG:
	//	this.dump(DEBUG)

	/*
    def dump(this, DEBUG=1):
        fprintf(
            this.logfile,
            "DID=%d name=%r etype=%d DIDs(left=%d right=%d root=%d parent=%d kids=%r) first_SID=%d tot_size=%d\n",
            this.DID, this.name, this.etype, this.left_DID,
            this.right_DID, this.root_DID, this.parent, this.children, this.first_SID, this.tot_size
            )
        if DEBUG == 2:
            # cre_lo, cre_hi, mod_lo, mod_hi = tsinfo
            print("timestamp info", this.tsinfo, file=this.logfile)
	*/
}
function buildFamilyTree(dirlist, parent_DID, child_DID){
    if (child_DID < 0) return;
    buildFamilyTree(dirlist, parent_DID, dirlist[child_DID].left_DID);
    dirlist[parent_DID].children.push(child_DID);
    dirlist[child_DID].parent = parent_DID;
    buildFamilyTree(dirlist, parent_DID, dirlist[child_DID].right_DID);
    if( dirlist[child_DID].etype == 1) // storage
        buildFamilyTree(dirlist, child_DID, dirlist[child_DID].root_DID);
}
		
exports.CompDoc = function (fd, totalSize){
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
	this.sec_size = sec_size = 1 << ssz;
    this.short_sec_size = 1 << sssz;
    if (this.sec_size != 512 || this.short_sec_size != 64)
            console.log("@@@@ sec_size=%d short_sec_size=%d",this.sec_size, this.short_sec_size);
    wbuf.pull(44,76);
    var SAT_tot_secs = wbuf.getInt32LE(0);
	this.dir_first_sec_sid = wbuf.getInt32LE(4);
	var _unused = wbuf.getInt32LE(8);
	this.min_size_std_stream = wbuf.getInt32LE(12);
    var SSAT_first_sec_sid = wbuf.getInt32LE(16);
	var SSAT_tot_secs = wbuf.getInt32LE(20);
    var MSATX_first_sec_sid = wbuf.getInt32LE(24);
	var MSATX_tot_secs = wbuf.getInt32LE(28);
        
    var mem_data_len = totalSize - 512;
	var mem_data_secs = Math.floor(mem_data_len / sec_size) ,
		left_over =  mem_data_len % sec_size;
	if (left_over){
		//#### raise CompDocError("Not a whole number of sectors")
		mem_data_secs += 1;
		console.log("WARNING *** file size (%d) not 512 + multiple of sector size (%d)",totalSize, sec_size);
	}
	this.mem_data_secs = mem_data_secs; // use for checking later
	this.mem_data_len = mem_data_len;
	var seen = this.seen = new Array(mem_data_secs);
	for(i =0; i < mem_data_secs; i++ ) seen[i] = 0;
	/*
	if DEBUG:
		print('sec sizes', ssz, sssz, sec_size, self.short_sec_size, file=logfile)
		print("mem data: %d bytes == %d sectors" % (mem_data_len, mem_data_secs), file=logfile)
		print("SAT_tot_secs=%d, dir_first_sec_sid=%d, min_size_std_stream=%d" \
			% (SAT_tot_secs, self.dir_first_sec_sid, self.min_size_std_stream,), file=logfile)
		print("SSAT_first_sec_sid=%d, SSAT_tot_secs=%d" % (SSAT_first_sec_sid, SSAT_tot_secs,), file=logfile)
		print("MSATX_first_sec_sid=%d, MSATX_tot_secs=%d" % (MSATX_first_sec_sid, MSATX_tot_secs,), file=logfile)
	*/
	var nent = Math.floor(sec_size/4) ;// number of SID entries in a sector
	var trunc_warned = 0;
	//
	// === build the MSAT ===
	//
	wbuf.pull(76,512);
	var MSAT = wbuf.toInt32LEs(0,436);
	
	var SAT_sectors_reqd = Math.floor((mem_data_secs + nent - 1)/ nent);
	var expected_MSATX_sectors = Math.max(0, Math.floor((SAT_sectors_reqd - 109 + nent - 2)/ (nent - 1)));
	var actual_MSATX_sectors = 0;
	
	if( MSATX_tot_secs == 0  &&  [EOCSID, FREESID, 0].indexOf(MSATX_first_sec_sid) >=0)
		// Strictly, if there is no MSAT extension, then MSATX_first_sec_sid
		// should be set to EOCSID ... FREESID and 0 have been met in the wild.
		;//pass # Presuming no extension
	else{
		sid = MSATX_first_sec_sid;
		
		while( EOCSID != sid && FREESID != sid){
			// Above should be only EOCSID according to MS & OOo docs
			// but Excel doesn't complain about FREESID. Zero is a valid
			// sector number, not a sentinel.
			
			//if (DEBUG > 1)
			//	console.log('MSATX: sid=%d (0x%08X)',sid, sid)
			if (sid >= mem_data_secs){
				msg = util.format("MSAT extension: accessing sector %d but only %d in file",sid, mem_data_secs);
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
			actual_MSATX_sectors += 1;
			//if DEBUG and actual_MSATX_sectors > expected_MSATX_sectors:
			//	print("[1]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors, file=logfile)
			
			var offset = 512 + sec_size * sid;
			wbuf.pull(offset,offset+sec_size);
			var buf = wbuf.toInt32LEs();
			sid = buf.pop();
			MSAT = MSAT.concat(buf);
		}
	}

	//if DEBUG and actual_MSATX_sectors != expected_MSATX_sectors:
	//	print("[2]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors, file=logfile)
	//if DEBUG:
	//	print("MSAT: len =", len(MSAT), file=logfile)
	//	dump_list(MSAT, 10, logfile)
	//
	// === build the SAT ===
	//
	this.SAT = [];
	var actual_SAT_sectors = 0;
	var dump_again = 0;
	for(msidx =0; msidx < MSAT.length; msidx++){
		msid = MSAT[msidx];
		if (msid == FREESID || msid ==EOCSID)
			// Specification: the MSAT array may be padded with trailing FREESID entries.
			// Toleration: a FREESID or EOCSID entry anywhere in the MSAT array will be ignored.
			continue;
		if (msid >= mem_data_secs){
			if (!trunc_warned){
				console.log("WARNING *** File is truncated, or OLE2 MSAT is corrupt!!");
				console.log(util.format("INFO: Trying to access sector %d but only %d available",
				msid, mem_data_secs));
				trunc_warned = 1;
			}
			MSAT[msidx] = EVILSID;
			dump_again = 1;
			continue;
		}else if (msid < -2)
			console.log(util.format("MSAT: invalid sector id: %d", msid));
		if (seen[msid])
			console.log(util.format("MSAT extension corruption: seen[%d] == %d",msid, seen[msid]));
		seen[msid] = 2;
		actual_SAT_sectors += 1;
		//if DEBUG and actual_SAT_sectors > SAT_sectors_reqd:
		//	print("[3]===>>>", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors, actual_SAT_sectors, msid, file=logfile)

		var offset = 512 + sec_size * msid;
		wbuf.pull(offset,offset+sec_size);
		this.SAT = this.SAT.concat(wbuf.toInt32LEs());
	}
	/*
	if DEBUG:
		print("SAT: len =", len(this.SAT), file=logfile)
		dump_list(self.SAT, 10, logfile)
		# print >> logfile, "SAT ",
		# for i, s in enumerate(self.SAT):
			# print >> logfile, "entry: %4d offset: %6d, next entry: %4d" % (i, 512 + sec_size * i, s)
			# print >> logfile, "%d:%d " % (i, s),
		print(file=logfile)
	if DEBUG and dump_again:
		print("MSAT: len =", len(MSAT), file=logfile)
		dump_list(MSAT, 10, logfile)
		for satx in xrange(mem_data_secs, len(self.SAT)):
			self.SAT[satx] = EVILSID
		print("SAT: len =", len(self.SAT), file=logfile)
		dump_list(self.SAT, 10, logfile)
	*/
	        
	//
	// === build the directory ===
	//
	var dbytes = getStream(this,wbuf, 512, this.SAT, this.sec_size, this.dir_first_sec_sid,null,"directory",3);
	var dirlist = [];
	var did = -1;
	for(pos = 0; pos < dbytes.length; pos += 128){
		did += 1;
		dirlist.push(new DirNode(did, dbytes.slice(pos,pos+128)));
	}
	this.dirlist = dirlist;
	buildFamilyTree(dirlist, 0, dirlist[0].root_DID); // and stand well back ...
	/*
	if DEBUG:
		for d in dirlist:
			d.dump(DEBUG)
	*/
	
	//
	// === get the SSCS ===
	//
	var sscs_dir = this.dirlist[0];
	assert(sscs_dir.etype == 5 ,'sscs_dir.etype == 5 '); // root entry
	if(sscs_dir.first_SID < 0 || sscs_dir.tot_size == 0){
		this.SSCS = "";
	}else{
		this.SSCS = getStream(this,wbuf,512, this.SAT, sec_size, sscs_dir.first_SID,sscs_dir.tot_size, "SSCS", 4);
	}
	// if DEBUG: print >> logfile, "SSCS", repr(self.SSCS)
	
	//
	// === build the SSAT ===
	//
	this.SSAT = null;
	var sectors =[];
	
	if(SSAT_tot_secs > 0 && sscs_dir.tot_size == 0)
		console.log("WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero");
	if (sscs_dir.tot_size > 0){
		var sid = SSAT_first_sec_sid;
		var nsecs = SSAT_tot_secs;
		while( sid >= 0 && nsecs > 0){
			if(seen[sid])
				throw new  Error(util.format("SSAT corruption: seen[%d] == %d" ,sid, seen[sid]));
			seen[sid] = 5;
			nsecs -= 1;
			var start_pos = 512 + sid * sec_size;
			
			wbuf.pull(start_pos,start_pos+sec_size);
			this.SSAT = this.SSAT.concat(wbuf.toInt32LEs());
			
			sid = this.SAT[sid];
		}
		//if DEBUG: print("SSAT last sid %d; remaining sectors %d" % (sid, nsecs), file=logfile)
		assert(nsecs == 0 && sid == EOCSID, 'nsecs == 0 && sid == EOCSID');
	}
	/*
	if DEBUG:
		print("SSAT", file=logfile)
		dump_list(self.SSAT, 10, logfile)
	if DEBUG:
		print("seen", file=logfile)
		dump_list(seen, 20, logfile)
	*/
	
	
	/*
	* Interrogate the compound document's directory; return Buffer if found, otherwise
	* return null.
	* @param qname Name of the desired stream e.g. u'Workbook'. Should be in Unicode or convertible thereto.
	*/
	this.get_named_stream = function (qname){
		var d = dirSearch(this,qname.split("/"));
		if( d == null ) return null;
		if( d.tot_size >= this.min_size_std_stream)
			return getStream(this, wbuf, 512, this.SAT, this.sec_size, d.first_SID,d.tot_size, qname, d.DID+6);
		else
			return getStream(this, likeBuffer(this.SSCS), 0, this.SSAT, this.short_sec_size, d.first_SID,
				d.tot_size, name=qname + " (from SSCS)", null)
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
	this.locate_named_stream = function(qname){
		var d = dirSearch(this,qname.split("/"));
		if (d == null) return [null, 0, 0];
		if (d.tot_size > this.mem_data_len)
			throw new  Error(util.format("%r stream length (%d bytes) > file data size (%d bytes)",
				qname, d.tot_size, this.mem_data_len));
		if (d.tot_size >= this.min_size_std_stream){
			/*
			if this.DEBUG:
				print("\nseen", file=this.logfile)
				dump_list(this.seen, 20, this.logfile)
			*/
			return locateStream(this, wbuf, 512, this.SAT, this.sec_size, d.first_SID, 
				d.tot_size, qname, d.DID+6);
		}else
			return 
			[
				getStream(this, likeBuffer(this.SSCS), 0, this.SSAT, this.short_sec_size, d.first_SID,d.tot_size, qname + " (from SSCS)", null),
				0,
				d.tot_size
			];
	}
}

//return buffer;
function getStream(compDoc, wbuf, base, sat, sec_size, start_sid, size, name, seen_id){
	// print >> self.logfile, "getStream", base, sec_size, start_sid, size
	var sectors = [];
	var s = start_sid;
  var totalSz = 0;
	if(name == null || name == undefined) name = '';
	if (size == null || size == undefined){
		// nothing to check against
		while( s >= 0){
			if(seen_id){
				if (compDoc.seen[s])
					throw new  Error(util.format("%s corruption: seen[%d] == %d" ,name, s, compDoc.seen[s]));
				compDoc.seen[s] = seen_id;
			}
			var start_pos = base + s * sec_size;
			
			wbuf.pull(start_pos,start_pos+sec_size);
			sectors.push(wbuf.toBuf());
		    totalSz += sec_size;
		    if(sat.length <= s) throw new Error(util.format("OLE2 stream %r: sector allocation table invalid entry (%d)",name, s));
		    s = sat[s];
		}
		assert(s == EOCSID, 's == EOCSID');
	}else{
		var todo = size;
		while(s >= 0){
			if(seen_id == null || seen_id == undefined){
				if (compDoc.seen[s])
					throw new  Error(util.format("%s corruption: seen[%d] == %d" % (name, s, compDoc.seen[s])));
				compDoc.seen[s] = seen_id;
			}
			var start_pos = base + s * sec_size;
			var grab = sec_size;
			if (grab > todo) grab = todo;
			todo -= grab;

			wbuf.pull(start_pos,start_pos+grab);
			sectors.push(wbuf.toBuf());
		    totalSz += grab;
		    if(sat.length <= s) throw new Error(util.format)("OLE2 stream %r: sector allocation table invalid entry (%d)",name, s);
		    s = sat[s];
		}
		assert(s == EOCSID, 's == EOCSID');
		if (todo != 0)
			console.log(util.format("WARNING *** OLE2 stream %r: expected size %d, actual size %d\n",name, size, size - todo));
	}
	return Buffer.concat(sectors,totalSz);
}

function dirSearch(compDoc, path, storage_DID){
	// Return matching DirNode instance, or None
	if(storage_DID == null || storage_DID == undefined) storage_DID = 0;  
	var head = path[0];
	var tail = path.slice(1);
	var dl = compDoc.dirlist;
	var length = dl[storage_DID].children.length;
	for(i =0; i < length; i++){
		var child = dl[storage_DID].children[i];
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
	


function locateStream(compDoc, wbuf, base, sat, sec_size, start_sid, expected_stream_size, qname, seen_id){
        // print >> self.logfile, "locateStream", base, sec_size, start_sid, expected_stream_size
	var s = start_sid;
	if (s < 0) throw new  Error(util.format("locateStream: start_sid (%d) is -ve", start_sid));
	var p = -99 ; // dummy previous SID
	var start_pos = -9999;
	var end_pos = -8888;
	var slices = [];
	var tot_found = 0;
	var found_limit = Math.floor((expected_stream_size + sec_size - 1)/ sec_size);
	var totalSz =0;
	while (s >= 0){
		if(compDoc.seen[s]){
			//console.log(util.format("locateStream(%s): seen" , qname); dump_list(self.seen, 20, self.logfile)
			throw new  Error(util.format("%s corruption: seen[%d] == %d",qname, s, compDoc.seen[s]));
		}
		compDoc.seen[s] = seen_id;
		tot_found += 1;
		if (tot_found > found_limit){
			throw new  Error(util.format(
				"%s: size exceeds expected %d bytes; corrupt?"
				, qname, found_limit * sec_size)
				); // Note: expected size rounded up to higher sector
		}
		if( s == p+1)
			// contiguous sectors
			end_pos += sec_size;
		else{
			// start new slice
			if( p >= 0){
				// not first time
				var len = end_pos - start_pos;
				slices.push([start_pos, end_pos]);
				totalSz += len;
			}
			start_pos = base + s * sec_size;
			end_pos = start_pos + sec_size;
		}
		p = s;
		s = sat[s];
	}
	assert(s == EOCSID,'s == EOCSID');
	assert(tot_found == found_limit,'tot_found == found_limit');
	// print >> self.logfile, "locateStream(%s): seen" % qname; dump_list(self.seen, 20, self.logfile)
	if (!slices.length){
		// The stream is contiguous ... just what we like!
		wbuf.pull(start_pos,start_pos+expected_stream_size);
		wbuf.hasData = true;
		return [wbuf, start_pos, expected_stream_size];
	}
	slices.push([start_pos, end_pos]);
	// print >> self.logfile, "+++>>> %d fragments" % len(slices)
	var ret = [];
	slices.forEach(function(x){
		wbuf.pull(x[0], x[1]);
		ret.push(wbuf.toBuf());
	});
	return [Buffer.concat(ret,totalSz), 0, expected_stream_size];
}