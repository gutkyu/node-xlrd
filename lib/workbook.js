'use strict'
var fs = require('fs'),
	util = require('util'),
	assert = require('assert'),
	comm = require('./common'),
	compDoc = require('./compDoc'),
	Sheet = require('./sheet.js'),
	formatting = require('./formatting');

var USE_FANCY_CD = 1;

var _EOF = 0x0F00BAAA;  // not a 16-bit number

var supportedVers = [80, 70, 50, 45, 40, 30, 21, 20];


//openWorkbook(filename, xlType, function(err, workbook))
exports.openWorkbook = function(fd, options, callback){
	var opts = options || {};//todo: default 값으로 병합
	var interObj = {}
	var bk = new Workbook(interObj);

	interObj.loadBiffN(fd);
	
	var biffVer = interObj.getBOF(comm.XL_WORKBOOK_GLOBALS);
	if(!biffVer){
		throw new Error("Can't determine file's BIFF version");
	}
	if(supportedVers.indexOf(biffVer)<0){
		throw new Error("BIFF version "+ comm.biffTextFromNum[biffVer]+"");
	}

	bk.biffVersion = biffVer;
	if(biffVer <= 40){
		//no workbook glabals, only 1 worksheet
		if(opts.onDemand){
			console.log('*** WARNING: on_demand is not supported for this Excel version.');
			console.log('*** Setting on_demand to False.');
			bk.onDemand = opts.onDemand = false;
		}
		interObj.fakeGlobalsGetSheet();
	}else if(biffVer == 45){
		interObj.parseGlobals();
		if(opts.onDemand){
			console.log('*** WARNING: on_demand is not supported for this Excel version.');
			console.log('*** Setting on_demand to False.');
			bk.onDemand = opts.onDemand = false;
		}
	}else{
		interObj.parseGlobals();
		interObj.initSheetList();
		if(!opts.onDemand) interObj.getSheets();
		interObj.setSheetCount(interObj.getSheetListLength());
		if(biffVer == 45 && bk.sheetCount >1){
			console.log('*** WARNING: Excel 4.0 workbook (.XLW) file contains %d worksheets.',bk.sheetCount);
			console.log('*** Book-level data will be that of the last worksheet.\n');
		}
	}
	interObj = null;
	return  bk;
};


function Workbook(interObj){
	var self = this;

    ////
    // The number of worksheets present in the workbook file.
    // This information is available even when no sheets have yet been loaded.
    //nsheets = 0
	var _nsheets = 0;

    ////
    // Which date system was in force when this file was last saved.
    //    0 => 1900 system (the Excel for Windows default).
    //    1 => 1904 system (the Excel for Macintosh default).
    self.dateMode = 0; // In case it's not specified in the file.

    ////
    // Version of BIFF (Binary Interchange File Format) used to create the file.
    // Latest is 8.0 (represented here as 80), introduced with Excel 97.
    // Earliest supported by this module: 2.0 (represented as 20).
    self.biffVersion = 0;

    ////
    // List containing a Name object for each NAME record in the workbook.
    var nameObjList = [];

    ////
    // An integer denoting the character set used for strings in this file.
    // For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.
    // For earlier versions, this is used to derive the appropriate Python encoding
    // to be used to convert to Unicode.
    // Examples: 1252 -> 'cp1252', 10000 -> 'mac_roman'
    self.codePage = null;

    ////
    // The encoding that was derived from the codePage.
    self.encoding = null;

    ////
    // A tuple containing the (telephone system) country code for:  
    //    [0]: the user-interface setting when the file was created.  
    //    [1]: the regional settings.  
    // Example: (1, 61) meaning (USA, Australia).
    // This information may give a clue to the correct encoding for an unknown codePage.
    // For a long list of observed values, refer to the OpenOffice.org documentation for
    // the COUNTRY record.
    self.countries = [0, 0];

    ////
    // What (if anything) is recorded as the name of the last user to save the file.
    var userName = '';
    var userNameIsRaw = false;
    Object.defineProperty(self, 'lastUser', {get:function(){return userNameIsRaw?'':userName;}})
    ////
    // A list of Font class instances, each corresponding to a FONT record.
    //    
    self.fontList = [];

    ////
    // A list of XF class instances, each corresponding to an XF record.
    //    
    self.xfList = [];

    ////
    // A list of Format objects, each corresponding to a FORMAT record, in
    // the order that they appear in the input file.
    // It does  not  contain builtin formats.
    // If you are creating an output file using (for example) pyExcelerator,
    // use this list.
    // The collection to be used for all visual rendering purposes is formatMap.
    //    
    self.formatList = [];

    ////
    // The mapping from XF.formatKey to Format object.
    self.formatMap = {};

    ////
    // This provides access via name to the extended format information for
    // both built-in styles and user-defined styles.  
    // It maps  name  to ( built_in ,  xfIndex ), where:  
    //  name  is either the name of a user-defined style,
    // or the name of one of the built-in styles. Known built-in names are
    // Normal, RowLevel_1 to RowLevel_7,
    // ColLevel_1 to ColLevel_7, Comma, Currency, Percent, "Comma [0]",
    // "Currency [0]", Hyperlink, and "Followed Hyperlink".  
    //  built_in  1 = built-in style, 0 = user-defined  
    //  xfIndex  is an index into Book.xfList.  
    // References: OOo docs s6.99 (STYLE record); Excel UI Format/Style
    // open_workbook(..., formattingInfo=True)
    var styleNameMap = {};

    ////
    // This provides definitions for colour indexes. Please refer to the
    // above section "The Palette; Colour Indexes" for an explanation
    // of how colours are represented in Excel.  
    // Colour indexes into the palette map into (red, green, blue) tuples.
    // "Magic" indexes e.g. 0x7FFF map to None.
    //  colourMap  is what you need if you want to render cells on screen or in a PDF
    // file. If you are writing an output XLS file, use  paletteRecord .
    //    Extracted only if open_workbook(..., formattingInfo=True)
    self.colourMap = {};

    ////
    // If the user has changed any of the colours in the standard palette, the XLS
    // file will contain a PALETTE record with 56 (16 for Excel 4.0 and earlier)
    // RGB values in it, and this list will be e.g. [(r0, b0, g0), ..., (r55, b55, g55)].
    // Otherwise this list will be empty. This is what you need if you are
    // writing an output XLS file. If you want to render cells on screen or in a PDF
    // file, use colourMap.
    //    Extracted only if open_workbook(..., formattingInfo=True)
    var paletteRecord = [];

    ////
    // Time in seconds to extract the XLS image as a contiguous string (or mmap equivalent).
    self.loadTimeStage1 = -1.0;

    ////
    // Time in seconds to parse the data from the contiguous string (or mmap equivalent).
    self.loadTimeStage2 = -1.0;
	
	//#
    // @return A list of all sheets in the book.
    // All sheets not already loaded will be loaded.
	Object.defineProperty(self, 'sheets', {get:function(){
		for (var sheetIdx =0; sheetIdx<_nsheets;sheetIdx++){
            if (! _sheetList[sheetIdx])
                getSheet(sheetIdx);}
        return _sheetList;
	}});
	
	self.sheet = {};
	Object.defineProperty(self.sheet, 'count', {get:function(){return _nsheets;}});

    //#
    // @param sheetIdx Sheet index in range(nsheets)
    // @return An object of the Sheet class
    self.sheet.byIndex=function(sheetIdx){
        return _sheetList[sheetIdx] || getSheet(sheetIdx);
	}
    //#
    // @param sheetName Name of sheet required
    // @return An object of the Sheet class
    self.sheet.byName=function(sheetName){
        
        var sheetIdx = _sheetNames.indexOf(sheetName);
		if(sheetIdx<0) throw new Error(util.format('No sheet named <%s>' , sheetName));
        return self.sheet.byIndex(sheetIdx);
	}
    //#
    // @return A list of the names of all the worksheets in the workbook file.
    // This information is available even when no sheets have yet been loaded.
	Object.defineProperty(self.sheet, 'names', {get:function(){return _sheetNames;}});
	
    //#
    // @param sheetId Name or index of sheet enquired upon
    // @return true if sheet is loaded, false otherwise
    self.sheet.loaded=function(sheetId){
		
        if (!(typeof sheetId == 'number')){
			sheetId = _sheetNames.indexOf(sheetId);
			if(sheetId<0) throw new Error(util.format('No sheet named <%s>' , sheetId));
		}
        return _sheetList[sheetId] ? true:false;
	}
    //#
    // @param sheetId Name or index of sheet to be unloaded.
    self.sheet.unload=function(sheetId){
		if (!(typeof sheetId == 'number')){
			sheetId = _sheetNames.indexOf(sheetId);
			if(sheetId<0) throw new Error(util.format('No sheet named <%s>' , sheetId));
		}
        delete _sheetList[sheetId];
    }
	
	var _sheetNames = [];//_sheet_names
	var _sheetList = [];//_sheet_list
	self._position = 0;
	// Book private variables
	//var _sheet_list = [];
	//var _sheet_names = [];
	self._sheetVisibility = []; // from BOUNDSHEET record
	var _shAbsPosList = [] ;// sheet's absolute position in the stream
	self._sharedStrings = [];
	var _richTextRunlistMap = {};
	var _sheetHdrCount = 0; // BIFF 4W only
	var builtInFmtCount = -1; // unknown as yet. BIFF 3, 4S, 4W
	initialiseFormatFnfo();
	self._allSheetsCount = 0; // includes macro & VBA sheets
	var _supbookCount = 0;
	var _supbookLocalsInx = null;
	var _supbookAddinsInx = null;
	var _allSheetsMap = []; // maps an all_sheets index to a calc-sheets index (or -1)
	var _externSheetInfo = [];
	var _externSheetTypeB57 = [];
	var _externShtNameFromNum = {};
	var _sheetNumFromName = {};
	var _externShtCount = 0;
	var _supbookTypes = [];
	var _resourcesReleased = 0;
	var addinFuncNames = [];


	self.streamLength = 0;
	
	var _unusedBiffVersion ;
	
	var _fd = null;

	if(interObj){
		interObj.getBOF = getBOF;
		interObj.loadBiffN = loadBiffN;
		interObj.parseGlobals = parseGlobals;
		interObj.fakeGlobalsGetSheet = fakeGlobalsGetSheet;
		interObj.setSheetCount = function(count){_nsheets = count;}
		interObj.initSheetList = function(){_sheetList[_sheetNames.length -1] = null;}
		interObj.getSheetListLength = function(){return _sheetList.length;}
		interObj.getSheets = getSheets;
	}

	// This method has a dual purpose. You can call it to release
    // memory-consuming objects when you have finished loading sheets in
    // on_demand mode, but still require the Book object to examine the
    // loaded sheets. It is also called automatically (a) when open_workbook
    // raises an exception. Calling this method multiple times on the 
    // same object has no ill effect.
	
	// release resource
    self.cleanUp = function(){
	    _resourcesReleased = 1;
        delete self.mem;
        _fd = null;
        delete self._sharedStrings;
        _richTextRunlistMap = null;
	}
	
	//biff2 ~ 8
	function loadBiffN(fd){
		var stats = fs.fstatSync(fd);
		if(!stats.size) throw new Error('File size is 0 bytes');
		self.streamLength = stats.size;
		_fd = fd;
		self.base = 0;
		var buf = new Buffer(8);
		fs.readSync(fd,buf,0,8,0);
		var bsig =buf.toString('hex').toLowerCase();
		
		if(bsig== compDoc.SIGNATURE){
			var cd = new compDoc.CompDoc(_fd, self.streamLength);
            if(USE_FANCY_CD){
                if(!['Workbook', 'Book'].some(function(qname){
					var ret = cd.locateNamedStream(qname);
					self.mem = ret[0];
					self.base = ret[1];
					self.streamLength = ret[2];
                    return self.mem && (self.mem.length || self.mem.hasData) ? true: false;
				})) throw new Error("Can't find workbook in OLE2 compound document");
            }else{
				if(!['Workbook', 'Book'].some(function(qname){
					self.mem = cd.getNamedStream(qname);
					return self.mem ? true: false;
				})) throw new Error("Can't find workbook in OLE2 compound document");
                self.streamLength = self.mem.length;
			}
		}
		self._position = self.base;
	}
	
	function initialiseFormatFnfo(){
        // needs to be done once per sheet for BIFF 4W :-(
        self.formatMap = {};
        self.formatList = [];
        self.xfCount = 0;
        self.actualFmtCount = 0; // number of FORMAT records seen so far
        self._xf_index_to_xl_type_map = {0: comm.XL_CELL_NUMBER};
        self._xfEpilogueDone = 0;
        self.xfList = [];
        self.fontList = [];
	}
	
	//return UInt16
	var get2Bytes = function(){
		var pos = self._position;
		var remain = self.mem.length - pos;
		if(remain < 2) { self._position += remain; return _EOF;}
		self._position += 2;
		return self.mem.readUInt16LE(pos);
	};

	function getRecordPartsConditional(reqd_record){
        var pos = self._position;
        var mem = self.mem;
        var code = mem.readUInt16LE(pos), 
            length = mem.readUInt16LE(pos+2);
        if (code != reqd_record)
            return [null, 0, ''];
        pos += 4;
        var data = mem.slice(pos,pos+length);
        self._position = pos + length;
        return [code, length, data];
	}
	//todo : fill this  function.
	//var formatting.initialiseWorkbook = function(){};
	
	function parseGlobals(){
		// DEBUG = 0
		// no need to position, just start reading (after the BOF)
		//todo : uncomment
		formatting.initialiseWorkbook(self);
		//console.log('parse globals',self.constructor.name, self.constructor.prototype);
		while (true){
			var record =self.getRecordParts();
			var rc = record[0], length = record[1], data = record[2];
			//if DEBUG: print("parseGlobals: record code is 0x%04x" % record.rc, file=self.logfile)
			switch(rc){
				case comm.XL_SST:
					handleSST(data);
					break;
				case comm.XL_FONT:
				case comm.XL_FONT_B3B4:
					;//workbook.handle_font(data);
					break;
				case comm.XL_FORMAT:// comm.XL_FORMAT2 is Bif(F <= 3.0, can't appear in globals
					self.handleFormat(data);
					break;
				case comm.XL_XF:
					self.handleXf(data);
					break;
				case comm.XL_BOUNDSHEET:
					handleBoundSheet(data);
					break;
				case comm.XL_DATEMODE:
					handleDateMode(data);
					break;
				case comm.XL_CODEPAGE:
					handleCodePage(data);
					break;
				case comm.XL_COUNTRY:
					handleCountry(data);
					break;
				case comm.XL_EXTERNNAME:
					;//workbook.handle_externname(data);
					break;
				case comm.XL_EXTERNSHEET:
					;//workbook.handle_externsheet(data);
					break;
				case comm.XL_FILEPASS:
					;//workbook.handle_filepass(data);
					break;
				case comm.XL_WRITEACCESS:
					handleWriteAccess(data);
					break;
				case comm.XL_SHEETSOFFSET:
					;//workbook.handle_sheetsoffset(data);
					break;
				case comm.XL_SHEETHDR:
					handleSheetHdr(data);
					break;
				case comm.XL_SUPBOOK:
					;//workbook.handle_supbook(data);
					break;
				case comm.XL_NAME:
					;//workbook.handle_name(data);
					break;
				case comm.XL_PALETTE:				
					;//workbook.handle_palette(data);
					break;
				case comm.XL_STYLE:
					;//workbook.handle_style(data);
					break;
				case comm.XL_EOF:
					self.xfEpilogue();
					//workbook.names_epilogue();
					//workbook.palette_epilogue();
					if( !self.encoding)
						;//workbook.deriveEncoding();
					if( self.biffVersion == 45)
						// DEBUG = 0
						//if( DEBUG) print("global EOF: position", self._position, file=self.logfile);
						console.log("global EOF: position", self._position);
						// if DEBUG:
						//     pos = self._position - 4
						//     print repr(self.mem[pos:pos+40])
					return;
				default:
					if( rc & 0xff == 9 && self.verbosity){
						console.log(util.format("*** Unexpected BOF at posn %d: %d len=%d data=%r"), self._position - length - 4, rc, length, data);
						/*fprintf(self.logfile, "*** Unexpected BOF at posn %d: 0x%04x len=%d data=%r\n",
							self._position - length - 4, record.rc, length, data);
						*/
					}else{
						// if DEBUG:
						//     print >> self.logfile, "parseGlobals: ignoring record code 0x%04x" % record.rc
						;
					}
			}
		}
	}
	//return Buffer;
	function read(pos, length){
		/*
		var buf = new Buffer(length);
		var len = fs.readSync(_fd,buf,0,length,pos);
		self._position = pos + len;
		return length == len ? buf : buf.slice(0,len);
		*/
		var pos = self._position;
		var remain = self.mem.length - pos;
		var len = remain < length ? remain : length;
		self._position += len;
		return self.mem.slice(pos,pos+len);
	}

	self.getRecordParts = function(){
		var pos = self._position;
        var mem = self.mem;
		var code = mem.readUInt16LE(pos) ;
		var length = mem.readUInt16LE(pos + 2);
        pos += 4;
        var data = mem.slice(pos,pos+length);
        self._position = pos + length;
        return [code, length, data];
	}
    
	function getSheet(sheetIndex, updatePos){
		if(updatePos === undefined) updatePos = true;
		if(_resourcesReleased)
			throw new Error("Can't load sheets after releasing resources.");
		if( updatePos) self._position = _shAbsPosList[sheetIndex];
		_unusedBiffVersion = getBOF(comm.XL_WORKSHEET)
		// assert biffVersion == self.biffVersion ### FAILS
		// Have an example where book is v7 but sheet reports v8!!!
		// It appears to work OK if the sheet version is ignored.
		// Confirmed by Daniel Rentz: happens when Excel does "save as"
		// creating an old version file; ignore version details on sheet BOF.
		
		var sh = new Sheet(self, self._position,_sheetNames[sheetIndex],sheetIndex);
		sh.read(self);
		_sheetList[sheetIndex] = sh;
		return sh;
	}
	
	function getSheets(){
        var DEBUG = 0
        
		var sheetIdx = 0;
		_sheetNames.forEach(function(x){
            if (DEBUG) console.log("GET_SHEETS: sheetIdx =", sheetIdx, _sheetNames, _shAbsPosList);
			getSheet(sheetIdx++);
		});
	}
	function fakeGlobalsGetSheet(){// for BIFF 4.0 and earlier
		formatting.initialiseWorkbook(self);
		var fakeSheetName = 'Sheet 1';
		_sheetNames = [fakeSheetName];
		_shAbsPosList = [0];
		self._sheetVisibility = [0]; // one sheet, visible
		_sheetList.push(null); // get_sheet updates _sheet_list but needs a None beforehand
		self.getSheets();
	}
	function handleBoundSheet(data){
        var DEBUG = 0;
        var bv = self.biffVersion;
        deriveEncoding();
        if (DEBUG)
            console.log(util.format("BOUNDSHEET: bv=%d data %r", bv, data));
		var sheetName = '';
		var visibility = 0;
		var sheetType = 0;
		var absPos = 0;
        if (bv == 45){ // BIFF4W
            //////// Not documented in OOo docs ...
            // In fact, the *only* data is the name of the sheet.
            sheetName = unpackString(data, 0, self.encoding,1);
            visibility = 0;
            sheetType = comm.XL_BOUNDSHEET_WORKSHEET; // guess, patch later
            if (!_shAbsPosList.length)
                absPos = self._sheetsoffset + self.base;
                // Note (a) this won't be used
                // (b) it's the position of the SHEETHDR record
                // (c) add 11 to get to the worksheet BOF record
            else
                absPos = -1; // unknown
        }else{
            var offset = data.readInt32LE(0), 
				visibility = data.readUInt8(4), 
				sheetType = data.readUInt8(4);
            absPos = offset + self.base; // because global BOF is always at posn 0 in the stream
            if (bv < comm.BIFF_FIRST_UNICODE)
                sheetName = comm.unpackString(data, 6, self.encoding, 1);
            else
                sheetName = comm.unpackUnicode(data, 6, 1);
		}
        if(DEBUG || self.verbosity >= 2)
            console.log(util.format(
                "BOUNDSHEET: index=%d visibility=%r sheetName=%r absPos=%d sheetType=0x%02x",
                self._allSheetsCount, visibility, sheetName, absPos, sheetType));
        self._allSheetsCount += 1;
        if (sheetType != comm.XL_BOUNDSHEET_WORKSHEET){
            _allSheetsMap.push(-1);
            var descr = {
                1: 'Macro sheet',
                2: 'Chart',
                6: 'Visual Basic module',
            }[sheetType];
			if(!descr ) sheetType ='UNKNOWN';

            if (DEBUG || self.verbosity >= 1)
                console.log(util.format("NOTE *** Ignoring non-worksheet data named %r (type 0x%02x = %s)",
                    sheetName, sheetType, descr));
        }else{
            var snum = _sheetNames.length;
            _allSheetsMap.push(snum);
            _sheetNames.push(sheetName);
            _shAbsPosList.push(absPos);
            self._sheetVisibility.push(visibility);
            _sheetNumFromName[sheetName] = snum;
		}
	}
	function getBOF(rqd_stream){
		// DEBUG = 1
		// if DEBUG: print >> self.logfile, "getbof(): position", self._position
		//if DEBUG: print("reqd: 0x%04x" % rqd_stream, file=self.logfile)
		function bofError(msg){
			throw new Error('Unsupported format, or corrupt file: ' + msg);
		}
		var savPos = self._position;
		var opcode = get2Bytes();
		if( opcode == _EOF)
			bofError('Expected BOF record; met end of file');
		if(comm.bofCodes.indexOf(opcode)<0){
			var buf = new Buffer(8);
			fs.readSync(_fd,buf,0,8,savPos);
			bofError('Expected BOF record; found %r', buf.toString('hex') );
		}
		var length = get2Bytes();
		if( length == _EOF)
			bofError('Incomplete BOF record[1]; met end of file')
		if( ! (4 <= length &&  length <= 20)){
			var opcd = new Buffer(4);
			opcd.writeUInt32LE(opcode,0);
			bofError(util.format('Invalid length (%d) for BOF record type 0x%s', length, opcd.toString('hex')));
		}
		var padding = new Buffer(Math.max(0, comm.bofLen[opcode] - length));
		padding.fill(0);
		var data = read(self._position, length);//Type Buffer
		//if DEBUG: fprintf(self.logfile, "\ngetbof(): data=%r\n", data);
		//if DEBUG: fprintf(self.logfile, "\ngetbof(): data=%r\n", data);
		if( data.length < length) bofError('Incomplete BOF record[2]; met end of file');
		data = Buffer.concat([data,padding]);
		var version1 = opcode >> 8;
		var version2 = data.readUInt16LE(0);
		var streamType = data.readUInt16LE(2);
		/*if DEBUG:
			print("getbof(): op=0x%04x version2=0x%04x streamType=0x%04x" \
				% (opcode, version2, streamType), file=self.logfile)
		*/
		var bofOffset = self._position - 4 - length;
		/*if DEBUG:
			print("getbof(): BOF found at offset %d; savPos=%d" \
				% (bofOffset, savPos), file=self.logfile)
		*/
		var version =0, build =0, year = 0;
		if( version1 == 0x08){
			build = data.readUInt16LE(4);
			year = data.readUInt16LE(6);
			if( version2 == 0x0600) version = 80;
			else if (version2 == 0x0500)
				if( year < 1994 ||   [2412, 3218, 3321].indexOf(build) >= 0 ) version = 50;
				else version = 70;
			else
				// dodgy one, created by a 3rd-party tool
				try{version = {0x0000: 21,0x0007: 21,0x0200: 21,0x0300: 30,0x0400: 40}[version2];}catch(e){}
		}else if( [0x04, 0x02, 0x00].indexOf(version1) >= 0){
			version = {0x04: 40, 0x02: 30, 0x00: 21}[version1];
		}
		if(version == 40 && streamType == comm.XL_WORKBOOK_GLOBALS_4W)  version = 45; // i.e. 4W

		/*if DEBUG or self.verbosity >= 2:
			print("BOF: op=0x%04x vers=0x%04x stream=0x%04x buildid=%d buildyr=%d -> BIFF%d" \
				% (opcode, version2, streamType, build, year, version), file=self.logfile)
		*/
		var gotGlobals = streamType == comm.XL_WORKBOOK_GLOBALS || 
							(version == 45 && streamType == comm.XL_WORKBOOK_GLOBALS_4W);
		if ((rqd_stream == comm.XL_WORKBOOK_GLOBALS && gotGlobals) || streamType == rqd_stream) return version;
		if( version < 50 && streamType == comm.XL_WORKSHEET) return version;
		if (version >= 50 && streamType == 0x0100) bofError("Workspace file -- no spreadsheet data");
		bofError(
			util.format('BOF not workbook/worksheet: op=%d vers=%d strm=%d build=%d year=%d -> BIFF%d',
			opcode, version2, streamType, build, year, version)
		);
	}
	
	function deriveEncoding(){
        if(self.encodingOverride)
            self.encoding = self.encodingOverride
        else if(!self.codePage){
            if (self.biffVersion < 80)
                //fprintf(self.logfile,"*** No CODEPAGE record, no encodingOverride: will use 'ascii'\n")
                self.encoding = 'ascii';
            else
                self.codePage = 1200; // utf16le
                if (self.verbosity >= 2)
                    //fprintf(self.logfile, "*** No CODEPAGE record; assuming 1200 (utf_16_le)\n")
					;
        }else{
            var codePage = self.codePage,
				encoding = self.encoding;
            if (Object.keys(comm.encodingFromCodePage).indexOf(codePage.toString())>=0) 
				encoding = comm.encodingFromCodePage[codePage];
            else if(300 <= codePage && codePage <= 1999)
                encoding = 'cp' + codePage;
            else
                encoding = 'unknown_codepage_' + codePage;
            //if DEBUG or (self.verbosity and encoding != self.encoding) 
            //    fprintf(self.logfile, "CODEPAGE: codePage %r -> encoding %r\n", codePage, encoding)
            self.encoding = encoding;
		}
        if (self.codePage != 1200) // utf_16_le
			//todo iconv로 구현할 것
			throw new Error('codePage',self.codePage,'unsupported now');
			/*
            // If we don't have a codec that can decode ASCII into Unicode,
            // we're well & truly stuffed -- let the punter know ASAP.
            try:
                _unused = unicode(b'trial', self.encoding)
            except BaseException as e:
                fprintf(self.logfile,
                    "ERROR *** codePage %r -> encoding %r -> %s: %s\n",
                    self.codePage, self.encoding, type(e).__name__.split(".")[-1], e)
                raise
			*/
        if (userNameIsRaw){
			//todo iconv로 구현할 것
			throw new Error('codePage(',self.codePage,') unsupported now');
			
            var str = comm.encode(userName, 0, self.encoding, 1)
            var str = str.rtrim();
            // if DEBUG:
            //     print "CODEPAGE: user name decoded from %r to %r" % (self.userName, str)
            userName = str;
            userNameIsRaw = false;
			
		}
        return self.encoding;
	}
    function handleCodePage(data){
        // DEBUG = 0
        var codePage = data.readUInt16LE(0);
        self.codePage = codePage;
        deriveEncoding();
	}
    function handleCountry(data){
        var countries = [data.readUInt16LE(0),data.readUInt16LE(2)];
        if (self.verbosity) console.log("Countries:", countries[0],',',countries[1]);
        // Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        assert((self.countries[0] ==0 && self.countries[1] == 0) || 
				(self.countries[0] ==countries[0] && self.countries[1] == countries[1]),
			'self.countries == [0, 0] or self.countries == countries'
		);
        self.countries = countries;
	}
    function handleDateMode(data){
        var dateMode = data.readUInt16LE(0);
		/*
        if DEBUG or self.verbosity:
            fprintf(self.logfile, "DATEMODE: dateMode %r\n", dateMode)
		*/
        assert((dateMode == 0 || dateMode == 1),'dateMode in (0, 1)');
        self.dateMode = dateMode;
	}
	function handleSheetHdr(data){
        // This a BIFF 4W special.
        // The SHEETHDR record is followed by a (BOF ... EOF) substream containing
        // a worksheet.
        // DEBUG = 1
        deriveEncoding();
        var sheetLen = data.readInt32LE(0);
        var sheetName = comm.encode(data, 4, self.encoding,1)
        var sheetno = _sheetHdrCount;
        assert(sheetName == _sheetNames[sheetno],'sheetName == _sheetNames[sheetno]');
        _sheetHdrCount += 1;
        var bofPos = self._position;
        var posn = bofPos - 4 - data.length;
        //if DEBUG: fprintf(self.logfile, 'SHEETHDR %d at posn %d: len=%d name=%r\n', sheetno, posn, sheetLen, sheetName)
		console.log(util.format('SHEETHDR %d at posn %d: len=%d name=%s\n', sheetno, posn, sheetLen, sheetName));
        initialiseFormatFnfo();
        //if DEBUG: print('SHEETHDR: xf epilogue flag is %d' % self._xfEpilogueDone, file=self.logfile)
        _sheetList.push(null); // get_sheet updates _sheet_list but needs a None beforehand
        getSheet(sheetno, false)
        //if DEBUG: print('SHEETHDR: posn after get_sheet() =', self._position, file=self.logfile)
        self._position = bofPos + sheetLen;
	}
	
	function handleSST(data){
		// DEBUG = 1
		//if DEBUG:
		//	print("SST Processing", file=self.logfile)
		//	t0 = time.time()
		var nbt = data.length;
		var strList = [data];
		var uniqueStrings = data.readInt32LE(4);
		//if DEBUG  or self.verbosity >= 2:
		//	fprintf(self.logfile, "SST: unique strings: %d\n", uniqueStrings)
		while(1){
			var rc = getRecordPartsConditional(comm.XL_CONTINUE);
			var code = rc[0], nb = rc[1], data = rc[2];
			if (code == null)
				break;
			nbt += nb;
			//if DEBUG >= 2:
			//	fprintf(self.logfile, "CONTINUE: adding %d bytes to SST -> %d\n", nb, nbt)
			strList.push(data);
		}
		var sstt= parseSST(strList, uniqueStrings)
		self._sharedStrings = sstt[0];
		var rtRunList = sstt[1];
		if (self.formattingInfo)
			self._richTextRunlistMap = rtRunList;
		//if DEBUG:
		//	t1 = time.time()
		//	print("SST processing took %.2f seconds" % (t1 - t0, ), file=self.logfile)
	}
    
    function handleWriteAccess(data){
        var DEBUG = 0;
        var str = '';
        if (self.biffVersion < 80){
            if (! self.encoding){
                userNameIsRaw = true;
                userName = data
                return;
            }
            str = comm.unpackString(data, 0, self.encoding, 1)
        }else{
            str = comm.unpackUnicode(data, 0, 2)
        }
        if (DEBUG)  
            console.log("WRITEACCESS: %d bytes; raw=%s %r\n", data.length, userNameIsRaw, str)
        userName = str.trimRight()
    }
}

function parseSST(dataTable, stringLength){
    //"Return list of strings"
    var dataIdx = 0;
    var dataTabLen = dataTable.length;
    var data = dataTable[0];
    var dataLen = data.length;
    var pos = 8;
    var strings = [];
    var richtextRuns = {};
    var latin1 = "latin1";
    for(var _unused_i =0;_unused_i < stringLength; _unused_i++){
        var nchars =  data.readUInt16LE(pos);
        pos += 2;
        var options = data[pos];
        pos += 1;
        var rtCount = 0;
        var phoSize = 0;
        if (options & 0x08){ // richtext
            rtCount = data.readUInt16LE(pos);
            pos += 2;
		}
        if (options & 0x04){ // phonetic
            phoSize = data.readInt32LE(pos);
            pos += 4;
		}
        var accStr ='';
        var charsGot = 0;
        while(1){
            var charsNeed = nchars - charsGot;
			var charsAvail = null;
			var rawStr= null;
            if (options & 0x01){
                // Uncompressed UTF-16
                charsAvail = Math.min((dataLen - pos) >> 1, charsNeed);
                rawStr = data.slice(pos,pos+2*charsAvail);
                // if DEBUG: print "SST U16: nchars=%d pos=%d rawStr=%r" % (nchars, pos, rawStr)
                try{
                    accStr += comm.decode(rawStr,0,rawStr.length, "utf16le");
                }catch(e){
                    // print "SST U16: nchars=%d pos=%d rawStr=%r" % (nchars, pos, rawStr)
                    // Probable cause: dodgy data e.g. unfinished surrogate pair.
                    // E.g. file unicode2.xls in pyExcelerator's examples has cells containing
                    // unichr(i) for i in range(0x100000)
                    // so this will include 0xD800 etc
                    throw e;
				}
                pos += 2*charsAvail;
            }else{
                // Note: this is COMPRESSED (not ASCII!) encoding!!!
                charsAvail = Math.min(dataLen - pos, charsNeed);
                rawStr = data.slice(pos,pos+charsAvail);
                // if DEBUG: print "SST CMPRSD: nchars=%d pos=%d rawStr=%r" % (nchars, pos, rawStr)
                accStr += comm.decode(rawStr,0,rawStr.length, latin1);
                pos += charsAvail;
			}
            charsGot += charsAvail;
            if (charsGot == nchars)
                break;
            dataIdx += 1;
            data = dataTable[dataIdx];
            dataLen = data.length;
            options = data[0];
            pos = 1;
        }
        if(rtCount){
            var runs = [];
            for(var runindex =0; runindex <rtCount; runindex++){
                if (pos == dataLen){
                    pos = 0;
                    dataIdx += 1;
                    data = dataTable[dataIdx];
                    dataLen = data.length;
				}
                runs.push(data.readUInt16LE(pos));
				runs.push(data.readUInt16LE(pos+2));
                pos += 4;
			}
            richtextRuns[strings.length] = runs;
        }        
        pos += phoSize; // size of the phonetic stuff to skip
        if(pos >= dataLen){
            // adjust to correct position in next record
            pos = pos - dataLen;
            dataIdx += 1;
            if(dataIdx < dataTabLen){
                data = dataTable[dataIdx];
                dataLen = data.length;
            }else
                assert (_unused_i == stringLength - 1, '_unused_i == stringLength - 1;');
		}
        strings.push(accStr);
	}
    return [strings, richtextRuns];
}
