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
		interObj.parse_globals();
		if(opts.onDemand){
			console.log('*** WARNING: on_demand is not supported for this Excel version.');
			console.log('*** Setting on_demand to False.');
			bk.onDemand = opts.onDemand = false;
		}
	}else{
		interObj.parse_globals();
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
    var name_obj_list = [];

    ////
    // An integer denoting the character set used for strings in this file.
    // For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.
    // For earlier versions, this is used to derive the appropriate Python encoding
    // to be used to convert to Unicode.
    // Examples: 1252 -> 'cp1252', 10000 -> 'mac_roman'
    self.codePage = null;

    ////
    // The encoding that was derived from the codepage.
    self.encoding = null;

    ////
    // A tuple containing the (telephone system) country code for:  
    //    [0]: the user-interface setting when the file was created.  
    //    [1]: the regional settings.  
    // Example: (1, 61) meaning (USA, Australia).
    // This information may give a clue to the correct encoding for an unknown codepage.
    // For a long list of observed values, refer to the OpenOffice.org documentation for
    // the COUNTRY record.
    self.countries = [0, 0];

    ////
    // What (if anything) is recorded as the name of the last user to save the file.
    var user_name = '';

    ////
    // A list of Font class instances, each corresponding to a FONT record.
    //    
    self.font_list = [];

    ////
    // A list of XF class instances, each corresponding to an XF record.
    //    
    self.xf_list = [];

    ////
    // A list of Format objects, each corresponding to a FORMAT record, in
    // the order that they appear in the input file.
    // It does  not  contain builtin formats.
    // If you are creating an output file using (for example) pyExcelerator,
    // use this list.
    // The collection to be used for all visual rendering purposes is format_map.
    //    
    self.format_list = [];

    ////
    // The mapping from XF.format_key to Format object.
    self.format_map = {};

    ////
    // This provides access via name to the extended format information for
    // both built-in styles and user-defined styles.  
    // It maps  name  to ( built_in ,  xf_index ), where:  
    //  name  is either the name of a user-defined style,
    // or the name of one of the built-in styles. Known built-in names are
    // Normal, RowLevel_1 to RowLevel_7,
    // ColLevel_1 to ColLevel_7, Comma, Currency, Percent, "Comma [0]",
    // "Currency [0]", Hyperlink, and "Followed Hyperlink".  
    //  built_in  1 = built-in style, 0 = user-defined  
    //  xf_index  is an index into Book.xf_list.  
    // References: OOo docs s6.99 (STYLE record); Excel UI Format/Style
    // open_workbook(..., formattingInfo=True)
    var style_name_map = {};

    ////
    // This provides definitions for colour indexes. Please refer to the
    // above section "The Palette; Colour Indexes" for an explanation
    // of how colours are represented in Excel.  
    // Colour indexes into the palette map into (red, green, blue) tuples.
    // "Magic" indexes e.g. 0x7FFF map to None.
    //  colour_map  is what you need if you want to render cells on screen or in a PDF
    // file. If you are writing an output XLS file, use  palette_record .
    //    Extracted only if open_workbook(..., formattingInfo=True)
    var colour_map = {};

    ////
    // If the user has changed any of the colours in the standard palette, the XLS
    // file will contain a PALETTE record with 56 (16 for Excel 4.0 and earlier)
    // RGB values in it, and this list will be e.g. [(r0, b0, g0), ..., (r55, b55, g55)].
    // Otherwise this list will be empty. This is what you need if you are
    // writing an output XLS file. If you want to render cells on screen or in a PDF
    // file, use colour_map.
    //    Extracted only if open_workbook(..., formattingInfo=True)
    var palette_record = [];

    ////
    // Time in seconds to extract the XLS image as a contiguous string (or mmap equivalent).
    var load_time_stage_1 = -1.0;

    ////
    // Time in seconds to parse the data from the contiguous string (or mmap equivalent).
    var load_time_stage_2 = -1.0;
	
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
    // @param sheet_name Name of sheet required
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
	self._sheet_visibility = []; // from BOUNDSHEET record
	var _sh_abs_posn = [] ;// sheet's absolute position in the stream
	self._sharedstrings = [];
	var _rich_text_runlist_map = {};
	var raw_user_name = false;
	var _sheethdr_count = 0; // BIFF 4W only
	var builtinfmtcount = -1; // unknown as yet. BIFF 3, 4S, 4W
	initialise_format_info();
	self._all_sheets_count = 0; // includes macro & VBA sheets
	var _supbook_count = 0;
	var _supbook_locals_inx = null;
	var _supbook_addins_inx = null;
	var _all_sheets_map = []; // maps an all_sheets index to a calc-sheets index (or -1)
	var _externsheet_info = [];
	var _externsheet_type_b57 = [];
	var _extnsht_name_from_num = {};
	var _sheet_num_from_name = {};
	var _extnsht_count = 0;
	var _supbook_types = [];
	var _resources_released = 0;
	var addin_func_names = [];


	self.stream_len = 0;
	
	var _unused_biff_version ;
	
	var _fd = null;

	if(interObj){
		interObj.getBOF = getBOF;
		interObj.loadBiffN = loadBiffN;
		interObj.parse_globals = parse_globals;
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
	    _resources_released = 1;
        delete self.mem;
        _fd = null;
        delete self._sharedstrings;
        _rich_text_runlist_map = null;
	}
	
	//biff2 ~ 8
	function loadBiffN(fd){
		var stats = fs.fstatSync(fd);
		if(!stats.size) throw new Error('File size is 0 bytes');
		self.stream_len = stats.size;
		_fd = fd;
		self.base = 0;
		var buf = new Buffer(8);
		fs.readSync(fd,buf,0,8,0);
		var bsig =buf.toString('hex').toLowerCase();
		
		if(bsig== compDoc.SIGNATURE){
			var cd = new compDoc.CompDoc(_fd, self.stream_len);
            if(USE_FANCY_CD){
                if(!['Workbook', 'Book'].some(function(qname){
					var ret = cd.locateNamedStream(qname);
					self.mem = ret[0];
					self.base = ret[1];
					self.stream_len = ret[2];
                    return self.mem && (self.mem.length || self.mem.hasData) ? true: false;
				})) throw new Error("Can't find workbook in OLE2 compound document");
            }else{
				if(!['Workbook', 'Book'].some(function(qname){
					self.mem = cd.getNamedStream(qname);
					return self.mem ? true: false;
				})) throw new Error("Can't find workbook in OLE2 compound document");
                self.stream_len = self.mem.length;
			}
		}
		self._position = self.base;
	}
	
	function initialise_format_info(){
        // needs to be done once per sheet for BIFF 4W :-(
        self.format_map = {};
        self.format_list = [];
        self.xfcount = 0;
        self.actualfmtcount = 0; // number of FORMAT records seen so far
        self._xf_index_to_xl_type_map = {0: comm.XL_CELL_NUMBER};
        self._xf_epilogue_done = 0;
        self.xf_list = [];
        self.font_list = [];
	}
	
	//return UInt16
	var get2bytes = function(){
		var pos = self._position;
		var remain = self.mem.length - pos;
		if(remain < 2) { self._position += remain; return _EOF;}
		self._position += 2;
		return self.mem.readUInt16LE(pos);
	};

	function get_record_parts_conditional(reqd_record){
        var pos = self._position;
        var mem = self.mem;
        var code = data.readUInt16LE(pos), 
			length = data.readUInt16LE(pos+2);
        if (code != reqd_record)
            return [null, 0, ''];
        pos += 4;
        var data = mem.slice(pos,pos+length);
        self._position = pos + length
        return [code, length, data];
	}
	//todo : fill this  function.
	//var formatting.initialise_book = function(){};
	
	function parse_globals(){
		// DEBUG = 0
		// no need to position, just start reading (after the BOF)
		//todo : uncomment
		formatting.initialise_book(self);
		//console.log('parse globals',self.constructor.name, self.constructor.prototype);
		while (true){
			var record =self.get_record_parts();
			var rc = record[0], length = record[1], data = record[2];
			//if DEBUG: print("parse_globals: record code is 0x%04x" % record.rc, file=self.logfile)
			switch(rc){
				case comm.XL_SST:
					handle_sst(data);
					break;
				case comm.XL_FONT:
				case comm.XL_FONT_B3B4:
					;//workbook.handle_font(data);
					break;
				case comm.XL_FORMAT:// comm.XL_FORMAT2 is Bif(F <= 3.0, can't appear in globals
					self.handle_format(data);
					break;
				case comm.XL_XF:
					self.handle_xf(data);
					break;
				case comm.XL_BOUNDSHEET:
					handle_boundsheet(data);
					break;
				case comm.XL_DATEMODE:
					handle_datemode(data);
					break;
				case comm.XL_CODEPAGE:
					handle_codepage(data);
					break;
				case comm.XL_COUNTRY:
					handle_country(data);
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
					;//workbook.handle_writeaccess(data);
					break;
				case comm.XL_SHEETSOFFSET:
					;//workbook.handle_sheetsoffset(data);
					break;
				case comm.XL_SHEETHDR:
					handle_sheethdr(data);
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
					self.xf_epilogue();
					//workbook.names_epilogue();
					//workbook.palette_epilogue();
					if( !self.encoding)
						;//workbook.derive_encoding();
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
						//     print >> self.logfile, "parse_globals: ignoring record code 0x%04x" % record.rc
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

	self.get_record_parts = function(){
		var pos = self._position;
        var mem = self.mem;
		var code = mem.readUInt16LE(pos) ;
		var length = mem.readUInt16LE(pos + 2);
        pos += 4;
        var data = mem.slice(pos,pos+length);
        self._position = pos + length;
        return [code, length, data];
	}
    function get_record_parts_conditional(reqd_record){
        var pos = self._position;
        var mem = self.mem;
        var code = mem.readUInt16LE(pos), length = mem.readUInt16LE(pos+2);
        if (code != reqd_record)
            return [null, 0, ''];
        pos += 4;
        var data = mem.slice(pos,pos+length);
        self._position = pos + length;
        return [code, length, data];
	}
	
	function getSheet(sh_number, update_pos){
		if(update_pos === undefined) update_pos = true;
		if(_resources_released)
			throw new Error("Can't load sheets after releasing resources.");
		if( update_pos) self._position = _sh_abs_posn[sh_number];
		_unused_biff_version = getBOF(comm.XL_WORKSHEET)
		// assert biffVersion == self.biffVersion ### FAILS
		// Have an example where book is v7 but sheet reports v8!!!
		// It appears to work OK if the sheet version is ignored.
		// Confirmed by Daniel Rentz: happens when Excel does "save as"
		// creating an old version file; ignore version details on sheet BOF.
		
		var sh = new Sheet(self, self._position,_sheetNames[sh_number],sh_number);
		sh.read(self);
		_sheetList[sh_number] = sh;
		return sh;
	}
	
	function getSheets(){
        var DEBUG = 0
        
		var sheetno = 0;
		_sheetNames.forEach(function(x){
            if (DEBUG) console.log("GET_SHEETS: sheetno =", sheetno, _sheetNames, _sh_abs_posn);
			getSheet(sheetno++);
		});
	}
	function fakeGlobalsGetSheet(){// for BIFF 4.0 and earlier
		formatting.initialise_book(self);
		var fake_sheet_name = 'Sheet 1';
		_sheetNames = [fake_sheet_name];
		_sh_abs_posn = [0];
		self._sheet_visibility = [0]; // one sheet, visible
		_sheetList.push(null); // get_sheet updates _sheet_list but needs a None beforehand
		self.getSheets();
	}
	function handle_boundsheet(data){
        var DEBUG = 0;
        var bv = self.biffVersion;
        derive_encoding();
        if (DEBUG)
            console.log(util.format("BOUNDSHEET: bv=%d data %r", bv, data));
		var sheet_name = '';
		var visibility = 0;
		var sheet_type = 0;
		var abs_posn = 0;
        if (bv == 45){ // BIFF4W
            //////// Not documented in OOo docs ...
            // In fact, the *only* data is the name of the sheet.
            sheet_name = unpack_string(data, 0, self.encoding,1);
            visibility = 0;
            sheet_type = comm.XL_BOUNDSHEET_WORKSHEET; // guess, patch later
            if (!_sh_abs_posn.length)
                abs_posn = self._sheetsoffset + self.base;
                // Note (a) this won't be used
                // (b) it's the position of the SHEETHDR record
                // (c) add 11 to get to the worksheet BOF record
            else
                abs_posn = -1; // unknown
        }else{
            var offset = data.readInt32LE(0), 
				visibility = data.readUInt8(4), 
				sheet_type = data.readUInt8(4);
            abs_posn = offset + self.base; // because global BOF is always at posn 0 in the stream
            if (bv < comm.BIFF_FIRST_UNICODE)
                sheet_name = comm.unpack_string(data, 6, self.encoding, 1);
            else
                sheet_name = comm.unpack_unicode(data, 6, 1);
		}
        if(DEBUG || self.verbosity >= 2)
            console.log(util.format(
                "BOUNDSHEET: inx=%d vis=%r sheet_name=%r abs_posn=%d sheet_type=0x%02x",
                self._all_sheets_count, visibility, sheet_name, abs_posn, sheet_type));
        self._all_sheets_count += 1;
        if (sheet_type != comm.XL_BOUNDSHEET_WORKSHEET){
            _all_sheets_map.push(-1);
            var descr = {
                1: 'Macro sheet',
                2: 'Chart',
                6: 'Visual Basic module',
            }[sheet_type];
			if(!descr ) sheet_type ='UNKNOWN';

            if (DEBUG || self.verbosity >= 1)
                console.log(util.format("NOTE *** Ignoring non-worksheet data named %r (type 0x%02x = %s)",
                    sheet_name, sheet_type, descr));
        }else{
            var snum = _sheetNames.length;
            _all_sheets_map.push(snum);
            _sheetNames.push(sheet_name);
            _sh_abs_posn.push(abs_posn);
            self._sheet_visibility.push(visibility);
            _sheet_num_from_name[sheet_name] = snum;
		}
	}
	function getBOF(rqd_stream){
		// DEBUG = 1
		// if DEBUG: print >> self.logfile, "getbof(): position", self._position
		//if DEBUG: print("reqd: 0x%04x" % rqd_stream, file=self.logfile)
		function bofError(msg){
			throw new Error('Unsupported format, or corrupt file: ' + msg);
		}
		var savpos = self._position;
		var opcode = get2bytes();
		if( opcode == _EOF)
			bofError('Expected BOF record; met end of file');
		if(comm.bofcodes.indexOf(opcode)<0){
			var buf = new Buffer(8);
			fs.readSync(_fd,buf,0,8,savpos);
			bofError('Expected BOF record; found %r', buf.toString('hex') );
		}
		var length = get2bytes();
		if( length == _EOF)
			bofError('Incomplete BOF record[1]; met end of file')
		if( ! (4 <= length &&  length <= 20)){
			var opcd = new Buffer(4);
			opcd.writeUInt32LE(opcode,0);
			bofError(util.format('Invalid length (%d) for BOF record type 0x%s', length, opcd.toString('hex')));
		}
		var padding = new Buffer(Math.max(0, comm.boflen[opcode] - length));
		padding.fill(0);
		var data = read(self._position, length);//Type Buffer
		//if DEBUG: fprintf(self.logfile, "\ngetbof(): data=%r\n", data);
		//if DEBUG: fprintf(self.logfile, "\ngetbof(): data=%r\n", data);
		if( data.length < length) bofError('Incomplete BOF record[2]; met end of file');
		data = Buffer.concat([data,padding]);
		var version1 = opcode >> 8;
		var version2 = data.readUInt16LE(0);
		var streamtype = data.readUInt16LE(2);
		/*if DEBUG:
			print("getbof(): op=0x%04x version2=0x%04x streamtype=0x%04x" \
				% (opcode, version2, streamtype), file=self.logfile)
		*/
		var bof_offset = self._position - 4 - length;
		/*if DEBUG:
			print("getbof(): BOF found at offset %d; savpos=%d" \
				% (bof_offset, savpos), file=self.logfile)
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
		if(version == 40 && streamtype == comm.XL_WORKBOOK_GLOBALS_4W)  version = 45; // i.e. 4W

		/*if DEBUG or self.verbosity >= 2:
			print("BOF: op=0x%04x vers=0x%04x stream=0x%04x buildid=%d buildyr=%d -> BIFF%d" \
				% (opcode, version2, streamtype, build, year, version), file=self.logfile)
		*/
		var got_globals = streamtype == comm.XL_WORKBOOK_GLOBALS || 
							(version == 45 && streamtype == comm.XL_WORKBOOK_GLOBALS_4W);
		if ((rqd_stream == comm.XL_WORKBOOK_GLOBALS && got_globals) || streamtype == rqd_stream) return version;
		if( version < 50 && streamtype == comm.XL_WORKSHEET) return version;
		if (version >= 50 && streamtype == 0x0100) bofError("Workspace file -- no spreadsheet data");
		bofError(
			util.format('BOF not workbook/worksheet: op=%d vers=%d strm=%d build=%d year=%d -> BIFF%d',
			opcode, version2, streamtype, build, year, version)
		);
	}
	
	function derive_encoding(){
        if(self.encoding_override)
            self.encoding = self.encoding_override
        else if(!self.codePage){
            if (self.biffVersion < 80)
                //fprintf(self.logfile,"*** No CODEPAGE record, no encoding_override: will use 'ascii'\n")
                self.encoding = 'ascii';
            else
                self.codePage = 1200; // utf16le
                if (self.verbosity >= 2)
                    //fprintf(self.logfile, "*** No CODEPAGE record; assuming 1200 (utf_16_le)\n")
					;
        }else{
            var codePage = self.codePage,
				encoding = self.encoding;
            if (Object.keys(comm.encoding_from_codepage).indexOf(codePage.toString())>=0) 
				encoding = comm.encoding_from_codepage[codePage];
            else if(300 <= codePage && codePage <= 1999)
                encoding = 'cp' + codePage;
            else
                encoding = 'unknown_codepage_' + codePage;
            //if DEBUG or (self.verbosity and encoding != self.encoding) 
            //    fprintf(self.logfile, "CODEPAGE: codepage %r -> encoding %r\n", codePage, encoding)
            self.encoding = encoding;
		}
        if (self.codePage != 1200) // utf_16_le
			//todo iconv로 구현할 것
			throw new Error('codepage',self.codePage,'unsupported now');
			/*
            // If we don't have a codec that can decode ASCII into Unicode,
            // we're well & truly stuffed -- let the punter know ASAP.
            try:
                _unused = unicode(b'trial', self.encoding)
            except BaseException as e:
                fprintf(self.logfile,
                    "ERROR *** codepage %r -> encoding %r -> %s: %s\n",
                    self.codepage, self.encoding, type(e).__name__.split(".")[-1], e)
                raise
			*/
        if (raw_user_name){
			//todo iconv로 구현할 것
			throw new Error('codepage(',self.codePage,') unsupported now');
			
            var strg = comm.encode(user_name, 0, self.encoding, 1)
            var strg = strg.rtrim();
            // if DEBUG:
            //     print "CODEPAGE: user name decoded from %r to %r" % (self.user_name, strg)
            self.user_name = strg;
            self.raw_user_name = false;
			
		}
        return self.encoding;
	}
    function handle_codepage(data){
        // DEBUG = 0
        var codePage = data.readUInt16LE(0);
        self.codePage = codePage;
        derive_encoding();
	}
    function handle_country(data){
        var countries = [data.readUInt16LE(0),data.readUInt16LE(2)];
        if (self.verbosity) console.log("Countries:", countries[0],',',countries[1]);
        // Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        assert((self.countries[0] ==0 && self.countries[1] == 0) || 
				(self.countries[0] ==countries[0] && self.countries[1] == countries[1]),
			'self.countries == [0, 0] or self.countries == countries'
		);
        self.countries = countries;
	}
    function handle_datemode(data){
        var dateMode = data.readUInt16LE(0);
		/*
        if DEBUG or self.verbosity:
            fprintf(self.logfile, "DATEMODE: dateMode %r\n", dateMode)
		*/
        assert((dateMode == 0 || dateMode == 1),'dateMode in (0, 1)');
        self.dateMode = dateMode;
	}
	function handle_sheethdr(data){
        // This a BIFF 4W special.
        // The SHEETHDR record is followed by a (BOF ... EOF) substream containing
        // a worksheet.
        // DEBUG = 1
        derive_encoding();
        var sheet_len = data.readInt32LE(0);
        var sheet_name = comm.encode(data, 4, self.encoding,1)
        var sheetno = _sheethdr_count;
        assert(sheet_name == _sheetNames[sheetno],'sheet_name == _sheetNames[sheetno]');
        _sheethdr_count += 1;
        var BOF_posn = self._position;
        var posn = BOF_posn - 4 - data.length;
        //if DEBUG: fprintf(self.logfile, 'SHEETHDR %d at posn %d: len=%d name=%r\n', sheetno, posn, sheet_len, sheet_name)
		console.log(util.format('SHEETHDR %d at posn %d: len=%d name=%s\n', sheetno, posn, sheet_len, sheet_name));
        initialise_format_info();
        //if DEBUG: print('SHEETHDR: xf epilogue flag is %d' % self._xf_epilogue_done, file=self.logfile)
        _sheetList.push(null); // get_sheet updates _sheet_list but needs a None beforehand
        getSheet(sheetno, false)
        //if DEBUG: print('SHEETHDR: posn after get_sheet() =', self._position, file=self.logfile)
        self._position = BOF_posn + sheet_len;
	}
	
	function handle_sst(data){
		// DEBUG = 1
		//if DEBUG:
		//	print("SST Processing", file=self.logfile)
		//	t0 = time.time()
		var nbt = data.length;
		var strlist = [data];
		var uniquestrings = data.readInt32LE(4);
		//if DEBUG  or self.verbosity >= 2:
		//	fprintf(self.logfile, "SST: unique strings: %d\n", uniquestrings)
		while(1){
			var rc = get_record_parts_conditional(comm.XL_CONTINUE);
			var code = rc[0], nb = rc[1], data = rc[2];
			if (code == null)
				break;
			nbt += nb;
			//if DEBUG >= 2:
			//	fprintf(self.logfile, "CONTINUE: adding %d bytes to SST -> %d\n", nb, nbt)
			strlist.push(data);
		}
		var sstt= parseSST(strlist, uniquestrings)
		self._sharedstrings = sstt[0];
		var rt_runlist = sstt[1];
		if (self.formattingInfo)
			self._rich_text_runlist_map = rt_runlist;
		//if DEBUG:
		//	t1 = time.time()
		//	print("SST processing took %.2f seconds" % (t1 - t0, ), file=self.logfile)
	}
}

function parseSST(datatab, nstrings){
    //"Return list of strings"
    var datainx = 0;
    var ndatas = datatab.length;
    var data = datatab[0];
    var datalen = data.length;
    var pos = 8;
    var strings = [];
    var richtext_runs = {};
    var latin1 = "latin1";
    for(var _unused_i =0;_unused_i < nstrings; _unused_i++){
        var nchars =  data.readUInt16LE(pos);
        pos += 2;
        var options = data[pos];
        pos += 1;
        var rtcount = 0;
        var phosz = 0;
        if (options & 0x08){ // richtext
            rtcount = data.readUInt16LE(pos);
            pos += 2;
		}
        if (options & 0x04){ // phonetic
            phosz = data.readInt32LE(pos);
            pos += 4;
		}
        var accstrg ='';
        var charsgot = 0;
        while(1){
            var charsneed = nchars - charsgot;
			var charsavail = null;
			var rawstrg= null;
            if (options & 0x01){
                // Uncompressed UTF-16
                charsavail = Math.min((datalen - pos) >> 1, charsneed);
                rawstrg = data.slice(pos,pos+2*charsavail);
                // if DEBUG: print "SST U16: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
                try{
                    accstrg += comm.decode(rawstrg,0,rawstrg.length, "utf16le");
                }catch(e){
                    // print "SST U16: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
                    // Probable cause: dodgy data e.g. unfinished surrogate pair.
                    // E.g. file unicode2.xls in pyExcelerator's examples has cells containing
                    // unichr(i) for i in range(0x100000)
                    // so this will include 0xD800 etc
                    throw e;
				}
                pos += 2*charsavail;
            }else{
                // Note: this is COMPRESSED (not ASCII!) encoding!!!
                charsavail = Math.min(datalen - pos, charsneed);
                rawstrg = data.slice(pos,pos+charsavail);
                // if DEBUG: print "SST CMPRSD: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
                accstrg += comm.decode(rawstrg,0,rawstrg.length, latin1);
                pos += charsavail;
			}
            charsgot += charsavail;
            if (charsgot == nchars)
                break;
            datainx += 1;
            data = datatab[datainx];
            datalen = data.length;
            options = data[0];
            pos = 1;
        }
        if(rtcount){
            var runs = [];
            for(var runindex =0; runindex <rtcount; runindex++){
                if (pos == datalen){
                    pos = 0;
                    datainx += 1;
                    data = datatab[datainx];
                    datalen = data.length;
				}
                runs.push(data.readUInt16LE(pos));
				runs.push(data.readUInt16LE(pos+2));
                pos += 4;
			}
            richtext_runs[strings.length] = runs;
        }        
        pos += phosz; // size of the phonetic stuff to skip
        if(pos >= datalen){
            // adjust to correct position in next record
            pos = pos - datalen;
            datainx += 1;
            if(datainx < ndatas){
                data = datatab[datainx];
                datalen = data.length;
            }else
                assert (_unused_i == nstrings - 1, '_unused_i == nstrings - 1;');
		}
        strings.push(accstrg);
	}
    return [strings, richtext_runs];
}