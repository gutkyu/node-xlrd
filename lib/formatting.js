'use strict'
var util = require('util');
var comm = require('./common');
var assert = require('assert');
	

var DEBUG = 0;

var fmtTtoCellT = {};
fmtTtoCellT[comm.Format.FNU]=comm.XL_CELL_NUMBER
fmtTtoCellT[comm.Format.FUN]=comm.XL_CELL_NUMBER
fmtTtoCellT[comm.Format.FGE]=comm.XL_CELL_NUMBER
fmtTtoCellT[comm.Format.FDT]=comm.XL_CELL_DATE
fmtTtoCellT[comm.Format.FTX]=comm.XL_CELL_NUMBER // Yes, a number can be formatted as text.

	
function initialiseColourMap(book){

}

// === "Number formats" ===

//#
// "Number format" information from a FORMAT record.
function Format(formatKey, typ, formatStr){
	var self = this;
	
	if(formatKey === undefined) formatKey = 0 ;
	if(typ === undefined) typ = comm.Format.FUN;
	
    // The key into Book.formatMap
	self.formatKey = formatKey;
    // A classification that has been inferred from the format string.
    // Currently, this is used only to distinguish between numbers and dates.
    //   Values:
    //   FUN = 0 # unknown
    //   FDT = 1 # date
    //   FNU = 2 # number
    //   FGE = 3 # general
    //   FTX = 4 # text
    //#
    self.type = typ;
	
	self.formatStr = formatStr;
}

var stdFmtStrs = {
    // "std" == "standard for US English locale"
    // #### TODO ... a lot of work to tailor these to the user's locale.
    // See e.g. gnumeric-1.x.y/src/formats.c
    0x00: "General",
    0x01: "0",
    0x02: "0.00",
    0x03: "#,##0",
    0x04: "#,##0.00",
    0x05: "$#,##0_);($#,##0)",
    0x06: "$#,##0_);[Red]($#,##0)",
    0x07: "$#,##0.00_);($#,##0.00)",
    0x08: "$#,##0.00_);[Red]($#,##0.00)",
    0x09: "0%",
    0x0a: "0.00%",
    0x0b: "0.00E+00",
    0x0c: "# ?/?",
    0x0d: "# ??/??",
    0x0e: "m/d/yy",
    0x0f: "d-mmm-yy",
    0x10: "d-mmm",
    0x11: "mmm-yy",
    0x12: "h:mm AM/PM",
    0x13: "h:mm:ss AM/PM",
    0x14: "h:mm",
    0x15: "h:mm:ss",
    0x16: "m/d/yy h:mm",
    0x25: "#,##0_);(#,##0)",
    0x26: "#,##0_);[Red](#,##0)",
    0x27: "#,##0.00_);(#,##0.00)",
    0x28: "#,##0.00_);[Red](#,##0.00)",
    0x29: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    0x2a: "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
    0x2b: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    0x2c: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    0x2d: "mm:ss",
    0x2e: "[h]:mm:ss",
    0x2f: "mm:ss.0",
    0x30: "##0.0E+0",
    0x31: "@",
    };


var fmtCodeRanges = [ // both-inclusive ranges of "standard" format codes
    // Source: the openoffice.org doc't
    // and the OOXML spec Part 4, section 3.8.30
    [ 0,  0, comm.Format.FGE],
    [ 1, 13, comm.Format.FNU],
    [14, 22, comm.Format.FDT],
    [27, 36, comm.Format.FDT], // CJK date formats
    [37, 44, comm.Format.FNU],
    [45, 47, comm.Format.FDT],
    [48, 48, comm.Format.FNU],
    [49, 49, comm.Format.FTX],
    // Gnumeric assumes (or assumed) that built-in formats finish at 49, not at 163
    [50, 58, comm.Format.FDT], // CJK date formats
    [59, 62, comm.Format.FNU], // Thai number (currency?) formats
    [67, 70, comm.Format.FNU], // Thai number (currency?) formats
    [71, 81, comm.Format.FDT], // Thai date formats
];

var stdFmtCdTyps = {};
fmtCodeRanges.forEach(function(x){
	var lo =x[0], hi =x[1], typ =x[2];
	var len = hi+1;
	for(var i = lo; i<len; i++){
		stdFmtCdTyps[i] = typ;
	}
});

var dateChars = 'ymdhs';// year, month/minute, day, hour, second
dateChars += dateChars.toUpperCase() ;
var dateCharDict = {};
for(var i = 0;i<dateChars.length; i++) dateCharDict[dateChars[i]] = 5;


var skipChars = '$-+/(): ';
var skipCharDict = {};
for(i=0;i< skipChars.length;i++) skipCharDict[skipChars[i]] = 1;

var numCharDict = {
    '0': 5,
    '#': 5,
    '?': 5,
}


var nonDateFormats = {
    '0.00E+00':1,
    '##0.0E+0':1,
    'General' :1,
    'GENERAL' :1, // OOo Calc 1.1.4 does this.
    'general' :1,  // pyExcelerator 0.6.3 does this.
    '@'       :1,
};

var fmtBracketed = /\[[^]]*\]/;


function isDateFormatString(fmt){
	//self is book
	var self = this;
    // Heuristics:
    // Ignore "text" and [stuff in square brackets (aarrgghh -- see below)].
    // Handle backslashed-escaped chars properly.
    // E.g. hh\hmm\mss\s should produce a display like 23h59m59s
    // Date formats have one or more of ymdhs (caseless) in them.
    // Numeric formats have // and 0.
    // N.B. 'General"."' hence get rid of "text" first.
    // TODO: Find where formats are interpreted in Gnumeric
    // TODO: '[h]\\ \\h\\o\\u\\r\\s' ([h] means don't care about hours > 23)
    var state = 0;
    var s = '';
	
    for(i = 0; i < fmt.length; i++){
		var c = fmt[i];
        if( state == 0)
            if( c == '"')
                state = 1;
            else if( ['\\','_','*'][c])
                state = 2;
            else if( skipCharDict[c])
                ;
            else
                s += c;
        else if( state == 1)
            if (c == '"')
                state = 0;
        else if( state == 2)
            // Ignore char after backslash, underscore or asterisk
            state = 0;
        assert(0 <= state && state <= 2, '0 <= state <= 2') ;
	}
    if( self.verbosity >= 4)
        console.log(util.format("isDateFormatString: reduced format is %s", REPR(s)));
    s = s.replace(fmtBracketed,'');
    if(nonDateFormats[s])
        return false;
    state = 0;
    var separator = ";";
    var gotSep = 0;
	var numCount = 0;
    var dateCount = numCount;
    for(i =0; i< s.length ;i++){
		var c = s[i];
        if( dateCharDict[c] )
            dateCount += dateCharDict[c];
        else if(numCharDict[c])
            numCount += numCharDict[c];
        else if( c == separator)
            gotSep = 1;
	}
    // print numCount, dateCount, repr(fmt)
    if( dateCount && ! numCount)
        return true;
    if( numCount && ! dateCount)
        return false;
    if( dateCount)
        if( self.verbosity)
            console.log(util.format(
                'WARNING *** is_date_format: ambiguous d=%d n=%d fmt=%r\n',
                dateCount, numCount, fmt));
    else if( ! gotSep)
        if( self.verbosity)
            console.log(util.format(
                "WARNING *** format %s produces constant result\n",
                fmt));
    return dateCount > numCount;
}
function handleFormat(data, recType){
	//self is book
	var self = this;
	recType=recType||comm.XL_FORMAT;
    var DEBUG = 0;
    var bv = self.biffVersion;
    if( recType == comm.XL_FORMAT2)
        bv = Math.min(bv, 30);
    if( ! self.encoding)
        self.deriveEncoding();
    var strPos = 2;
	var fmtKey = null;
    if( bv >= 50)
        fmtKey = data.readUInt16LE(0);
    else{
        fmtKey = self.actualFmtCount;
        if( bv <= 30)
            strPos = 0;
	}
    self.actualFmtCount += 1;
	var uniStrg = null;
    if( bv >= 80)
        uniStrg = comm.unpackUnicode(data, 2);
    else
        uniStrg = comm.unpackString(data, strPos, self.encoding, 1);
    var blah = DEBUG || self.verbosity >= 3;
    if( blah)
        console.log(util.format(
            "FORMAT: count=%d fmtKey=0x%04x (%d) s=%r",
            self.actualFmtCount, fmtKey, fmtKey, uniStrg));
    var isDateStr = self.isDateFormatString(uniStrg);
    var typ = isDateStr?comm.Format.FDT:comm.Format.FGE;
    if( !(fmtKey > 163 || bv < 50)){
        // user_defined if fmtKey > 163
        // N.B. Gnumeric incorrectly starts these at 50 instead of 164 :-(
        // if earlier than BIFF 5, standard info is useless
        var stdTyp = stdFmtCdTyps[fmtKey];
		if(stdTyp === undefined || stdTyp === null)  stdTyp =  comm.Format.FUN;
        // print "std typ", stdTyp
        var isDateCode = stdTyp == comm.Format.FDT;
        if( self.verbosity && 0 < fmtKey && fmtKey < 50 && (isDateCode ^ isDateStr)){
            DEBUG = 2;
            console.log(util.format(
                "WARNING *** Conflict between std format key %d and its format string %s",
                fmtKey, uniStrg));
		}
	}
    if( DEBUG == 2)
        console.log(util.format(
            "typ: %d; isDateCode: %s; isDateStr: %s; fmt_strg: %s",
            typ, isDateCode, isDateStr, uniStrg));
    var fmtObj = new Format(fmtKey, typ, uniStrg);
    if( blah)
        fmtObj.dump(util.format("--- handleFormat [%d] ---" , self.actualFmtCount-1));
    self.formatMap[fmtKey] = fmtObj;
    self.formatList.push(fmtObj);
}

function checkColourIndexesInObj(book, obj, origIndex){}


function fillStandardFormats(book){
	for(var key in stdFmtCdTyps){
		if(book.formatMap[key] != undefined)  continue;
		var typ = stdFmtCdTyps[key];
		// Note: many standard format codes (mostly CJK date formats) have
		// format strings that vary by locale; xlrd does not (yet)
		// handle those; the type (date or numeric) is recorded but the fmtStr will be None.
		var fmtStr = stdFmtStrs[key];
		var fmtObj = new Format(key, typ, fmtStr);
		book.formatMap[key] = fmtObj;
	}
}

function handleXf(data){
	var self = this;
    //## self is a Book instance
    // DEBUG = 0
    var blah = DEBUG || self.verbosity >= 3
    var bv = self.biffVersion
    var xf = new XF();
	//todo 나중에 
	/*
    xf.alignment = XFAlignment()
    xf.alignment.indent_level = 0
    xf.alignment.shrink_to_fit = 0
    xf.alignment.text_direction = 0
    xf.border = XFBorder()
    xf.border.diag_up = 0
    xf.border.diag_down = 0
    xf.border.diag_colour_index = 0
    xf.border.diag_line_style = 0 // no line
    xf.background = XFBackground()
    xf.protection = XFProtection()
	*/
    // fill in the known standard formats
    if( bv >= 50 && ! self.xfCount)
        // i.e. do this once before we process the first XF record
        fillStandardFormats(self)
    if( bv >= 80){
		xf.formatKey = data.readUInt16LE(2);
		var pkdTypePar = data.readUInt16LE(4);
		xf.is_style = (pkdTypePar & 0x0004) >> 2;
		xf.parent_style_index = (pkdTypePar & 0xFFF0) >> 4;
		//todo 나중에 
		/*
        unpack_fmt = '<HHHBBBBIiH'
        (xf.font_index, xf.formatKey, pkdTypePar,
        pkd_align1, xf.alignment.rotation, pkd_align2,
        pkd_used, pkd_brdbkg1, pkd_brdbkg2, pkd_brdbkg3,
        ) = unpack(unpack_fmt, data[0:20])
        upkbits(xf.protection, pkdTypePar, (
            (0, 0x01, 'cell_locked'),
            (1, 0x02, 'formula_hidden'),
            ))
        upkbits(xf, pkdTypePar, (
            (2, 0x0004, 'is_style'),
            // Following is not in OOo docs, but is mentioned
            // in Gnumeric source and also in (deep breath)
            // org.apache.poi.hssf.record.ExtendedFormatRecord.java
            (3, 0x0008, 'lotus_123_prefix'), // Meaning is not known.
            (4, 0xFFF0, 'parent_style_index'),
            ))
        upkbits(xf.alignment, pkd_align1, (
            (0, 0x07, 'hor_align'),
            (3, 0x08, 'text_wrapped'),
            (4, 0x70, 'vert_align'),
            ))
        upkbits(xf.alignment, pkd_align2, (
            (0, 0x0f, 'indent_level'),
            (4, 0x10, 'shrink_to_fit'),
            (6, 0xC0, 'text_direction'),
            ))
        reg = pkd_used >> 2
        for attr_stem in \
            "format font alignment border background protection".split():
            attr = "_" + attr_stem + "_flag"
            setattr(xf, attr, reg & 1)
            reg >>= 1
        upkbitsL(xf.border, pkd_brdbkg1, (
            (0,  0x0000000f,  'left_line_style'),
            (4,  0x000000f0,  'right_line_style'),
            (8,  0x00000f00,  'top_line_style'),
            (12, 0x0000f000,  'bottom_line_style'),
            (16, 0x007f0000,  'left_colour_index'),
            (23, 0x3f800000,  'right_colour_index'),
            (30, 0x40000000,  'diag_down'),
            (31, 0x80000000, 'diag_up'),
            ))
        upkbits(xf.border, pkd_brdbkg2, (
            (0,  0x0000007F, 'top_colour_index'),
            (7,  0x00003F80, 'bottom_colour_index'),
            (14, 0x001FC000, 'diag_colour_index'),
            (21, 0x01E00000, 'diag_line_style'),
            ))
        upkbitsL(xf.background, pkd_brdbkg2, (
            (26, 0xFC000000, 'fill_pattern'),
            ))
        upkbits(xf.background, pkd_brdbkg3, (
            (0, 0x007F, 'pattern_colour_index'),
            (7, 0x3F80, 'background_colour_index'),
            ))
		*/
    }else if( bv >= 50){
		xf.formatKey = data.readUInt16LE(2);
		var pkdTypePar = data.readUInt16LE(4);
		xf.is_style = (pkdTypePar & 0x0004) >> 2;
		xf.parent_style_index = (pkdTypePar & 0xFFF0) >> 4;
		//todo 나중에 
		/*
        unpack_fmt = '<HHHBBIi'
        (xf.font_index, xf.formatKey, pkdTypePar,
        pkd_align1, pkd_orient_used,
        pkd_brdbkg1, pkd_brdbkg2,
        ) = unpack(unpack_fmt, data[0:16])
        upkbits(xf.protection, pkdTypePar, (
            (0, 0x01, 'cell_locked'),
            (1, 0x02, 'formula_hidden'),
            ))
        upkbits(xf, pkdTypePar, (
            (2, 0x0004, 'is_style'),
            (3, 0x0008, 'lotus_123_prefix'), // Meaning is not known.
            (4, 0xFFF0, 'parent_style_index'),
            ))
        upkbits(xf.alignment, pkd_align1, (
            (0, 0x07, 'hor_align'),
            (3, 0x08, 'text_wrapped'),
            (4, 0x70, 'vert_align'),
            ))
        orientation = pkd_orient_used & 0x03
        xf.alignment.rotation = [0, 255, 90, 180][orientation]
        reg = pkd_orient_used >> 2
        for attr_stem in \
            "format font alignment border background protection".split():
            attr = "_" + attr_stem + "_flag"
            setattr(xf, attr, reg & 1)
            reg >>= 1
        upkbitsL(xf.background, pkd_brdbkg1, (
            ( 0, 0x0000007F, 'pattern_colour_index'),
            ( 7, 0x00003F80, 'background_colour_index'),
            (16, 0x003F0000, 'fill_pattern'),
            ))
        upkbitsL(xf.border, pkd_brdbkg1, (
            (22, 0x01C00000,  'bottom_line_style'),
            (25, 0xFE000000, 'bottom_colour_index'),
            ))
        upkbits(xf.border, pkd_brdbkg2, (
            ( 0, 0x00000007, 'top_line_style'),
            ( 3, 0x00000038, 'left_line_style'),
            ( 6, 0x000001C0, 'right_line_style'),
            ( 9, 0x0000FE00, 'top_colour_index'),
            (16, 0x007F0000, 'left_colour_index'),
            (23, 0x3F800000, 'right_colour_index'),
            ))
		*/
    }else if( bv >= 40){
		xf.formatKey = data.readUInt8(1);
		var pkdTypePar = data.readUInt8(2);
		xf.is_style = (pkdTypePar & 0x0004) >> 2;
		xf.parent_style_index = (pkdTypePar & 0xFFF0) >> 4;
		//todo 나중에 
		/*
        unpack_fmt = '<BBHBBHI'
        (xf.font_index, xf.formatKey, pkdTypePar,
        pkd_align_orient, pkd_used,
        pkd_bkg_34, pkd_brd_34,
        ) = unpack(unpack_fmt, data[0:12])
        upkbits(xf.protection, pkdTypePar, (
            (0, 0x01, 'cell_locked'),
            (1, 0x02, 'formula_hidden'),
            ))
        upkbits(xf, pkdTypePar, (
            (2, 0x0004, 'is_style'),
            (3, 0x0008, 'lotus_123_prefix'), // Meaning is not known.
            (4, 0xFFF0, 'parent_style_index'),
            ))
        upkbits(xf.alignment, pkd_align_orient, (
            (0, 0x07, 'hor_align'),
            (3, 0x08, 'text_wrapped'),
            (4, 0x30, 'vert_align'),
            ))
        orientation = (pkd_align_orient & 0xC0) >> 6
        xf.alignment.rotation = [0, 255, 90, 180][orientation]
        reg = pkd_used >> 2
        for attr_stem in \
            "format font alignment border background protection".split():
            attr = "_" + attr_stem + "_flag"
            setattr(xf, attr, reg & 1)
            reg >>= 1
        upkbits(xf.background, pkd_bkg_34, (
            ( 0, 0x003F, 'fill_pattern'),
            ( 6, 0x07C0, 'pattern_colour_index'),
            (11, 0xF800, 'background_colour_index'),
            ))
        upkbitsL(xf.border, pkd_brd_34, (
            ( 0, 0x00000007,  'top_line_style'),
            ( 3, 0x000000F8,  'top_colour_index'),
            ( 8, 0x00000700,  'left_line_style'),
            (11, 0x0000F800,  'left_colour_index'),
            (16, 0x00070000,  'bottom_line_style'),
            (19, 0x00F80000,  'bottom_colour_index'),
            (24, 0x07000000,  'right_line_style'),
            (27, 0xF8000000, 'right_colour_index'),
            ))
		*/
    }else if( bv == 30){
		xf.formatKey = data.readUInt8(1);
		var pkdTypePar = data.readUInt8(2);
		xf.is_style = (pkdTypePar & 0x0004) >> 2;
		xf.parent_style_index = (pkdTypePar & 0xFFF0) >> 4;
		//todo 나중에 
		/*
        unpack_fmt = '<BBBBHHI'
        (xf.font_index, xf.formatKey, pkd_type_prot,
        pkd_used, pkd_align_par,
        pkd_bkg_34, pkd_brd_34,
        ) = unpack(unpack_fmt, data[0:12])
        upkbits(xf.protection, pkd_type_prot, (
            (0, 0x01, 'cell_locked'),
            (1, 0x02, 'formula_hidden'),
            ))
        upkbits(xf, pkd_type_prot, (
            (2, 0x0004, 'is_style'),
            (3, 0x0008, 'lotus_123_prefix'), // Meaning is not known.
            ))
        upkbits(xf.alignment, pkd_align_par, (
            (0, 0x07, 'hor_align'),
            (3, 0x08, 'text_wrapped'),
            ))
        upkbits(xf, pkd_align_par, (
            (4, 0xFFF0, 'parent_style_index'),
            ))
        reg = pkd_used >> 2
        for attr_stem in \
            "format font alignment border background protection".split():
            attr = "_" + attr_stem + "_flag"
            setattr(xf, attr, reg & 1)
            reg >>= 1
        upkbits(xf.background, pkd_bkg_34, (
            ( 0, 0x003F, 'fill_pattern'),
            ( 6, 0x07C0, 'pattern_colour_index'),
            (11, 0xF800, 'background_colour_index'),
            ))
        upkbitsL(xf.border, pkd_brd_34, (
            ( 0, 0x00000007,  'top_line_style'),
            ( 3, 0x000000F8,  'top_colour_index'),
            ( 8, 0x00000700,  'left_line_style'),
            (11, 0x0000F800,  'left_colour_index'),
            (16, 0x00070000,  'bottom_line_style'),
            (19, 0x00F80000,  'bottom_colour_index'),
            (24, 0x07000000,  'right_line_style'),
            (27, 0xF8000000, 'right_colour_index'),
            ))
        xf.alignment.vert_align = 2 // bottom
        xf.alignment.rotation = 0
		*/
    }else if( bv == 21){
        //### Warning: incomplete treatment; formattingInfo not fully supported.
        //### Probably need to offset incoming BIFF2 XF[n] to BIFF8-like XF[n+16],
        //### and create XF[0:16] like the standard ones in BIFF8
        //### *AND* add 16 to all XF references in cell records :-(
		
		xf.formatKey = data.readUInt8(2) & 0x3F ;
		xf.parent_style_index = 0;
		//todo 나중에 
		/*
        (xf.font_index, format_etc, halign_etc) = unpack('<BxBB', data)
        xf.formatKey = format_etc & 0x3F
        upkbits(xf.protection, format_etc, (
            (6, 0x40, 'cell_locked'),
            (7, 0x80, 'formula_hidden'),
            ))
        upkbits(xf.alignment, halign_etc, (
            (0, 0x07, 'hor_align'),
            ))
        for mask, side in ((0x08, 'left'), (0x10, 'right'), (0x20, 'top'), (0x40, 'bottom')):
            if( halign_etc & mask:
                colour_index, line_style = 8, 1 // black, thin
            else
                colour_index, line_style = 0, 0 // none, none
            setattr(xf.border, side + '_colour_index', colour_index)
            setattr(xf.border, side + '_line_style', line_style)
        bg = xf.background
        if( halign_etc & 0x80)
            bg.fill_pattern = 17
        else
            bg.fill_pattern = 0
        bg.background_colour_index = 9 // white
        bg.pattern_colour_index = 8 // black
        xf.parent_style_index = 0 // ???????????
        xf.alignment.vert_align = 2 // bottom
        xf.alignment.rotation = 0
        for attr_stem in \
            "format font alignment border background protection".split():
            attr = "_" + attr_stem + "_flag"
            setattr(xf, attr, 1)
		*/
    }else
        throw new Error('programmer stuff-up: bv='+ bv);

    xf.xfIndex = self.xfList.length
    self.xfList.push(xf);
    self.xfCount += 1;
    if( blah)
        xf.dump(util.format("--- handleXf: xf[%d] ---" , xf.xfIndex));
        
	var fmt = null;
	var cellTyp = null;
	if(	(fmt = self.formatMap[xf.formatKey]) == undefined || 
		fmt.type == undefined || 
		(cellTyp=fmtTtoCellT[fmt.type]) == undefined ){
		cellTyp = comm.XL_CELL_NUMBER;
	}
 
    self._xf_index_to_xl_type_map[xf.xfIndex] = cellTyp;
    // Now for some assertions ...
    if( self.formattingInfo)
        if( self.verbosity && xf.is_style && xf.parent_style_index != 0x0FFF){
            var msg = "WARNING *** XF[%d] is a style XF but parent_style_index is 0x%04x, not 0x0fff";
            console.log(util.format(msg, xf.xfIndex, xf.parent_style_index));
		}
        checkColourIndexesInObj(self, xf, xf.xfIndex);
    if( !self.formatMap[xf.formatKey]){
        var msg = "WARNING *** XF[%d] unknown (raw) format key (%d)";
        if( self.verbosity)
            console.log(util.format(msg,xf.xfIndex, xf.formatKey));
        xf.formatKey = 0;
	}
}

function xfEpilogue(){
    // self is a Book instance.
	var self = this;
    self._xfEpilogueDone = 1;
    var numXfs = self.xfList.length;
    var blah = DEBUG  ||  self.verbosity >= 3;
    var blah1 = DEBUG  ||  self.verbosity >= 1;
    if (blah) console.log( "xfEpilogue called ...");

    function checkSame(book_arg, xfArg, parentArg, attr){
        //the _arg caper is to avoid a Warning msg from Python 2.1 :-(
        if(getattr(xfArg, attr) != getattr(parentArg, attr))
            console.log(util.format("NOTE !!! XF[%d] parent[%d] %s different",
                xfArg.xfIndex, parentArg.xfIndex, attr));
	}
	
    for(var xfx =0 ; xfx < numXfs; xfx++){
        var xf = self.xfList[xfx];
        if(self.formatMap[xf.formatKey] == undefined){
            var msg = "ERROR *** XF[%d] unknown format key (%d, 0x%04x)";
            console.log(util.format(msg,xf.xfIndex, xf.formatKey, xf.formatKey));
            xf.formatKey = 0;
		}
        var fmt = self.formatMap[xf.formatKey];
        var cellTyp = fmtTtoCellT[fmt.type];
        self._xf_index_to_xl_type_map[xf.xfIndex] = cellTyp;
        // Now for some assertions etc
        if (! self.formattingInfo)
            continue;
        if( xf.is_style)
            continue;
        if (!(0 <= xf.parent_style_index && xf.parent_style_index  < numXfs)){
            if (blah1)
                console.log(util.format("WARNING *** XF[%d]: is_style=%d but parent_style_index=%d",
                    xf.xfIndex, xf.is_style, xf.parent_style_index));
            // make it conform
            xf.parent_style_index = 0;
		}
        if(self.biffVersion >= 30){
            if (blah1){
                if (xf.parent_style_index == xf.xfIndex)
                    console.log(util.format(
                        "NOTE !!! XF[%d]: parent_style_index is also %d",
                        xf.xfIndex, xf.parent_style_index));
                else if(! self.xfList[xf.parent_style_index].is_style)
                    console.log(util.format(
                        "NOTE !!! XF[%d]: parent_style_index is %d; style flag not set",
                        xf.xfIndex, xf.parent_style_index));
			}
            if(blah1 && xf.parent_style_index > xf.xfIndex)
                console.log(util.format(
                    "NOTE !!! XF[%d]: parent_style_index is %d; out of order?",
                    xf.xfIndex, xf.parent_style_index));
            var parent = self.xfList[xf.parent_style_index];
            if( xf._alignment_flag && ! parent._alignment_flag)
                if (blah1) checkSame(self, xf, parent, 'alignment');
            if( xf._background_flag && ! parent._background_flag)
                if (blah1) checkSame(self, xf, parent, 'background');
            if( xf._border_flag && ! parent._border_flag)
                if (blah1) checkSame(self, xf, parent, 'border');
            if( xf._protection_flag && ! parent._protection_flag)
                if (blah1) checkSame(self, xf, parent, 'protection');
            if( xf._format_flag && ! parent._format_flag)
                if (blah1 && xf.formatKey != parent.formatKey)
                    console.log(util.format(
                        "NOTE !!! XF[%d] fmtk=%d, parent[%d] fmtk=%r\n%s / %s",
                        xf.xfIndex, xf.formatKey, parent.xfIndex, parent.formatKey,
                        self.formatMap[xf.formatKey].formatStr,
                        self.formatMap[parent.formatKey].formatStr));
            if( xf._font_flag && ! parent._font_flag)
                if (blah1 && xf.font_index != parent.font_index)
                    console.log(util.format(
                        "NOTE !!! XF[%d] fontx=%d, parent[%d] fontx=%r\n",
                        xf.xfIndex, xf.font_index, parent.xfIndex, parent.font_index));
		}
	}
	
}
exports.initialiseWorkbook = function (book){
    initialiseColourMap(book);
    book._xfEpilogueDone = 0;
    var methods = [	
				//handle_font,
				//handle_efont,
				handleFormat,
				isDateFormatString,
				//handle_palette, 
				//palette_epilogue,
				//handle_style,
				handleXf,
				xfEpilogue
			];
	var bkProto = Object.getPrototypeOf(book);
    methods.forEach(function(x){bkProto[x.name] = x;});
	
}


// eXtended Formatting information for cells, rows, columns and styles.
//
//  Each of the 6 flags below describes the validity of
// a specific group of attributes.
//   
// In cell XFs, flag==0 means the attributes of the parent style XF are used,
// (but only if the attributes are valid there); flag==1 means the attributes
// of this XF are used. 
// In style XFs, flag==0 means the attribute setting is valid; flag==1 means
// the attribute should be ignored. 
// Note that the API
// provides both "raw" XFs and "computed" XFs -- in the latter case, cell XFs
// have had the above inheritance mechanism applied.
//  

function XF(){
	var self = this;

    // 0 = cell XF, 1 = style XF
    self.is_style = 0;

    // cell XF: Index into Book.xfList
    // of this XF's style XF  
    // style XF: 0xFFF
    self.parent_style_index = 0;
    self._format_flag = 0;
    self._font_flag = 0;
    self._alignment_flag = 0;
    self._border_flag = 0;
    self._background_flag = 0;
    self._protection_flag = 0;
    // Index into Book.xfList
    self.xfIndex = 0;
    // Index into Book.fontList
    self.font_index = 0;
    // Key into Book.formatMap
    //  
    // Warning: OOo docs on the XF record call this "Index to FORMAT record".
    // It is not an index in the Python sense. It is a key to a map.
    // It is true  only  for Excel 4.0 and earlier files
    // that the key into formatMap from an XF instance
    // is the same as the index into formatList, and  only 
    // if the index is less than 164.
    //  
    self.formatKey = 0;
    // An instance of an XFProtection object.
    self.protection = null;
    // An instance of an XFBackground object.
    self.background = null;
    // An instance of an XFAlignment object.
    self.alignment = null;
    // An instance of an XFBorder object.
    self.border = null;
}