var iconv = require('iconv-lite');
// unknown, date, number, general, text
var format = exports.Format = {FUN:0, FDT:1, FNU:2, FGE:3, FTX:4 };
exports.DATEFORMAT = format.FDT;
exports.NUMBERFORMAT = format.FNU;


exports.XL_CELL_EMPTY =0;
exports.XL_CELL_TEXT =1;
exports.XL_CELL_NUMBER =2;
exports.XL_CELL_DATE =3;
exports.XL_CELL_BOOLEAN =4;
exports.XL_CELL_ERROR =5;
exports.XL_CELL_BLANK =6; // for use in debugging, gathering stats, etc


exports.biffCodeStringMap = {
		0:  "(not BIFF)",
		20: "2.0",
		21: "2.1",
		30: "3",
		40: "4S",
		45: "4W",
		50: "5",
		70: "7",
		80: "8",
		85: "8X",
};

/*
* This dictionary can be used to produce a text version of the internal codes
* that Excel uses for error cells. Here are its contents:
* 
* 0x00: '#NULL!',  // Intersection of two cell ranges is empty
* 0x07: '#DIV/0!', // Division by zero
* 0x0F: '#VALUE!', // Wrong type of operand
* 0x17: '#REF!',   // Illegal or deleted cell reference
* 0x1D: '#NAME?',  // Wrong exports.function or range name
* 0x24: '#NUM!',   // Value range overflow
* 0x2A: '#N/A',    // Argument or exports.function not available
* 
*/
exports.errorCodeStringMap = {
    0x00: '#NULL!',  // Intersection of two cell ranges is empty
    0x07: '#DIV/0!', // Division by zero
    0x0F: '#VALUE!', // Wrong type of operand
    0x17: '#REF!',   // Illegal or deleted cell reference
    0x1D: '#NAME?',  // Wrong exports.function or range name
    0x24: '#NUM!',   // Value range overflow
    0x2A: '#N/A',    // Argument or exports.function not available
};

exports.BIFF_FIRST_UNICODE = 80;

exports.XL_WORKBOOK_GLOBALS = exports.WBKBLOBAL = 0x5;
exports.XL_WORKBOOK_GLOBALS_4W = 0x100;
exports.XL_WORKSHEET = exports.WRKSHEET = 0x10;

exports.XL_BOUNDSHEET_WORKSHEET = 0x00;
exports.XL_BOUNDSHEET_CHART     = 0x02;
exports.XL_BOUNDSHEET_VB_MODULE = 0x06;

// XL_RK2 = 0x7e
exports.XL_ARRAY  = 0x0221;
exports.XL_ARRAY2 = 0x0021;
exports.XL_BLANK = 0x0201;
exports.XL_BLANK_B2 = 0x01;
exports.XL_BOF = 0x809;
exports.XL_BOOLERR = 0x205;
exports.XL_BOOLERR_B2 = 0x5;
exports.XL_BOUNDSHEET = 0x85;
exports.XL_BUILTINFMTCOUNT = 0x56;
exports.XL_CF = 0x01B1;
exports.XL_CODEPAGE = 0x42;
exports.XL_COLINFO = 0x7D;
exports.XL_COLUMNDEFAULT = 0x20; // BIFF2 only
exports.XL_COLWIDTH = 0x24; // BIFF2 only
exports.XL_CONDFMT = 0x01B0;
exports.XL_CONTINUE = 0x3c;
exports.XL_COUNTRY = 0x8C;
exports.XL_DATEMODE = 0x22;
exports.XL_DEFAULTROWHEIGHT = 0x0225;
exports.XL_DEFCOLWIDTH = 0x55;
exports.XL_DIMENSION = 0x200;
exports.XL_DIMENSION2 = 0x0;
exports.XL_EFONT = 0x45;
exports.XL_EOF = 0x0a;
exports.XL_EXTERNNAME = 0x23;
exports.XL_EXTERNSHEET = 0x17;
exports.XL_EXTSST = 0xff;
exports.XL_FEAT11 = 0x872;
exports.XL_FILEPASS = 0x2f;
exports.XL_FONT = 0x31;
exports.XL_FONT_B3B4 = 0x231;
exports.XL_FORMAT = 0x41e;
exports.XL_FORMAT2 = 0x1E; // BIFF2, BIFF3
exports.XL_FORMULA = 0x6;
exports.XL_FORMULA3 = 0x206;
exports.XL_FORMULA4 = 0x406;
exports.XL_GCW = 0xab;
exports.XL_HLINK = 0x01B8;
exports.XL_QUICKTIP = 0x0800;
exports.XL_HORIZONTALPAGEBREAKS = 0x1b;
exports.XL_INDEX = 0x20b;
exports.XL_INTEGER = 0x2; // BIFF2 only
exports.XL_IXFE = 0x44; // BIFF2 only
exports.XL_LABEL = 0x204;
exports.XL_LABEL_B2 = 0x04;
exports.XL_LABELRANGES = 0x15f;
exports.XL_LABELSST = 0xfd;
exports.XL_LEFTMARGIN = 0x26;
exports.XL_TOPMARGIN = 0x28;
exports.XL_RIGHTMARGIN = 0x27;
exports.XL_BOTTOMMARGIN = 0x29;
exports.XL_HEADER = 0x14;
exports.XL_FOOTER = 0x15;
exports.XL_HCENTER = 0x83;
exports.XL_VCENTER = 0x84;
exports.XL_MERGEDCELLS = 0xE5;
exports.XL_MSO_DRAWING = 0x00EC;
exports.XL_MSO_DRAWING_GROUP = 0x00EB;
exports.XL_MSO_DRAWING_SELECTION = 0x00ED;
exports.XL_MULRK = 0xbd;
exports.XL_MULBLANK = 0xbe;
exports.XL_NAME = 0x18;
exports.XL_NOTE = 0x1c;
exports.XL_NUMBER = 0x203;
exports.XL_NUMBER_B2 = 0x3;
exports.XL_OBJ = 0x5D;
exports.XL_PAGESETUP = 0xA1;
exports.XL_PALETTE = 0x92;
exports.XL_PANE = 0x41;
exports.XL_PRINTGRIDLINES = 0x2B;
exports.XL_PRINTHEADERS = 0x2A;
exports.XL_RK = 0x27e;
exports.XL_ROW = 0x208;
exports.XL_ROW_B2 = 0x08;
exports.XL_RSTRING = 0xd6;
exports.XL_SCL = 0x00A0;
exports.XL_SHEETHDR = 0x8F; // BIFF4W only
exports.XL_SHEETPR = 0x81;
exports.XL_SHEETSOFFSET = 0x8E; // BIFF4W only
exports.XL_SHRFMLA = 0x04bc;
exports.XL_SST = 0xfc;
exports.XL_STANDARDWIDTH = 0x99;
exports.XL_STRING = 0x207;
exports.XL_STRING_B2 = 0x7;
exports.XL_STYLE = 0x293;
exports.XL_SUPBOOK = 0x1AE; // aka EXTERNALBOOK in OOo docs
exports.XL_TABLEOP = 0x236;
exports.XL_TABLEOP2 = 0x37;
exports.XL_TABLEOP_B2 = 0x36;
exports.XL_TXO = 0x1b6;
exports.XL_UNCALCED = 0x5e;
exports.XL_UNKNOWN = 0xffff;
exports.XL_VERTICALPAGEBREAKS = 0x1a;
exports.XL_WINDOW2    = 0x023E;
exports.XL_WINDOW2_B2 = 0x003E;
exports.XL_WRITEACCESS = 0x5C;
exports.XL_WSBOOL = exports.XL_SHEETPR;
exports.XL_XF = 0xe0;
exports.XL_XF2 = 0x0043; // BIFF2 version of XF record
exports.XL_XF3 = 0x0243; // BIFF3 version of XF record
exports.XL_XF4 = 0x0443; // BIFF4 version of XF record

exports.bofLen = {0x0809: 8, 0x0409: 6, 0x0209: 6, 0x0009: 4};
exports.bofCodes = [0x0809, 0x0409, 0x0209, 0x0009];

exports.XL_FORMULA_OPCODES = [0x0006, 0x0406, 0x0206];

exports.cellOpcodeList = [
    exports.XL_BOOLERR,
    exports.XL_FORMULA,
    exports.XL_FORMULA3,
    exports.XL_FORMULA4,
    exports.XL_LABEL,
    exports.XL_LABELSST,
    exports.XL_MULRK,
    exports.XL_NUMBER,
    exports.XL_RK,
    exports.XL_RSTRING,
    ];
exports.cellOpcodeDic = {};
exports.cellOpcodeList.forEach(function(x){ exports.cellOpcodeDic[exports.cellOpcode] = 1;});

//todo 
/*
exports.function is_cell_opcode(c){
    return c in  cellOpcodeDic;
}
exports.function upkbits(tgt_obj, src, manifest, local_setattr=setattr){
    for n, mask, attr in manifest:
        local_setattr(tgt_obj, attr, (src & mask) >> n)
}
exports.function upkbitsL(tgt_obj, src, manifest, local_setattr=setattr, local_int=int){
    for n, mask, attr in manifest:
        local_setattr(tgt_obj, attr, local_int((src & mask) >> n))
}
*/

exports.decode = decode = function(data,pos,len,encoding){
	if(Buffer.isEncoding(encoding)) 
		return data.toString(encoding,pos, pos+len);
	var buf = pos==0 && data.length == len ? data : data.slice(pos,pos+len);
	return iconv.decode(buf,encoding);
}

//unpackString
exports.unpackString = function(data,pos,encoding,lenRecordLength){
	lenRecordLength == lenRecordLength || 1;
	if(lenRecordLength >2) throw new Error('decode error : lenRecordLength ',lenRecordLength,' > 2');
	var len = lenRecordLength == 1? data.readUInt8(pos):data.readUInt16LE(pos);
    pos += lenRecordLength;
	return decode( data,pos,len,encoding);
}

exports.unpackStringUpdatePos = function(data, pos, encoding, lenRecordLength, knownLength){
	lenRecordLength = lenRecordLength || 1;
	knownLength = knownLength || null;
	var nchars = 0;
    if (knownLength)
        // On a NAME record, the length byte is detached from the front of the string.
        nchars = knownLength;
    else{
        nchars = lenRecordLength == 1? data.readUInt8(pos):data.readUInt16LE(pos);
        pos += lenRecordLength;
	}
    newpos = pos + nchars;
    return [decode(data, pos, nchars,encoding), newpos];
}
exports.unpackUnicode = function(data, pos, lenRecordLength){
	lenRecordLength = lenRecordLength || 2;
    nchars = lenRecordLength == 1? data.readUInt8(pos):data.readUInt16LE(pos);
	var str = null;
    if (! nchars)
        // Ambiguous whether 0-length string should have an "options" byte.
        // Avoid crash if missing.
        return '';
    pos += lenRecordLength;
    var options = data[pos];//BOM
    pos += 1;
    // phonetic = options & 0x04
    // richtext = options & 0x08
    if (options & 0x08)
        // rt = unpack('<H', data[pos:pos+2])[0] // unused
        pos += 2;
    if (options & 0x04)
        // sz = unpack('<i', data[pos:pos+4])[0] // unused
        pos += 4;
    if (options & 0x01){
        // Uncompressed UTF-16-LE
        str = data.toString('utf16le',pos,pos+2*nchars);
        // if DEBUG: print "nchars=%d pos=%d rawStr=%r" % (nchars, pos, rawStr)
        
        // pos += 2*nchars
    }else{
        // Note: this is COMPRESSED (not ASCII!) encoding!!!
        // Merely returning the raw bytes would work OK 99.99% of the time
        // if the local codePage was cp1252 -- however this would rapidly go pear-shaped
        // for other codepages so we grit our Anglocentric teeth and return Unicode :-)

        //strg = unicode(data[pos:pos+nchars], "latin_1");
		return decode(data,pos,nchars,'latin1');
        // pos += nchars
	}
    // if richtext:
    //     pos += 4 * rt
    // if phonetic:
    //     pos += sz
    // return (strg, pos)
    return str;
}
/*
exports.function unpack_unicode_update_pos(data, pos, lenRecordLength=2, knownLength=None){
    "Return (unicode_strg, updated value of pos)"
    if knownLength is not None:
        // On a NAME record, the length byte is detached from the front of the string.
        nchars = knownLength
    else:
        nchars = unpack('<' + 'BH'[lenRecordLength-1], data[pos:pos+lenRecordLength])[0]
        pos += lenRecordLength
    if not nchars and not data[pos:]:
        // Zero-length string with no options byte
        return (UNICODE_LITERAL(""), pos)
    options = BYTES_ORD(data[pos])
    pos += 1
    phonetic = options & 0x04
    richtext = options & 0x08
    if richtext:
        rt = unpack('<H', data[pos:pos+2])[0]
        pos += 2
    if phonetic:
        sz = unpack('<i', data[pos:pos+4])[0]
        pos += 4
    if options & 0x01:
        // Uncompressed UTF-16-LE
        strg = unicode(data[pos:pos+2*nchars], 'utf_16_le')
        pos += 2*nchars
    else:
        // Note: this is COMPRESSED (not ASCII!) encoding!!!
        strg = unicode(data[pos:pos+nchars], "latin_1")
        pos += nchars
    if richtext:
        pos += 4 * rt
    if phonetic:
        pos += sz
    return (strg, pos)
}
exports.function unpack_cell_range_address_list_update_pos(
    output_list, data, pos, biffVersion, addr_size=6){
    // output_list is updated in situ
    assert addr_size in (6, 8)
    // Used to assert size == 6 if not BIFF8, but pyWLWriter writes
    // BIFF8-only MERGEDCELLS records in a BIFF5 file!
    n, = unpack("<H", data[pos:pos+2])
    pos += 2
    if n:
        if addr_size == 6:
            fmt = "<HHBB"
        else:
            fmt = "<HHHH"
        for _unused in xrange(n):
            ra, rb, ca, cb = unpack(fmt, data[pos:pos+addr_size])
            output_list.append((ra, rb+1, ca, cb+1))
            pos += addr_size
    return pos
}
_brecstrg = """\
0000 DIMENSIONS_B2
0001 BLANK_B2
0002 INTEGER_B2_ONLY
0003 NUMBER_B2
0004 LABEL_B2
0005 BOOLERR_B2
0006 FORMULA
0007 STRING_B2
0008 ROW_B2
0009 BOF_B2
000A EOF
000B INDEX_B2_ONLY
000C CALCCOUNT
000D CALCMODE
000E PRECISION
000F REFMODE
0010 DELTA
0011 ITERATION
0012 PROTECT
0013 PASSWORD
0014 HEADER
0015 FOOTER
0016 EXTERNCOUNT
0017 EXTERNSHEET
0018 NAME_B2,5+
0019 WINDOWPROTECT
001A VERTICALPAGEBREAKS
001B HORIZONTALPAGEBREAKS
001C NOTE
001D SELECTION
001E FORMAT_B2-3
001F BUILTINFMTCOUNT_B2
0020 COLUMNDEFAULT_B2_ONLY
0021 ARRAY_B2_ONLY
0022 DATEMODE
0023 EXTERNNAME
0024 COLWIDTH_B2_ONLY
0025 DEFAULTROWHEIGHT_B2_ONLY
0026 LEFTMARGIN
0027 RIGHTMARGIN
0028 TOPMARGIN
0029 BOTTOMMARGIN
002A PRINTHEADERS
002B PRINTGRIDLINES
002F FILEPASS
0031 FONT
0032 FONT2_B2_ONLY
0036 TABLEOP_B2
0037 TABLEOP2_B2
003C CONTINUE
003D WINDOW1
003E WINDOW2_B2
0040 BACKUP
0041 PANE
0042 CODEPAGE
0043 XF_B2
0044 IXFE_B2_ONLY
0045 EFONT_B2_ONLY
004D PLS
0051 DCONREF
0055 DEFCOLWIDTH
0056 BUILTINFMTCOUNT_B3-4
0059 XCT
005A CRN
005B FILESHARING
005C WRITEACCESS
005D OBJECT
005E UNCALCED
005F SAVERECALC
0063 OBJECTPROTECT
007D COLINFO
007E RK2_mythical_?
0080 GUTS
0081 WSBOOL
0082 GRIDSET
0083 HCENTER
0084 VCENTER
0085 BOUNDSHEET
0086 WRITEPROT
008C COUNTRY
008D HIDEOBJ
008E SHEETSOFFSET
008F SHEETHDR
0090 SORT
0092 PALETTE
0099 STANDARDWIDTH
009B FILTERMODE
009C FNGROUPCOUNT
009D AUTOFILTERINFO
009E AUTOFILTER
00A0 SCL
00A1 SETUP
00AB GCW
00BD MULRK
00BE MULBLANK
00C1 MMS
00D6 RSTRING
00D7 DBCELL
00DA BOOKBOOL
00DD SCENPROTECT
00E0 XF
00E1 INTERFACEHDR
00E2 INTERFACEEND
00E5 MERGEDCELLS
00E9 BITMAP
00EB MSO_DRAWING_GROUP
00EC MSO_DRAWING
00ED MSO_DRAWING_SELECTION
00EF PHONETIC
00FC SST
00FD LABELSST
00FF EXTSST
013D TABID
015F LABELRANGES
0160 USESELFS
0161 DSF
01AE SUPBOOK
01AF PROTECTIONREV4
01B0 CONDFMT
01B1 CF
01B2 DVAL
01B6 TXO
01B7 REFRESHALL
01B8 HLINK
01BC PASSWORDREV4
01BE DV
01C0 XL9FILE
01C1 RECALCID
0200 DIMENSIONS
0201 BLANK
0203 NUMBER
0204 LABEL
0205 BOOLERR
0206 FORMULA_B3
0207 STRING
0208 ROW
0209 BOF
020B INDEX_B3+
0218 NAME
0221 ARRAY
0223 EXTERNNAME_B3-4
0225 DEFAULTROWHEIGHT
0231 FONT_B3B4
0236 TABLEOP
023E WINDOW2
0243 XF_B3
027E RK
0293 STYLE
0406 FORMULA_B4
0409 BOF
041E FORMAT
0443 XF_B4
04BC SHRFMLA
0800 QUICKTIP
0809 BOF
0862 SHEETLAYOUT
0867 SHEETPROTECTION
0868 RANGEPROTECTION
"""

exports.biff_rec_name_dict = {}
for _buff in _brecstrg.splitlines():
    _numh, _name = _buff.split()
    biff_rec_name_dict[int(_numh, 16)] = _name
exports.del _buff=null;
exports._name=null;
exports._brecstrg=null;

exports.function hex_char_dump(strg, ofs, dlen, base=0, fout=sys.stdout, unnumbered=False){
    endpos = min(ofs + dlen, len(strg))
    pos = ofs
    numbered = not unnumbered
    num_prefix = ''
    while pos < endpos:
        endsub = min(pos + 16, endpos)
        substrg = strg[pos:endsub]
        lensub = endsub - pos
        if lensub <= 0 or lensub != len(substrg):
            fprintf(
                sys.stdout,
                '??? hex_char_dump: ofs=%d dlen=%d base=%d -> endpos=%d pos=%d endsub=%d substrg=%r\n',
                ofs, dlen, base, endpos, pos, endsub, substrg)
            break
        hexd = ''.join(["%02x " % BYTES_ORD(c) for c in substrg])
        
        chard = ''
        for c in substrg:
            c = chr(BYTES_ORD(c))
            if c == '\0':
                c = '~'
            elif not (' ' <= c <= '~'):
                c = '?'
            chard += c
        if numbered:
            num_prefix = "%5d: " %  (base+pos-ofs)
        
        fprintf(fout, "%s     %-48s %s\n", num_prefix, hexd, chard)
        pos = endsub
}
exports.function biff_dump(mem, stream_offset, streamLen, base=0, fout=sys.stdout, unnumbered=False){
    pos = stream_offset
    stream_end = stream_offset + streamLen
    adj = base - stream_offset
    dummies = 0
    numbered = not unnumbered
    num_prefix = ''
    while stream_end - pos >= 4:
        rc, length = unpack('<HH', mem[pos:pos+4])
        if rc == 0 and length == 0:
            if mem[pos:] == b'\0' * (stream_end - pos):
                dummies = stream_end - pos
                savPos = pos
                pos = stream_end
                break
            if dummies:
                dummies += 4
            else:
                savPos = pos
                dummies = 4
            pos += 4
        else:
            if dummies:
                if numbered:
                    num_prefix =  "%5d: " % (adj + savPos)
                fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
                dummies = 0
            recname = biff_rec_name_dict.get(rc, '<UNKNOWN>')
            if numbered:
                num_prefix = "%5d: " % (adj + pos)
            fprintf(fout, "%s%04x %s len = %04x (%d)\n", num_prefix, rc, recname, length, length)
            pos += 4
            hex_char_dump(mem, pos, length, adj+pos, fout, unnumbered)
            pos += length
    if dummies:
        if numbered:
            num_prefix =  "%5d: " % (adj + savPos)
        fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
    if pos < stream_end:
        if numbered:
            num_prefix = "%5d: " % (adj + pos)
        fprintf(fout, "%s---- Misc bytes at end ----\n", num_prefix)
        hex_char_dump(mem, pos, stream_end-pos, adj + pos, fout, unnumbered)
    elif pos > stream_end:
        fprintf(fout, "Last dumped record has length (%d) that is too large\n", length)
}
exports.function biff_count_records(mem, stream_offset, streamLen, fout=sys.stdout){
    pos = stream_offset
    stream_end = stream_offset + streamLen
    tally = {}
    while stream_end - pos >= 4:
        rc, length = unpack('<HH', mem[pos:pos+4])
        if rc == 0 and length == 0:
            if mem[pos:] == b'\0' * (stream_end - pos):
                break
            recname = "<Dummy (zero)>"
        else:
            recname = biff_rec_name_dict.get(rc, None)
            if recname is None:
                recname = "Unknown_0x%04X" % rc
        if recname in tally:
            tally[recname] += 1
        else:
            tally[recname] = 1
        pos += length + 4
    slist = sorted(tally.items())
    for recname, count in slist:
        print("%8d %s" % (count, recname), file=fout)
}
*/
exports.encodingFromCodePage = {
    1200 : 'utf16le',
    10000: 'mac_roman',
    10006: 'mac_greek', // guess
    10007: 'mac_cyrillic', // guess
    10029: 'mac_latin2', // guess
    10079: 'mac_iceland', // guess
    10081: 'mac_turkish', // guess
    32768: 'mac_roman',
    32769: 'cp1252',
    }
/*
// some more guessing, for Indic scripts
// codePage 57000 range:
// 2 Devanagari [0]
// 3 Bengali [1]
// 4 Tamil [5]
// 5 Telegu [6]
// 6 Assamese [1] c.f. Bengali
// 7 Oriya [4]
// 8 Kannada [7]
// 9 Malayalam [8]
// 10 Gujarati [3]
// 11 Gurmukhi [2]
*/

//Windows Country Identifier 
exports.windowsCountryIdList = {
    1:'USA',
    2:'Canada',
    7:'Russia',
    20:'Egypt',
    27:'South Africa',
    30:'Greece',
    31:'Netherlands',
    32:'Belgium',
    33:'France',
    34:'Spain',
    36:'Hungary',
    39:'Italy',
    40:'Romania',
    41:'Switzerland',
    43:'Austria',
    44:'United Kingdom',
    45:'Denmark',
    46:'Sweden',
    47:'Norway',
    48:'Poland',
    49:'Germany',
    51:'Peru',
    52:'Mexico',
    53:'Cuba',
    54:'Argentinia',
    55:'Brazil',
    56:'Chile',
    57:'Colombia',
    58:'Venezuela',
    60:'Malaysia',
    61:'Australia',
    62:'Indonesia',
    63:'Philippines',
    64:'New Zealand',
    65:'Singapore',
    66:'Thailand',
    81:'Japan',
    82:'South Korea',
    84:'Vietnam',
    86:'PR China',
    90:'Turkey',
    91:'India',
    92:'Pakistan',
    93:'Afghanistan',
    94:'Sri Lanka',
    95:'Burma (Myanmar)',
    212:'Morocco',
    213:'Algeria',
    216:'Tunisia',
    218:'Libya',
    220:'Gambia',
    221:'Senegal',
    222:'Mauritania',
    223:'Mali',
    224:'Guinea',
    225:"Côte d'Ivoire",
    226:'Burkina Farso',
    227:'Niger',
    228:'Togo',
    229:'Benin',
    230:'Mauritius',
    231:'Liberia',
    232:'Sierra Leone',
    233:'Ghana',
    234:'Nigeria',
    235:'Chad',
    236:'Central African Rep.',
    237:'Cameroon',
    238:'Cape Verde',
    239:'Sao Tome',
    240:'Equatorial Guinea',
    241:'Gabon',
    242:'Congo',
    243:'Zaire',
    244:'Angola',
    245:'Guinea-Bissau',
    246:'Diego Garcia',
    247:'Ascension Island',
    248:'Seychelles',
    249:'Sudan',
    250:'Rwanda',
    251:'Ethiopia',
    252:'Somalia',
    253:'Djibouti',
    254:'Kenya',
    255:'Tanzania',
    256:'Uganda',
    257:'Burundi',
    258:'Mozambique',
    259:'Zanzibar',
    260:'Zambia',
    261:'Madagascar',
    262:'Reunion Island',
    263:'Zimbabwe',
    264:'Namibia',
    265:'Malawi',
    266:'Lesotho',
    267:'Botswana',
    268:'Swaziland',
    269:'Comoros, Mayotte',
    290:'St. Helena',
    291:'Eritrea',
    297:'Aruba',
    298:'Faeroe Islands',
    299:'Green Island',
    350:'Gibraltar',
    351:'Portugal',
    352:'Luxembourg',
    353:'Ireland',
    354:'Iceland',
    355:'Albania',
    356:'Malta',
    357:'Cyprus',
    358:'Finland',
    359:'Bulgaria',
    370:'Lithuania',
    371:'Latvia',
    372:'Estonia',
    373:'Moldova',
    374:'Armenia',
    375:'Belarus',
    376:'Andorra',
    377:'Monaco',
    378:'San Marino',
    379:'Vatican City',
    380:'Ukraine',
    381:'Serbia',
    385:'Croatia',
    386:'Slovenia',
    387:'Bosnia, Herzegovina',
    389:'Macedonia',
    420:'Czech',
    421:'Slovak',
    423:'Liechtenstein',
    500:'Falkland Islands',
    501:'Belize',
    502:'Guatemala',
    503:'El Salvador',
    504:'Honduras',
    505:'Nicaragua',
    506:'Costa Rica',
    507:'Panama',
    508:'St. Pierre',
    509:'Haiti',
    590:'Guadeloupe',
    591:'Bolivia',
    592:'Guyana',
    593:'Ecuador',
    594:'French Guiana',
    595:'Paraguay',
    596:'Martinique',
    597:'Suriname',
    598:'Uruguay',
    599:'Netherlands Antilles',
    670:'East Timor',
    672:'Antarctica',
    673:'Brunei Darussalam',
    674:'Narupu',
    675:'Papua New Guinea',
    676:'Tonga',
    677:'Solomon Islands',
    678:'Vanuatu',
    679:'Fiji',
    680:'Palau',
    681:'Wallis and Futuna',
    682:'Cook Islands',
    683:'Niue Island',
    684:'American Samoa',
    685:'Western Samoa',
    686:'Kiribati',
    687:'New Caledonia',
    688:'Tuvalu',
    689:'French Polynesia',
    690:'Tokelau',
    691:'Micronesia',
    692:'Marshall Islands',
    850:'North Korea',
    852:'Hong Kong S.A.R.',
    853:'Macao S.A.R.',
    855:'Cambodia',
    856:'Laos',
    880:'Bangladesh',
    886:'Taiwan',
    960:'Maldives',
    961:'Lebanon',
    962:'Jordan',
    963:'Syria',
    964:'Iraq',
    965:'Kuwait',
    966:'Saudi Arabia',
    967:'Yemen',
    968:'Oman',
    970:'Palestine',
    971:'U.A.E.',
    972:'Israel',
    973:'Bahrain',
    974:'Qatar',
    975:'Bhutan',
    976:'Mongolia',
    977:'Nepal',
    981:'Iran',
    992:'Tajikistan',
    993:'Turkmenistan',
    994:'Azerbaijan',
    995:'Georgia',
    996:'Kyrgyzstan',
    998:'Uzbekistan',
}