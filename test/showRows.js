var util = require('util'),
    xl = require('../lib/node-xlrd');

xl.open('./test.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	show(bk,10);
});

function show(bk, nshow, printit){
	var nshow= nshow ||65535;
	var printit= printit||1;
	
	showHeader(bk);
	/*
	if (0){
		var rclist = xlrd.sheet.rc_stats.items();
		rclist = sorted(rclist)
		print("rc stats")
		for k, v in rclist:
			print("0x%04x %7d" % (k, v))
	}
	*/
	for( shx = 0 ;shx <  bk.nsheets; shx++){
		var sh = bk.sheetByIndex(shx);
		var nrows = sh.nrows, ncols = sh.ncols;
		
		var anshow = Math.min(nshow, nrows);
		console.log(util.format("sheet %d: name = %s; nrows = %d; ncols = %d",shx, sh.name, sh.nrows, sh.ncols));
		if (nrows && ncols){
			// Beat the bounds
			for(rowx = 0 ; rowx < nrows ; rowx++){
				var nc = sh.rowLength(rowx);
				if (nc){
					var _junk = sh.rowTypes(rowx)[nc-1];
					_junk = sh.rowValues(rowx)[nc-1];
					_junk = sh.cell(rowx, nc-1);
				}
			}
		}
		for(rowx = 0; rowx <anshow-1; rowx++){
			if (! printit && rowx % 10000 == 1 && rowx > 1)
				console.log(util.format("done %d rows" ,rowx-1));
			showRow(bk, sh, rowx, ncols, printit);
		}
		if (anshow && nrows)
			showRow(bk, sh, nrows-1, ncols, printit);
	}
	
}

function showHeader(bk){
	console.log(util.format("BIFF version: %s; datemode: %s",xl.toBiffVersionText(bk.biffVersion), bk.datemode));
	console.log(util.format("codepage: %s (encoding: %s); countries: %s",bk.codepage, bk.encoding, bk.countries));
	console.log(util.format("Last saved by: %s",bk.user_name));
	console.log(util.format("Number of data sheets: %d" ,bk.nsheets));
	console.log(util.format("Ragged rows: %d" , bk.ragged_rows));
	if (bk.formatting_info)
		console.log(util.format("FORMATs: %d, FONTs: %d, XFs: %d",len(bk.format_list), len(bk.font_list), bk.xf_list.length));
	//if (! options.suppress_timing)
	//	console.log(util.format("Load time: %.2f seconds (stage 1) %.2f seconds (stage 2)",bk.load_time_stage_1, bk.load_time_stage_2));
	console.log();
}

function showRow(bk, sh, rowx, colLen, printit){
        if (bk.ragged_rows)
            colLen = sh.rowLength(rowx).length;
        if (!colLen) return;
        //if (printit) ;
        if (bk.formatting_info)
			getRowData(bk, sh, rowx, colLen).forEach(function(x){
				var colx=x[0], ty=x[1], val=x[2], cxfx=x[3];
				if (printit)
                    console.log(util.format("cell %s%d: type=%d, data: %s, xfx: %s",
                         xl.colname(colx), rowx+1, ty, val, cxfx));
			});
        else
			getRowData(bk, sh, rowx, colLen).forEach(function(x){
				var colx=x[0], ty=x[1], val=x[2], _unused=x[3];
				if (printit)
                    console.log(util.format("cell %s%d: type=%d, data: %s" ,xl.colname(colx), rowx+1, ty, val));
			});
}

function getRowData(bk, sh, rowx, colLen){
	var result = [];
	var dmode = bk.datemode;
	var ctys = sh.rowTypes(rowx);
	var cvals = sh.rowValues(rowx);
	for(colx = 0; colx <colLen; colx++){
		var cty = ctys[colx];
		var cval = cvals[colx];
		var cxfx = null;
		if (bk.formatting_info)
			cxfx = str(sh.cellXFIndex(rowx, colx));
		else
			cxfx = '';
		var showval = '';
		/* if (cty == comm.XL_CELL_DATE){
			//try:
				showval = comm.xldate_as_tuple(cval, dmode);
			//except xlrd.XLDateError as e:
			//	showval = "%s:%s" % (type(e).__name__, e)
			//	cty = xlrd.XL_CELL_ERROR
		}else if( cty == comm.XL_CELL_ERROR){
			showval = comm.error_text_from_code[cval]; 
			if(showval == undefined) showval = util.format('<Unknown error code 0x%02x>',cval);
		}else*/
			showval = cval;
		result.push([colx, cty, showval, cxfx]);
	}
	return result
}