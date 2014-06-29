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
	for( shx = 0 ;shx <  bk.sheet.count; shx++){
		var sh = bk.sheet.byIndex(shx);
		var rowCount = sh.row.count, colCount = sh.column.count;
		
		var anshow = Math.min(nshow, rowCount);
		console.log(util.format("sheet %d: name = %s; rowCount = %d; colCount = %d",shx, sh.name,rowCount, colCount));
		if (rowCount && colCount){
			// Beat the bounds
			for(rowx = 0 ; rowx < rowCount ; rowx++){
				var nc = sh.row.getCount(rowx);
				if (nc){
					var _junk = sh.row.getTypes(rowx)[nc-1];
					_junk = sh.row.getValues(rowx)[nc-1];
					_junk = sh.cell(rowx, nc-1);
				}
			}
		}
		for(rowx = 0; rowx <anshow-1; rowx++){
			if (! printit && rowx % 10000 == 1 && rowx > 1)
				console.log(util.format("done %d rows" ,rowx-1));
			showRow(bk, sh, rowx, colCount, printit);
		}
		if (anshow && rowCount)
			showRow(bk, sh, rowCount-1, colCount, printit);
	}
	
}

function showHeader(bk){
	console.log(util.format("BIFF version: %s; dateMode: %s",xl.common.toBiffVersionString(bk.biffVersion), bk.dateMode));
	console.log(util.format("codePage: %s (encoding: %s); countries: [%s,%s], %s",bk.codePage, bk.encoding, xl.common.toCountryName(bk.countries[0]), xl.common.toCountryName(bk.countries[1]), bk.countries));
	console.log(util.format("Last saved by: %s",bk.lastUser));
	console.log(util.format("Number of data sheets: %d" ,bk.sheet.count));
	console.log(util.format("Ragged rows: %d" , bk.raggedRows));
	if (bk.formattingInfo)
		console.log(util.format("FORMATs: %d, FONTs: %d, XFs: %d",len(bk.formatList), len(bk.fontList), bk.xfList.length));
	//if (! options.suppress_timing)
	//	console.log(util.format("Load time: %.2f seconds (stage 1) %.2f seconds (stage 2)",bk.loadTimeStage1, bk.loadTimeStage2));
	console.log();
}

function showRow(bk, sh, rowx, colLen, printit){
        if (bk.raggedRows)
            colLen = sh.row.getCount(rowx).length;
        if (!colLen) return;
        //if (printit) ;
        if (bk.formattingInfo)
			getRowData(bk, sh, rowx, colLen).forEach(function(x){
				var colx=x[0], typ=x[1], val=x[2], raw = x[3], cxfx=x[4];
				if (printit)
                    console.log(util.format("cell %s%d: type=%d, data: %s, raw data: %s, xfx: %s",
                         xl.toColumnName(colx), rowx+1, typ, val, raw, cxfx));
			});
        else
			getRowData(bk, sh, rowx, colLen).forEach(function(x){
				var colx=x[0], typ=x[1], val=x[2], raw = x[3], _unused=x[3];
				if (printit)
                    console.log(util.format("cell %s%d: type=%d, data: %s, raw data: %s", xl.common.toColumnName(colx), rowx+1, typ, val, raw));
			});
}

function getRowData(bk, sh, rowx, colLen){
	var result = [];
	var dmode = bk.dateMode;
	var ctys = sh.row.getTypes(rowx);
	var cvals = sh.row.getValues(rowx);
	var craws = sh.row.getRaws(rowx);
	for(colx = 0; colx <colLen; colx++){
		var cty = ctys[colx];
		var cval = cvals[colx];
		var craw = craws[colx];
		var cxfx = null;
		if (bk.formattingInfo)
			cxfx = str(sh.cell.getXFIndex(rowx, colx));
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
			showval = comm.errorCodeStringMap[cval]; 
			if(showval == undefined) showval = util.format('<Unknown error code 0x%02x>',cval);
		}else*/
			showval = cval;
		result.push([colx, cty, showval, craw, cxfx]);
	}
	return result
}