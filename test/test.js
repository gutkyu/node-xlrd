var xl = require('../lib/node-xlrd'),
	util = require('util');

xl.open('./testDate.xls',{onDemend:true}, function(err,bk){
//xl.open('./test.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	console.log('print cells by sheet index')
	var shtCount = bk.sheet.count;
	for(var sIdx = 0; sIdx < shtCount; sIdx++ ){
		var sht = bk.sheet.byIndex(sIdx),
			rCount = sht.row.count,
			cCount = sht.column.count;
		console.log(util.format("  sheet %d: name = %s; rowCount = %d; columnCount = %d",sIdx, sht.name,rCount, cCount));
		for(var rIdx = 0; rIdx < rCount; rIdx++){
			for(var cIdx = 0; cIdx < cCount; cIdx++){
				try{
					console.log(rIdx,cIdx, sht.cell(rIdx,cIdx));
					console.log(util.format("    row = %d, col = %d, value = %s",rIdx,cIdx, sht.cell(rIdx,cIdx)));
				}catch(e){
					console.log(e.message);
				}
			}
		}
		if(bk.sheet.loaded(sIdx))
			bk.sheet.unload(sIdx);
	}
});

