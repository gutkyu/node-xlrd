var xl = require('../lib/node-xlrd'),
	util = require('util');
xl.open('./testDate.xls',{onDemand:true}, function(err,bk){
//xl.open('./test.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	
	console.log('---------- 1. load all sheets, print cells ----------')
	
	console.log('---------- 2. load sheet, print cells, unload it by sheet index ----------')
	
	var shtCount = bk.sheet.count;
	for(var sIdx = 0; sIdx < shtCount; sIdx++ ){
		console.log('sheet "%d" ', sIdx);
		console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
		console.log('  try loading : index %d', sIdx );
		var sht = bk.sheet.byIndex(sIdx),
			rCount = sht.row.count,
			cCount = sht.column.count;
		console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
		console.log('  name = %s; index = %d; rowCount = %d; columnCount = %d', sht.name, sIdx, rCount, cCount);
		for(var rIdx = 0; rIdx < rCount; rIdx++){
			for(var cIdx = 0; cIdx < cCount; cIdx++){
				try{
					console.log('  cell : row = %d, col = %d, value = %s', rIdx, cIdx, sht.cell(rIdx,cIdx));
				}catch(e){
					console.log(e.message);
				}
			}
		}
		console.log('  try unloading : index %d', sIdx );
		if(bk.sheet.loaded(sIdx))
			bk.sheet.unload(sIdx);
		console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
	}
	
	console.log('---------- 3. load sheet, print cells and unload it by sheet name ----------')
	bk.sheet.names.forEach(function(shtNm){
		console.log('sheet "%s" ', shtNm);
		console.log('  check loaded : %s', bk.sheet.loaded(shtNm) );
		console.log('  try loading : name %s', shtNm );
		var sht = bk.sheet.byName(shtNm),
			rCount = sht.row.count,
			cCount = sht.column.count;
		console.log('  check loaded : %s', bk.sheet.loaded(shtNm) );
		console.log('  name = %s; index = %d; rowCount = %d; columnCount = %d', sht.name, sIdx, rCount, cCount);
		for(var rIdx = 0; rIdx < rCount; rIdx++){
			for(var cIdx = 0; cIdx < cCount; cIdx++){
				try{
					console.log('  cell : row = %d, col = %d, value = %s', rIdx, cIdx, sht.cell(rIdx,cIdx));
				}catch(e){
					console.log(e.message);
				}
			}
		}
		console.log('  try unloading : name %s', shtNm );
		if(bk.sheet.loaded(shtNm))
			bk.sheet.unload(shtNm);
		console.log('  check loaded : %s', bk.sheet.loaded(shtNm) );
	});
	bk.cleanUp();// if onDemand == true, function workbook.cleanUp() should be called.
});

