var xl = require('../lib/node-xlrd');
	
xl.open('./test.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	
	var shtCount = bk.sheet.count;
	for(var sIdx = 0; sIdx < shtCount; sIdx++ ){
		console.log('sheet "%d" ', sIdx);
		console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
		var sht = bk.sheets[sIdx],
			rCount = sht.row.count,
			cCount = sht.column.count;
		console.log('  name = %s; index = %d; rowCount = %d; columnCount = %d', sht.name, sIdx, rCount, cCount);
		for(var rIdx = 0; rIdx < rCount; rIdx++){
			for(var cIdx = 0; cIdx < cCount; cIdx++){
				try{
					console.log('  cell : row = %d, col = %d, value = "%s"', rIdx, cIdx, sht.cell(rIdx,cIdx));
				}catch(e){
					console.log(e.message);
				}
			}
		}
		
		//save memory
		//console.log('  try unloading : index %d', sIdx );
		//if(bk.sheet.loaded(sIdx))
		//	bk.sheet.unload(sIdx);
		//console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
	}
	// if onDemand == false, allow function 'workbook.cleanUp()' to be omitted,
	// because it is called by caller 'node-xlrd.open()' after callback finished.
	//bk.cleanUp();
});

