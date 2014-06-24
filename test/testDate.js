var util = require('util'),
    xl = require('../lib/node-xlrd');

xl.open('./testDate.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	var sh = bk.sheet.byIndex(0);
	console.log("Raw (%d, %d) : %d",0,0,sh.cell.getRaw(0,0)); // raw value
	console.log("Type(%d, %d) : %d",0,0,sh.cell.getType(0,0));
	console.log("Cell(%d, %d) : %s",0,0,sh.cell(0,0)); // converted to js data type
});
