var util = require('util'),
    xl = require('../lib/node-xlrd');

xl.open('./testDate.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
	var sh = bk.sheetByIndex(0);
	console.log("cell(",0,0,"')",sh.cellValue(0,0)); // raw value
	console.log("cellType(",0,0,"')",sh.cellType(0,0));
	console.log("cell(",0,0,"')",sh.cell(0,0)); // converted to js data type
});
