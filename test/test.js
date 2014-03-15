var xl = require('../lib/node-xlrd');

xl.open('./testDate.xls', function(err,bk){
	if(err) {console.log(err.name, err.message); return;}
    var sht = bk.sheetByIndex(0);
	console.log(sht.cell(0,0));
});

