var xl = require('../lib/node-xlrd');

xl.open('./basic.xls', showAllData);

function showAllData(err, bk) {
  if (err) {
    console.log(err.name, err.message);
    return;
  }
  bk.sheets.forEach(function (sht, sIdx) {
    var rCount = sht.row.count,
      cCount = sht.column.count;
    console.log('Sheet[%s]', sht.name);
    console.log(
      '  index : %d, row count : %d, column count : %d',
      sIdx,
      rCount,
      cCount
    );
    for (var rIdx = 0; rIdx < rCount; rIdx++) {
      for (var cIdx = 0; cIdx < cCount; cIdx++) {
        try {
          console.log(
            '    cell : row = %d, col = %d, value = "%s"',
            rIdx,
            cIdx,
            sht.cell(rIdx, cIdx)
          );
        } catch (e) {
          console.log(e.message);
        }
      }
    }

    //save memory
    //console.log('  try unloading : index %d', sIdx );
    //if(bk.sheet.loaded(sIdx))
    //	bk.sheet.unload(sIdx);
    //console.log('  check loaded : %s', bk.sheet.loaded(sIdx) );
  });
  // if onDemand == false, allow function 'workbook.cleanUp()' to be omitted,
  // because it is called by caller 'node-xlrd.open()' after callback finished.
  //bk.cleanUp();
}
