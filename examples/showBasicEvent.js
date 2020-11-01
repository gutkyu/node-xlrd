var xl = require('../lib/node-xlrd');

var bk = xl.create('./basic.xls', {eventMode: true});
bk.on('done', function () {
  console.log('workbook done');
}).on('error', function (err) {//Workbook event "error" handler should be defined
  console.log('workbook error');
  console.log(err.stack);
}).work(function (err) {
  if(err) return console.log(err);
  bk.sheets.forEach(function (sh) {
    console.log('sheet "%s"', sh.name);
    sh.on('row', function (rowData) {
      var txts = [];
      rowData.values.forEach(function (val, cIdx) {
        txts.push('[' + cIdx + ']:' + val);
      });
      if (txts.length) {
        console.log('        ' + txts.join(', '));
      }
    }).on('done', function (rowCount) {
      console.log('row count : %d', rowCount);
    }).on('error', function (err) {
      console.log('sheet error');
      console.log(err.stack);
    });
  });
});
