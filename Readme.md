# node-xlrd
node.js's module to extract data from Microsoft Excel™ File(.xls)

## Features
*  porting [python module xlrd](http://www.python-excel.org/) to javascript
*  pure javascript on node.js

## Status
*  supported file format : Excel 2 ~ 2003 File(.xls), not .xlsx file
*  only cell data without a formula, format, hyperlink.
*  new features and enhancements in node-xlrd version 1
  * event-based programming
    * minimizing memory usage
    * large file support

## Supported node.js versions
*  from v0.10 to v6 : no supported
*  from v7 to v15 : supported
*  future versions : maybe

## Changelog
### 1.0.0-rc.2
* fixed date parsing errors

### More Changelog
* [wiki Changelog](https://github.com/gutkyu/node-xlrd/wiki/Changelog)
	
## API Reference
The API reference documentation provides detailed information about a function or object in node-xlrd.js.
* [node-xlrd 0.3](https://github.com/gutkyu/node-xlrd/wiki/API-v0.3)
* [node-xlrd 1.0.0 rc](https://github.com/gutkyu/node-xlrd/wiki/API-v1.0.0-rc)

## Installation
* Latest stable release(curent v0.3)
```console
npm i node-xlrd --save
```
* Latest beta release (current v1.0.0-rc)
```console
npm i node-xlrd@beta --save
```

## Usage
### Event Mode : False
```js
var xl = require('node-xlrd');

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
  }
});
```

### Event Mode : True
```js
var xl = require('node-xlrd');

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
```
## Examples
* See [GitHub examples folder](https://github.com/gutkyu/node-xlrd/tree/master/examples) for more examples.

## License
This project is licensed under the [BSD](https://github.com/gutkyu/node-xlrd/blob/master/LICENSE) license.

