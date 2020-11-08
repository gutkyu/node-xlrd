# node-xlrd
node.js's module to extract data from Microsoft Excel™ File(.xls)

## Announcements
* v1.x.x
  * status : [release candidates (v1.0.0-rc.x)](https://github.com/gutkyu/node-xlrd/tree/main_v1)
  * new features and enhancements
* v0.3.x
  * stable version
  * working on bug-fixes

## Features
*  porting [python module xlrd](http://www.python-excel.org/) to javascript
*  pure javascript on node.js

## Status
*  supported file format : Excel 2 ~ 2003 File(.xls), not .xlsx file
*  only cell data without a formula, format, hyperlink.

## Supported Node.js Versions
*  from v0.10 to v12 : supported
*  from v14 to future versions : maybe

## Changelog
### 0.3.9
* fixed date parsing error
	* when cellType is date, the value of month is incorrect
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
```js
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
  });
}
```
## Examples
* See [GitHub examples folder](https://github.com/gutkyu/node-xlrd/tree/master/examples) for more examples.

## License
This project is licensed under the [BSD](https://github.com/gutkyu/node-xlrd/blob/master/LICENSE) license.

