# node-xlrd
node.js's module to extract data from Microsoft Excel™ File(.xls)

## Features
*  porting [python module xlrd](http://www.python-excel.org/) to javascript
*  pure javascript on node.js

## Status
*  supported file format : Excel 2 ~ 2003 File(.xls), not .xlsx file
*  only cell data without a formula, format.

## Changelog
### 0.2.4
* used lowerCamelCase for inner variables, properties and function names.  
    moved to javascript naming conventions from python.
* added 'toCountryName' function
* added 'lastUser' property
* fixed bugs

### 0.2.2
* fixed bugs

### 0.2.1  
* added 'onDemand' option
* add api
	workbook.cleanUp()
* fixed bugs

### 0.2.0
* changed api  

|Before | After|
|-----------|-----------|
|workbook.nsheets | workbook.sheet.count|
|workbook.codepage | workbook.codePage|
|workbook.sheets() function | workbook.sheets property|
|workbook.sheetByIndex() | workbook.sheet.byIndex()|
|workbook.sheetByName() | workbook.sheet.byName()|
|workbook.sheetNames() function | workbook.sheetNames property|
|workbook.sheetList | removed, recommand workbook.sheetNames property|
|workbook.getSheets() | removed, inner function|
|workbook.fakeGlobalsGetSheet() | removed, inner function|
|workbook.getBOF() | removed, inner function|
|workbook.datemode | workbook.dateMode|
|sheet.ncols | sheet.column.count|
|sheet.nrows | sheet.row.count|
|sheet.number | sheet.index|
|sheet.cellValue() | sheet.cell.getRaw()|
|sheet.cellObj() | removed|
|sheet.cellType() | sheet.cell.getType()|
|sheet.cellXFIndex() | sheet.cell.getXFIndex()|
|sheet.row() | sheet.row.getValues()|
|sheet.rowLength() | sheet.row.getCount()|
|sheet.rowTypes() | sheet.row.getTypes()|
|sheet.rowValues() | sheet.row.getValues()|

* fixed bugs

### 0.1.2  
* fixed bugs

### 0.1.0
* first commit
	
## API
This description is based on python xlrd function's comments and document.  
note : It is possible to use not yet announced api, but it  may be modified in future versions.

### node-xlrd
#### xlrd.open(file, options, callback)
open a workbook.
* file : file path  
* options :
	* onDemand :  option to load worksheets on demand.  
		it allows saving memory and time by loading only those selected sheets , and releasing sheets when no longer required.  
		* onDemand = false (default),  
			xlrd.open() loads global data and all sheets, releases resources no longer required  
		* onDemand = true (only BIFF version >= 5.0)  
			* xlrd.open() loads global data and returns without releasing resources. At this stage, the only information available about sheets is workbook.sheet.count and workbook.sheet.names.  
			* workbook.sheet.byName() and workbook.sheet.byIndex() will load or reload the requested sheet if it is not already loaded or unloaded .  
			* workbook.sheets will load all/any unloaded sheets.  
			* The caller may save memory by calling workbook.sheet.unload() when finished with the sheet. This applies irrespective of the state of onDemand.  
			* workbook.sheet.loaded() checks whether a sheet is loaded or not.  
			* workbook.cleanUp() should be called at end of node-xlrd.open() callback.  
	* callback : function(err, workbook)
    
###	node-xlrd.common
#### node-xlrd.common.toColumnName(colunmIndex)
convert column name to zero-based index.  

#### node-xlrd.common.toBiffVersionString(version)
return the corresponding string for BIFF version number

    0:  "(not BIFF)"
    20: "2.0"
    21: "2.1"
    30: "3"
    40: "4S"
    45: "4W"
    50: "5"
    70: "7"
    80: "8"
    85: "8X"

#### node-xlrd.common.toCountryName(countryCode)
returns the corresponding country name for a country code  
refer [workbook.countries](https://github.com/gutkyu/node-xlrd/blob/devel/Readme.md#workbookcountries) property

### Class : Workbook
#### workbook.biffVersion
Version of BIFF (Binary Interchange File Format) used to create the file.  
Latest is 8.0 (represented here as 80), introduced with Excel 97.  
Earliest supported by this module: 2.0 (represented as 20).

#### workbook.lastUser
the name of the user who last created, opened, or modified the file.

#### workbook.codePage
An integer denoting the character set used for strings in this file.  
For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF16LE.

#### workbook.encoding
The encoding that was derived from the codepage.
#### workbook.countries
Array containing the (telephone system) country code for:  

	[0]: the user-interface setting when the file was created.  
	[1]: the regional settings.  
Example:

	(1, 61) meaning (USA, Australia)  
	(82, 82) meaning (South Korea, South Korea)  
This information may give a clue to the correct encoding for an unknown codepage.  
For a long list of observed values, refer to the OpenOffice.org documentation for
the COUNTRY record or [List of country calling codes](http://en.wikipedia.org/wiki/List_of_country_calling_codes)

refer [node-xlrd.common.toCountryName(countryCode)](https://github.com/gutkyu/node-xlrd/blob/devel/Readme.md#node-xlrdcommontocolumnnamecolunmindex)

#### workbook.sheets
return : A list of all sheets in the book.  
All sheets not already loaded will be loaded.
	 
#### workbook.cleanUp()
all resources( file descriptor, large caches) released.  
Once cleanUp() called, no more reload or parse sheets.
	
* onDemand == true,  
	Call this function ,if don't neet to load or parse sheets.  
* onDemand == false (default),  
	It is possible to omit workbook.cleanUp(), because xlrd.open() implicit calls it after the callback finished.  
		
* if any error raised in xlrd.open callback, workbook.cleanUp() implicit call by xlrd.open()
			
#### workbook.sheet.count
Zero-based index of sheet in workbook.
	
#### workbook.sheet.names
return A list of the names of all the worksheets in the workbook file.  
This information is available even when no sheets have yet been loaded.
	
#### workbook.sheet.byIndex(sheetIndex)
sheetIndex : Sheet index  
return : An object of the sheet class

#### workbook.sheet.byName(sheetName)
sheetName : Name of sheet required  
return : An object of the sheet class

#### workbook.sheet.loaded(sheetId)
sheetId : sheet name or index  
return : true if sheet is loaded, false otherwise

#### workbook.sheet.unload(sheetId)
sheetId : sheet name or index to be unloaded.
	
### Class : Sheet
#### sheet.name
Name of sheet
	
#### sheet.index
Index of sheet
	
#### sheet.cell( rowIndex, colIndex )
return : javascript value ( string, number, Date) converting from raw of the cell in the given row and column.

#### sheet.cell.getValue( rowIndex, colIndex )
getValue() is equal to .cell()

#### sheet.cell.getRaw( rowIndex, colIndex )
return : Raw value of the cell in the given row and column.
	
#### sheet.cell.getType( rowIndex, colIndex )
return : Type of the cell in the given row and column.

#### sheet.cell.getXFIndex( rowIndex, colIndex )
do not use, because a cell style parsing not implemented yet.  
return : XF index of the cell in the given row and column.

#### sheet.row.count
return : Number of rows in sheet.
	
#### sheet.row.getValues( rowIndex )
return :  a sequence of the cell values in the given row.

#### sheet.row.getTypes( rowIndex, startColIndex, endColIndex )
return :  a slice of the types of the cells in the given row.

#### sheet.row.getRaws( rowIndex, startColIndex, endColIndex )
return :  a slice of the values of the cells in the given row.

#### sheet.col.count
Nominal number of columns in sheet. It is 1 + the maximum column index found, ignoring trailing empty cells. 
	
## Usage
```js
var xl = require('node-xlrd');

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
```
## License
BSD
