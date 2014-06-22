# node-xlrd
node.js's module to extract data from Microsoft Excel¢â File

## Features
*  porting [python module xlrd](http://www.python-excel.org/) to javascript
*  pure javascript on node.js

## Status
*  supported file format : Excel 2 ~ 2003 File(.xls), not .xlsx file
*  only cell data without a formula, format.

## Changelog
### 0.2.0
	*changed api
		workbook.nsheets -> workbook.sheet.count
		workbook.codepage -> workbook.codePage
		workbook.sheets() function -> workbook.sheets property
		workbook.sheetByName() -> workbook.sheet.byName()
		workbook.sheetNames() function -> workbook.sheetNames property
		workbook.sheetList -> removed
		workbook.getSheets() -> removed
		workbook.fakeGlobalsGetSheet() -> removed
		workbook.getBOF() -> removed
		workbook.datemode -> workbook.dateMode
		sheet.ncols -> sheet.column.count
		sheet.nrows -> sheet.row.count
		sheet.number -> sheet.index
		sheet.cellValue() -> sheet.cell.getRaw()
		sheet.cellObj() -> removed
		sheet.cellType() -> sheet.cell.getType()
		sheet.cellXFIndex() -> sheet.cell.getXFIndex()
		sheet.row() -> sheet.row.getValues()
		sheet.rowLength() -> sheet.row.
		sheet.rowTypes() -> sheet.row.getTypes()
		sheet.rowValues() -> sheet.row.getValues()
		
	*fixed bugs
	
### 0.1.2
	*fixed bugs
### 0.1.0
	*first commit
	
## Api
	This description is based on python xlrd function's comments and document.
	note : It is possible to use not yet announced api, but it  may be modified in future versions.

### node-xlrd
	*xlrd.openWorkbook(file, options, callback)
		*file : excel file path
		*options :
			*onDemand :  option to load worksheets on demand.
				it allows saving memory and time by loading only those selected sheets , and releasing sheets when no longer required.
				onDemand = false (default)
					open_workbook() loads global data and all sheets, releases resources no longer required
				onDemand = true (only BIFF version >= 5.0)
					openWorkbook() loads global data and returns without releasing resources. At this stage, the only information available about sheets is workbook.sheet.count and workbook.sheet.names.
					workbook.sheet.byName() and workbook.sheet.byIndex() will load or reload the requested sheet if it is not already loaded or unloaded .
					workbook.sheets will load all/any unloaded sheets.
					The caller may save memory by calling workbook.sheet.unload() when finished with the sheet. This applies irrespective of the state of onDemand.
					workbook.sheet.loaded() checks whether a sheet is loadd or not.
		*callback : function(err, workbook)
		
### Class : Workbook
	#### workbook.biffVersion
		Version of BIFF (Binary Interchange File Format) used to create the file.
		Latest is 8.0 (represented here as 80), introduced with Excel 97.
		Earliest supported by this module: 2.0 (represented as 20).
		
	#### workbook.codePage
		An integer denoting the character set used for strings in this file.
		For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.

	#### workbook.encoding
		The encoding that was derived from the codepage.
	#### workbook.countries
		A array containing the (telephone system) country code for:  
			[0]: the user-interface setting when the file was created.  
			[1]: the regional settings.  
		Example: 
			(1, 61) meaning (USA, Australia)
			(82, 82) meaning (South Korea, South Korea)
		This information may give a clue to the correct encoding for an unknown codepage.
		For a long list of observed values, refer to the OpenOffice.org documentation for
		the COUNTRY record or [List of country calling codes](http://en.wikipedia.org/wiki/List_of_country_calling_codes)

	#### workbook.sheets
		return A list of all sheets in the book.
		All sheets not already loaded will be loaded.
	
	#### workbook.sheet.count
		zero-based index of sheet in workbook.
		
	*workbook.sheet.names
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
		
	#### sheet.cell( rowIndex, colIndex)
		return : javascript value ( string, number, Date) converting from raw of the cell in the given row and column.
	
	#### sheet.cell.getValue( rowIndex, colIndex)
		getValue() is equal to .cell()
	
	#### sheet.cell.getRaw( rowIndex, colIndex)
		return : Raw value of the cell in the given row and column.
		
	#### sheet.cell.getType( rowIndex, colIndex)
		return : Type of the cell in the given row and column.
	
	#### sheet.cell.getXFIndex( rowIndex, colIndex)
		do not use, because a cell style parsing not yet implemented 
		return : XF index of the cell in the given row and column.
	
	#### sheet.row.count
		return : Number of rows in sheet.
		
	#### sheet.row.getValues( rowIndex)
		return :  a sequence of the cell values in the given row.
	
	#### sheet.row.getTypes( rowIndex, startColIndex, endColIndex)
		return :  a slice of the types of the cells in the given row.
	
	#### sheet.row.getRaws( rowIndex, startColIndex, endColIndex)
		return :  a slice of the values of the cells in the given row.
	
	#### sheet.col.count
		Nominal number of columns in sheet. It is 1 + the maximum column index
		found, ignoring trailing empty cells. 
	
## Usage
```js
    var xl = require('node-xlrd');
    xl.open('./testDate.xls', function(err,bk){
        if(err) {console.log(err.name, err.message); return;}
        var sht = bk.sheet.byIndex(0);
        console.log(sht.cell(0,0));
    });
```
## License
BSD