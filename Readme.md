# node-xlrd
node.js's module to extract data from Microsoft Excel File

## Features
*  based on the python module xlrd http://www.python-excel.org/ 
*  pure javascript on node.js

##  Status
*  supported file format : Excel 2 ~ 2003 File(.xls)
*  only cell data

## Usage

    xl.open('./testDate.xls', function(err,bk){
        if(err) {console.log(err.name, err.message); return;}
        var sht = bk.sheetByIndex(0);
        console.log(sht.cell(0,0));
    });
