# node-xlrd
Microsoft Excel 파일에서 자료를 추출하기 위한 node.js 모듈

## 특징
*   python 모듈 xlrd(http://www.python-excel.org/)을 node.js환경에서 javascript으로 구현   

## 현재 상태
*   지원 파일 : Excel 2 ~ 2003 File(.xls)
*   셀 값 읽기만 가능
*   셀 포맷은 향후 지원 예정

## 사용방법
```js
    var xl = require('node-xlrd');
    xl.open('./testDate.xls', function(err,bk){
        if(err) {console.log(err.name, err.message); return;}
        var sht = bk.sheetByIndex(0);
        console.log(sht.cell(0,0));
    });
```
