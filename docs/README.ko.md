# node-xlrd
Microsoft Excel 파일에서 자료를 추출하기 위한 node.js 모듈

## 알림
* v1.x.x
  * status : [release candidates (v1.0.0-rc.x)](https://github.com/gutkyu/node-xlrd/tree/main_v1)
  * 새로운 특징과 개선 사항 추가
* v0.3.x
  * 안정 버전
  * 오류 수정

## 특징
*   python 모듈 xlrd(http://www.python-excel.org/)을 node.js환경에서 javascript으로 구현   

## 현재 상태
*   지원 파일 : Excel 2 ~ 2003 File(.xls)
*   셀 값 읽기만 가능
*   셀 포맷은 향후 지원 예정

## Node.js 지원 상황 (node-xlrd v0.3 기준)
*  v0.10 ~ v12 : 지원
*  v14 ~ 향후 버전 : maybe

## 변경 이력
### 0.3.9
* fixed date parsing error
	* when cellType is date, the value of month is incorrect
### 더 많은 내용
* [wiki Changelog](https://github.com/gutkyu/node-xlrd/wiki/Changelog)
	
## API 문서
* [node-xlrd 0.3](https://github.com/gutkyu/node-xlrd/wiki/API-v0.3)
* [node-xlrd 1.0.0 rc](https://github.com/gutkyu/node-xlrd/wiki/API-v1.0.0-rc)

## Install
* 최신 안정 버전 (현재 v0.3)
```console
npm i node-xlrd --save
```
* 최근 beta 버전 (현재 v1.0.0-rc)
```console
npm i node-xlrd@beta --save
```

## 사용 방법
```js
    var xl = require('node-xlrd');
    xl.open('./testDate.xls', function(err,bk){
        if(err) {console.log(err.name, err.message); return;}
        var sht = bk.sheetByIndex(0);
        console.log(sht.cell(0,0));
    });
```


## Usage
```js
var xl = require('../lib/node-xlrd');

xl.open('./basic.xls', showAllData);

function showAllData(err, bk){
  if (err) {
    console.log(err.name, err.message);
    return;
  }
  var shtCount = bk.sheet.count;
  for (var sIdx = 0; sIdx < shtCount; sIdx++) {
    var sht = bk.sheets[sIdx],
      rCount = sht.row.count,
      cCount = sht.column.count;
    for (var rIdx = 0; rIdx < rCount; rIdx++) {
      for (var cIdx = 0; cIdx < cCount; cIdx++) {
        try {
          console.log(
            '  cell : row = %d, col = %d, value = "%s"',
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
}
```
## 예제
* 더 많은 예제는 [GitHub examples folder](https://github.com/gutkyu/node-xlrd/tree/master/examples)에서 확인 가능.

## License
[BSD](https://github.com/gutkyu/node-xlrd/blob/master/LICENSE) license

