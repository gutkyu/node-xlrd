var assert = require('assert');
var xlrd = require('../lib/node-xlrd');

describe('Excel Test', function () {
  var bkInf = {};

  function loadFile(path) {
    return function (done) {
      xlrd.open(path, function (err, bk) {
        if (err) {
          return done(err);
        }
        bkInf = {};
        bkInf.biffVersion = bk.biffVersion;
        bkInf.lastUser = bk.lastUser;
        bkInf.codePage = bk.codePage;
        bkInf.encoding = bk.encoding;
        bkInf.countries = bk.countries;
        bkInf.sheetCount = bk.sheet.count;
        bkInf.sheets = [];
        bk.sheets.forEach(function (sh) {
          var shInf = {};
          shInf.name = sh.name;
          shInf.index = sh.index;
          shInf.rows = [];
          for (var r = 0; r < sh.row.count; r++) {
            var row = [];
            for (var c = 0; c < sh.column.count; c++) {
              row.push(sh.cell(r, c));
            }
            shInf.rows.push(row);
          }
          bkInf.sheets.push(shInf);
        });
        done();
      });
    };
  }

  describe('Excel97~2004 File Test', function () {
    before('Load File', loadFile('./test/test_excel97-2004.xls'));
    it('Workbook Test', function (done) {
      assert.strictEqual(bkInf.biffVersion, 80); //biff 80
      assert.strictEqual(bkInf.lastUser, 'Microsoft Office 사용자');
      assert.strictEqual(bkInf.codePage, 1200);
      assert.strictEqual(bkInf.encoding, 'utf16le');
      assert.ok(bkInf.countries[0] == 82 && bkInf.countries[1] == 82);
      assert.strictEqual(bkInf.sheetCount, 2);
      done();
    });
    it('Sheet Test', function (done) {
      var sh = bkInf.sheets[0];
      assert.strictEqual(sh.name, 'Sheet1');
      assert.strictEqual(sh.index, 0);
      assert.strictEqual(sh.rows[0][0], '');
      assert.strictEqual(sh.rows[1][0], ' ');
      assert.strictEqual(sh.rows[0][1], 1);
      assert.strictEqual(sh.rows[1][1], 3.14);
      assert.strictEqual(sh.rows[2][1], 2000000);
      assert.strictEqual(
        sh.rows[0][2].valueOf(),
        new Date(2020, 10, 4).valueOf() //js Date는 month가 zero-based, 11월 -> js month 10
      ); //date
      //assert.strictEqual(sh.rows[1][2].valueOf());//datetime
      //assert.strictEqual(sh.rows[2][2].valueOf());//time
      assert.strictEqual(sh.rows[0][3], 'Hello');
      assert.strictEqual(sh.rows[1][3], '안녕');
      assert.strictEqual(sh.rows[2][3], '你好');
      assert.strictEqual(sh.rows[3][3], 'こんにちは');
      assert.strictEqual(sh.rows[4][3], 'здраствуйте');
      assert.strictEqual(sh.rows[0][4], 1);
      assert.strictEqual(sh.rows[1][4], 0);
      done();
    });
  });
  return;
  describe('Excel5~95 File Test', function () {
    before('Load File', loadFile('./test/test_excel5-95.xls'));
    it('Workbook Test', function (done) {
      assert.strictEqual(bkInf.biffVersion, 70); //biff 70
      assert.strictEqual(bkInf.lastUser, 'Microsoft Office 사용자');
      assert.strictEqual(bkInf.codePage, 949);
      assert.strictEqual(bkInf.encoding, 'cp949');
      assert.ok(bkInf.countries[0] == 82 && bkInf.countries[1] == 82);
      assert.strictEqual(bkInf.sheetCount, 2);
      done();
    });
    it('Sheet Test', function (done) {
      var sh = bkInf.sheets[0];
      assert.strictEqual(sh.name, 'Sheet1');
      assert.strictEqual(sh.index, 0);
      assert.strictEqual(sh.rows[0][0], '');
      assert.strictEqual(sh.rows[1][0], ' ');
      assert.strictEqual(sh.rows[0][1], 1);
      assert.strictEqual(sh.rows[1][1], 3.14);
      assert.strictEqual(sh.rows[2][1], 2000000);
      assert.strictEqual(
        sh.rows[0][2].valueOf(),
        new Date('2020-11-03T15:00:00.000Z').valueOf()
      ); //date
      //assert.strictEqual(sh.rows[1][2].valueOf());//datetime
      //assert.strictEqual(sh.rows[2][2].valueOf());//time
      assert.strictEqual(sh.rows[0][3], 'Hello');
      assert.strictEqual(sh.rows[1][3], '안녕');
      assert.strictEqual(sh.rows[0][4], 1);
      assert.strictEqual(sh.rows[1][4], 0);
      done();
    });
  });
});
