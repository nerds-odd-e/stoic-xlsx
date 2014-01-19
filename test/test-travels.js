var expect = require('chai').expect;
var stoicXlsx = require('../lib/stoic-xlsx');
describe('When parsing the Travels spreadsheet', function() {
  var travels;
  before(function(done) {
    var filePath = 'test/assets/Travels.xlsx';
    stoicXlsx.parseXlsx(filePath, {evaluateFmt: false, skipEmptyCells: true}, function(e, sheets) {
      travels = sheets;
      done();
    });
  });
  it('does it', function() {

  });
});