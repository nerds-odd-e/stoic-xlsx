var expect = require('chai').expect;
var stoicXlsx = require('../lib/stoic-xlsx');
describe('When parsing the Travels spreadsheet', function() {
  var travels;
  before(function(done) {
    var filePath = 'test/assets/Travels.xlsx';
    stoicXlsx.parseXlsx(filePath, {evaluateFmt: true, skipEmptyCells: true, skipRawnf: false}, function(e, sheets) {
      travels = sheets;
      done();
    });
  });
  it('Must have parsed 6 sheets', function() {
    expect(Object.keys(travels)).to.deep.equal(['Applications', 'Companies', 'Fields', 'Hotels', 'Objects','Restaurants']);    
  });
  it('Must find a javascript date where expected', function() {
    var l = travels.Restaurants.values[0].length;
    var aDate = travels.Restaurants.values[1][l-5];
    expect(aDate).to.be.instanceOf(Date);
  });
});