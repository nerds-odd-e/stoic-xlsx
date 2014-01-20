var expect = require('chai').expect;
var stoicXlsx = require('../lib/stoic-xlsx');
describe('When parsing the Travels spreadsheet', function() {
  var travels;
  before(function(done) {
    var filePath = 'test/assets/Travels.xlsx';
    stoicXlsx.parseXlsx(filePath, function(e, spreadsheet) {
      travels = spreadsheet.sheets;
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
  it('Must find a raw number format where expected', function() {
    var l = travels.Restaurants.values[0].length;
    var aDateNF = travels.Restaurants.numberFormats[1][l-5];
    expect(aDateNF).to.equal('m/d/yyyy h:mm:ss');
  });
});
describe('When parsing the SmallDateFormats spreadsheet', function() {
  var temporary, name;
  before(function(done) {
    var filePath = 'test/assets/SmallDateFormats.xlsx';
    stoicXlsx.parseXlsx(filePath, function(e, spreadsheet) {
      temporary = spreadsheet.sheets;
      name = spreadsheet.name;
      done();
    });
  });
  it('Must have found a name', function() {
    expect(name).to.equal('SmallDateFormats');    
  });
  it('Must have parsed 1 sheet', function() {
    expect(Object.keys(temporary)).to.deep.equal(['Sheet1']);    
  });
});
