var expect = require('chai').expect;
var stoicXlsx = require('../lib/stoic-xlsx');
describe('When parsing the Travels spreadsheet', function() {
  var travels, name, spid, url, folder;
  before(function(done) {
    var filePath = 'test/assets/Travels.xlsx';
    stoicXlsx.parseXlsx(filePath, function(e, spreadsheet) {
      travels = spreadsheet.sheets;
      name = spreadsheet.name;
      spid = spreadsheet.spid;
      url = spreadsheet.url;
      folder = spreadsheet.folder;
      done();
    });
  });
  it('Must have parsed 6 sheets', function() {
    expect(Object.keys(travels)).to.deep.equal(['Applications', 'Companies', 'Fields', 'Hotels', 'Objects','Restaurants']);    
  });
  it('Must find a javascript date where expected', function() {
    var l = travels.Restaurants.values[0].length;
    var aDate = travels.Restaurants.values[1][l-6];
    expect(aDate).to.be.instanceOf(Date);
    expect(isNaN(aDate.valueOf())).to.equal(false);
  });
  it('Must find a raw number format where expected', function() {
    var l = travels.Restaurants.values[0].length;
    var aDateNF = travels.Restaurants.numberFormats[1][l-6];
    expect(aDateNF).to.equal('M/d/yyyy H:mm:ss');
  });
  it('Must have read the custom properties', function() {
    expect(name).to.equal('Travels');
    expect(spid).to.equal('0Ah3dPnCQn7k4dHNzY3JidkdnV1liQzdrVnB4WjZJdUE');
    expect(url).to.equal('https://docs.google.com/a/sutoiku.com/spreadsheet/ccc?key=tsscrbvGgWYbC7kVpxZ6IuA');
    expect(folder).to.deep.equal({
      id: '0Bx3dPnCQn7k4N2tTNUk5bVVyVFk',
      url: 'https://docs.google.com/a/sutoiku.com/open?id=0Bx3dPnCQn7k4N2tTNUk5bVVyVFk' });
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
