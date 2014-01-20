var XLSX = require('xlsx');

function parseXlsx(filePath, options, done) {
  var xlsx;
  try {
    xlsx = XLSX.readFile(filePath, options);
  } catch(x) {
    return done(x);
  }
  var header = makeHeader(filePath, xlsx);
  var current = -1;
  var sheets = {};
  header.sheets = sheets;
  var names = Object.keys(xlsx.Sheets);
  var processOneSheet = function (e) {
    current++;
    if(e || current === names.length) {
      return done(e, header);
    }
    var sheetName = names[current];
    var sheet = xlsx.Sheets[sheetName];
    processSheet(sheet, function(e, result){
      if(e) { return processOneSheet(e); }
      sheets[sheetName] = result;
      processOneSheet(e);
    });
  };
  processOneSheet();
}

function processSheet(sheet, done) {  
  var stringify = function stringify(val) {
    switch(val.t){
      case 'n': return String(val.v);
      case 's': case 'str':
        if(typeof val.v === 'undefined') {
          return "";
        }
        return JSON.stringify(val.v);
      case 'b': return val.v ? "TRUE" : "FALSE";
      case 'e': return ""; /* throw out value in case of error */
      default: throw 'unrecognized type ' + val.t;
    }
  };
  var jsonSheet;
  var err;
  try {
    jsonSheet = { values: [], numberFormats: [], formulas: [], notes: [] };
    var sheetRawRange = sheet["!ref"];
    if(!sheetRawRange) {
      return done();
    }
    var range = XLSX.utils.decode_range(sheet["!ref"]);
    for(var rowIndex = range.s.r; rowIndex <= range.e.r; ++rowIndex) {
      var values = [];
      var numberFormats = [];
      var notes = [];
      var formulas = [];
      for(var columnIndex = range.s.c; columnIndex <= range.e.c; ++columnIndex) {
        var val = sheet[XLSX.utils.encode_cell({c:columnIndex,r:rowIndex})];
        if (val) {
          if (isDateFormattedCell(val) && val.raw) {
           values[columnIndex] = parseDateCode(val.raw);
          } else if (val.v !== undefined) {
            values[columnIndex] = val.v;
          } else if (val.raw) {
            values[columnIndex] = val.raw;
          }
          if (val.f) {
            formulas[columnIndex] = val.f;
          }
          if (val.c) {
            var note = makeNote(val.c);
            if (note) {
              notes[columnIndex] = note;
            }
          }
          if (val.rawnf) {
            numberFormats[columnIndex] = val.rawnf;
          }
        }
      }
      if (values.length !== 0) { jsonSheet.values[rowIndex] = values; }
      if (numberFormats.length !== 0) { jsonSheet.numberFormats[rowIndex] = numberFormats; }
      if (notes.length !== 0) { jsonSheet.notes[rowIndex] = notes; }
      if (formulas.length !== 0) { jsonSheet.formulas[rowIndex] = formulas; }
    }
  } catch(x) {
    err = x;
  }
  done(err, jsonSheet);
}

function isDateFormattedCell(val) {
  return val.rawnf && val.rawnf.indexOf('yyyy') !== -1 && 
    typeof val.raw === 'number' && typeof val.v === 'string';
}

// return the text of the first comment that contains a single text.
// when it contains 2 texts it is actually a comment, not a note (to be checked with MS Excel)
function makeNote(comments) {
  if (comments.length === 0) { return; }
  for (var i = 0; i < comments.length; i++) {
    var comment = comments[i];
    if (!comment.t) {
      continue;
    }
    return comment.t;
  }
}

/**
  See MSDN doc for the numbers: search the web for 'OLE Automation date'
  http://stackoverflow.com/questions/10443325/how-to-convert-ole-automation-date-to-readable-format-using-javascript.

  The reverse: var oaDate = (date - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
 */
function parseDateCode(oaDate) {
  var date = new Date();
  date.setTime((oaDate - 25569) * 24 * 3600 * 1000);
  return date;
}

function makeHeader(filePath, xlsx) {
  var nameMatch = filePath.match(/([^\/]*)$/);
  var name = filePath;
  if (nameMatch && nameMatch[1]) {
    var toks = nameMatch[1].split('.');
    toks.pop();
    name = toks.join('.');
  }
  var header = {
    spid: filePath,
    name: name
  };

  return header;
}

exports.parseXlsx = parseXlsx;