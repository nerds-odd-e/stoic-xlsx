var XLSX = require('xlsx');

function parseXlsx(filePath, done) {
  var xlsx = XLSX.readFile(filePath);
  var current = -1;
  var sheets = {};
  var processOneSheet = function (e) {
    current++;
    if(e || current === xlsx.SheetNames.length) {
      return done(e);
    }
    var sheetName = xlsx.SheetNames[current];
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
  var jsonSheet = { values: [], numberFormat: [], formulas: [], notes: [] };
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
        if (val.r) {
          if (val.v) {
            values = val.v;
          }
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
        if (val.n) {
          numberFormats[columnIndex] = val.n;
        }
        // val = stringify(val).replace(/\\r\\n/g,"\n").replace(/\\t/g,"\t")
        //                     .replace(/\\\\/g,"\\").replace("\\\"","\"\"");
        values[columnIndex] = val;
      }
    }
    if (values.length !== 0) { jsonSheet.values[rowIndex] = values; }
    if (numberFormats.length !== 0) { jsonSheet.numberFormats[rowIndex] = numberFormats; }
    if (notes.length !== 0) { jsonSheet.notes[rowIndex] = notes; }
    if (formulas.length !== 0) { jsonSheet.formulas[rowIndex] = formulas; }
  }
  done(null, jsonSheet);
}

// return the text of the first comment that contains a single text.
// when it contains 2 texts it is actually a comment, not a note (to be checked with MS Excel)
function makeNote(comments) {
  if (comments.length === 0) { return; }
  for (var i = 0; i < comments.length; i++) {
    var comment = comments[i];
    if (!comment.t || comment.t.length !== 1) {
      continue;
    }
    return comment.t[0];
  }
}


exports.parseXlsx = parseXlsx;