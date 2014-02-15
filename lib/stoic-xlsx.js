var XLSX = require('xlsx');

function parseXlsx(filePath, done) {
  var xlsx;
  try {
    var options = { cellNF: true, sheetStubs: false };
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
    processSheet(sheetName, sheet, function(e, result){
      if(e) { return processOneSheet(e); }
      sheets[sheetName] = result;
      processOneSheet(e);
    });
  };
  processOneSheet();
}

function processSheet(sheetName, sheet, done) {  
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
          if (isDateFormattedCell(val)) {
           values[columnIndex] = parseDateCode(val.v);
          } else if (val.v !== undefined) {
            values[columnIndex] = val.v;
          } else if (val.w) {
            values[columnIndex] = val.w;
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
          if (val.z) {
            // should we strip the text part of the number format?
            var m = val.z.match(/^(.*);@$/);
            if (m) {
              numberFormats[columnIndex] = m[1];
            } else {
              numberFormats[columnIndex] = val.z;
            }
          }
        }
      }
      if (values.length !== 0) { jsonSheet.values[rowIndex] = values; }
      if (numberFormats.length !== 0) { jsonSheet.numberFormats[rowIndex] = numberFormats; }
      if (notes.length !== 0) {
        if (notes[0] === undefined) {
          // Google Spreadsheet shifts the notes around.
          // would be good to know that indeed this was coming from a google spreadsheet.
          for (var kn = 0; kn < notes.length; kn++) {
            // we know the note we want is json/json5 on the first cell.
            var anote = notes[kn];
            if (anote) {
              anote = anote.trim();
              if (anote.charAt(0) === '{' && anote.charAt(anote.length - 1) === '}') {
                // looks like JSON or JSON5; do the shifting
                while(notes.length !== 0) {
                  var n = notes[0];
                  if (n === undefined) {
                    notes.shift();
                  } else {
                    break;
                  }
                }
                break;
              }
            }
          }
        }
        jsonSheet.notes[rowIndex] = notes;
      }
      if (formulas.length !== 0) { jsonSheet.formulas[rowIndex] = formulas; }
    }
  } catch(x) {
    err = x;
  }
  done(err, jsonSheet);
}

// http://office.microsoft.com/en-sg/excel-help/create-a-custom-number-format-HP010342372.aspx
var dateOrTimeNumberFormatSignatures = [
  'm', 'd', 'yy', 's', 'm', 'h'
];

function isDateFormattedCell(val) {
  if (val.z && typeof val.v === 'number') {
    return dateOrTimeNumberFormatSignatures.some(function(sig) {
      if (val.z.indexOf(sig) !== -1) {
        return true;
      }
    });
  } else {
    return false;
  }
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
  var coreProps = xlsx.Props;
  if (coreProps.creator && coreProps.creator.indexOf('@') !== -1) {
    header.owner = { email: custProps.creator };
  }
  if (coreProps.title) {
    header.name = coreProps.title;
  }

  var custProps = xlsx.Custprops;
  if (!custProps) { return header; }
  [ 'url', 'spid', 'locale', 'timeAtZero', 'timezone', 'datasource' ].forEach(function(prop) {
    if (custProps[prop]) {
      header[prop] = custProps[prop];
    }
  });
  if (custProps.folderUrl || custProps.folderId) {
    header.folder = { id: custProps.folderId, url: custProps.folderUrl };
  }
  return header;
}

exports.parseXlsx = parseXlsx;