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
	var values = [];
	var sheetRawRange = sheet["!ref"];
	if(!sheetRawRange) {
		return done();
	}
	var range = XLSX.utils.decode_range(sheet["!ref"]);
	for(var rowIndex = range.s.r; rowIndex <= range.e.r; ++rowIndex) {
		var row = [];
		for(var columnIndex = range.s.c; columnIndex <= range.e.c; ++columnIndex) {
			var val = sheet[XLSX.utils.encode_cell({c:columnIndex,r:rowIndex})];
			if (val) {
				val = stringify(val).replace(/\\r\\n/g,"\n").replace(/\\t/g,"\t")
														.replace(/\\\\/g,"\\").replace("\\\"","\"\"");
				row[columnIndex] = val;
			}
		}
		if (row.length !== 0) {
			values[rowIndex] = row;
		}
	}
	done(null, jsonSheet);
}


exports.parseXlsx = parseXlsx;