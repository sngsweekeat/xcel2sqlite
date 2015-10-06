var express = require('express');
var router = express.Router();

var excel = require('xlsx');
var sql = require('sql.js');
var multer = require('multer');
var upload = multer({
	dest: "/tmp/",
	fileFilter: hasExcelExtension,
	limits: {
		fileSize: 409600
	}
});
var request = require('request');
var fs = require("fs");

router.post('/', upload.single('excelFile'), function(req, res, next) {
	var recaptchaResponse = req.body['g-recaptcha-response'];
	var recaptchaSecret = '6LdCJg4TAAAAAMbxGMBwha_dYCgfIe_bWOnYceTu';

	// Verify recaptcha
	request.post({
		url: 'https://www.google.com/recaptcha/api/siteverify',
		form: {
			secret: recaptchaSecret,
			response: recaptchaResponse
		}
	}, function(err, httpResponse, body) {
		var result = JSON.parse(body);
		if (result.success) {
			// Recaptcha verified
			var file = req.file;
			if (file) {
				console.log('Received file ' + file.originalname + " of size " + file.size + " to " + file.path);
				//
				// Get worksheets
				var workbook = excel.readFile(file.path);
				var worksheetNames = workbook.SheetNames;
				var db = new sql.Database();
				console.log('Found ' + worksheetNames.length + ' worksheets: ' + worksheetNames);

				for (var i = 0; i < worksheetNames.length; i++) {
					var worksheet = workbook.Sheets[worksheetNames[i]];
					console.log('Processing worksheet ' + worksheetNames[i]);
					var colNames = getColumnNamesFromSheet(worksheet);
					// Create table for worksheet
					createTableInDb(db, worksheetNames[i], colNames);
					var row = 2, col = 1;
					var rowVals = getExcelDataRow(worksheet, row, colNames.length);
					while (!isBlankRow(rowVals)) {
						insertDataIntoTableOfDb(db, worksheetNames[i], colNames, rowVals);
						row++;
						rowVals = getExcelDataRow(worksheet, row, colNames.length);
					}
				}
				var dbFileName = file.filename + '.sqlite';
				var dbFilePath = file.destination + '/' + dbFileName;
				saveDBToDisk(db, dbFilePath);
				db.close();
				res.attachment(dbFileName);
				res.sendFile(dbFilePath);
			} else {
				console.log('No valid upload received');
				res.send('No valid upload received. Only .XLS/.XLSX files up to 400KB are accepted.');
			}
		} else {
			res.send('Error: Are you a robot?');ls -l /tmp
		}
	});
});

module.exports = router;

function sanitize(s) {
	return s.replace(/;/, ''); // strip semi-colons
}

// Strip semi-colons, replace '.' with '_'
function sanitizeTableOrColumnNames(s) {
	return s.replace(/;/, '').replace(/\./, '_');
}

function hasExcelExtension(req, file, callback) {
	var regex = /.*\.(xlsx|xls)/;
	var matches = file.originalname.match(regex);
	if (matches !== null)
		callback(null, true);
	else
		callback(null, false);
}

function getColumnNamesFromSheet(sheet) {
	var row = 1, col = 1;
	var colNames = [];

	var colName = excelValueAtCellInSheet(sheet, excelAddressForCell(row, col));
	while (colName && colName != '') {
		colName = colName.replace(/^\s+|\s+$/g, ''); // Trim leading/trailing space
		if (colName != '') {
			colNames[col - 1] = colName;
			console.log("	Found a column " + colName);
		}
		col++;
		colName = excelValueAtCellInSheet(sheet, excelAddressForCell(row, col));
	}
	return colNames;
}

function createTableInDb(db, tableName, columns) {
	var command = "CREATE TABLE :tableName (:colNames)";

	command = command.replace(/:tableName/g, sanitizeTableOrColumnNames(tableName));
	var colNames = '';
	for (var i = 0; i < columns.length; i++) {
		colNames += sanitizeTableOrColumnNames(columns[i]);
		if (i != (columns.length - 1))
			colNames += ',';
	}
	command = command.replace(/:colNames/, colNames);
	console.log(command);
	db.run(command, null);
}

function getExcelDataRow(sheet, row, numCols) {
	var rowVals = [];
	for (var col = 1; col <= numCols; col++) {
		var val = excelValueAtCellInSheet(sheet, excelAddressForCell(row, col));
		if (!val)
			val = '';
		rowVals[col - 1] = val;
	}
	return rowVals;
}

function isBlankRow(rowData) {
	var isBlank = true;
	for (var i = 0; i < rowData.length; i++) {
		if (rowData[i] && rowData[i].length > 0) {
			isBlank = false;
			break;
		}
	}
	return isBlank;
}

function insertDataIntoTableOfDb(db, tableName, columns, vals) {
	var command = "INSERT INTO :tableName VALUES (:values)";
	command = command.replace(/:tableName/g, sanitizeTableOrColumnNames(tableName));
	var valPlaceHolders = '';
	for (var i = 0; i < columns.length; i++) {
		valPlaceHolders += ':val' + i;
		if (i != (columns.length - 1))
			valPlaceHolders += ',';
	}
	command = command.replace(/:values/g, valPlaceHolders);
	var values = {};
	for (var i = 0; i < columns.length; i++) {
		values[':val' + i] = vals[i];
	}
	console.log(command);
	db.run(command, values);
}

function saveDBToDisk(db, path) {
	var data = db.export();
	var buffer = new Buffer(data);
	fs.writeFileSync(path, buffer);
}

function excelValueAtCellInSheet(sheet, addr) {
	var cell = sheet[addr];
	if (cell)
		return cell.v;
	else
		return null;
}

function excelAddressForCell(row, col) {
	return excelColumnNameForIndex(col) + row;
}

function excelColumnNameForIndex(index) {
	// just handle up to 26 for now
	return String.fromCharCode(64 + index);
}
