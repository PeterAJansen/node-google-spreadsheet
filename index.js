//var async = require("async");
//var querystring = require("querystring");
var _ = require('lodash');
var GoogleAuth = require("google-auth-library");
var google = require('googleapis');

var GOOGLE_AUTH_SCOPE = ['https://www.googleapis.com/auth/spreadsheets'];

// The main class that represents a single sheet
// this is the main module.exports
var GoogleSpreadsheet = function( ss_key, auth_id, options ){
	var self = this;
	var google_auth = null;
	var visibility = 'public';
	var projection = 'values';

	var auth_mode = 'anonymous';

	var auth_client = new GoogleAuth();
	var jwt_client;

	var sheets = google.sheets('v4');
  
	options = options || {};

	// Check for google spreadsheet key
	if ( !ss_key ) {
		throw new Error("Spreadsheet key not provided.");
	}

	// auth_id may be null
	setAuthAndDependencies(auth_id);

	// Authentication Methods

	this.setAuthToken = function( auth_id ) {
		if (auth_mode == 'anonymous') auth_mode = 'token';
		setAuthAndDependencies(auth_id);
	}

	// deprecated username/password login method
	// leaving it here to help notify users why it doesn't work
	this.setAuth = function( username, password, cb ){
		return cb(new Error('Google has officially deprecated ClientLogin. Please upgrade this module and see the readme for more instrucations'))
	}

	this.useServiceAccountAuth = function( creds, cb ){
		if (typeof creds == 'string') {
			try {
				creds = require(creds);
			} catch (err) {
				return cb(err);
			}
		}
		jwt_client = new auth_client.JWT(creds.client_email, null, creds.private_key, GOOGLE_AUTH_SCOPE, null);
		renewJwtAuth(cb);		
	}

	function renewJwtAuth(cb) {
		auth_mode = 'jwt';
		jwt_client.authorize(function (err, token) {
			if (err) return cb(err);
			self.setAuthToken({
				type: token.token_type,
				value: token.access_token,
				expires: token.expiry_date
			});
			cb()
		});
	}

	this.isAuthActive = function() {
		return !!google_auth;
	}


	function setAuthAndDependencies( auth ) {
		google_auth = auth;
		if (!options.visibility) {
			visibility = google_auth ? 'private' : 'public';
		}
		if (!options.projection) {
			projection = google_auth ? 'full' : 'values';
		}
	}



	/*
	 * Public API Methods
	 */
	
	// Get Google Sheet information
	this.getInfo = function( cb ) {
		sheets.spreadsheets.get({
			auth: jwt_client,
			spreadsheetId: ss_key,
			fields: '',
		}, function(err, response) {
			if (err) {
				console.log('ERROR: ' + err);
				return cb(err);
			} else {
				//console.log('success');
				//console.log( JSON.stringify(response, null, 4) );
				//console.log(response.sheets);

				var ss_data = {
					id: response.spreadsheetUrl,
					title: response.properties.title,
					worksheets: []
				}
	
				var worksheets = forceArray(response.sheets);		// not sure what this does -- seems to just initialize the array to []
				worksheets.forEach( function( ws_data ) {
					ss_data.worksheets.push( new SpreadsheetWorksheet( self, ws_data ) );
				});
				self.info = ss_data;
				self.worksheets = ss_data.worksheets;
				cb( null, ss_data );
			}
		});		
	}

    
	// Retrieve all rows from the worksheet
	this.getRows = function( worksheet_id, opts, cb ) {		
		// opts is optional (Note: Currently unused)
		if ( typeof( opts ) == 'function' ){
			cb = opts;
			opts = {};
		}

		// Form request
		var A1Range = '' + worksheet_id;		// e.g. "Sheet1!A1:B2"
		var request = {
			auth: jwt_client,
			spreadsheetId: ss_key,
			range: A1Range,		
		};
	
		// Request data from Google Sheets
		sheets.spreadsheets.values.get(request, function(err, response) {
			if (err) {			
				console.log('getRows: ERROR: ' + err);
				//console.log('request:\n\n ' + JSON.stringify( request, null, 4 ));
				return cb(err);
			} else {		
				//console.log( JSON.stringify(response, null, 4) );

				var range = response.range;						// e.g. "range": "'Class Data'!A1:Z988",
				var majorDimension = response.majorDimension;	// e.g. "majorDimension": "ROWS",
		
				// Read rows
				var rows = [];
				var entries = forceArray( response.values );
			
				// Find header
				var header = [];
				if (entries.length > 0) {
					header = entries[0];
				}				
			
				// Populate spreadsheet rows
				for (var rowNum=0; rowNum<entries.length; rowNum++) {
					rows.push( new SpreadsheetRow( jwt_client, ss_key, worksheet_id, rowNum, entries[rowNum], header ) ); //## TODO
				}
			
				// Callback
				cb(null, rows);	
			}
		});        
	}

	
	// Retrieve a single column from the worksheet
	this.getColumn = function( worksheet_id, colNum, cb ) {		
		// The first row is not included in the data, but is kept in a 'header' field.
    
		// Form request
		var A1Range = '' + worksheet_id + '!' + IntToA1(colNum) + ':' + IntToA1(colNum);	// e.g. "Sheet1!A1:B2"
		var request = {
			auth: jwt_client,
			spreadsheetId: ss_key,
			range: A1Range,	
			majorDimension: 'COLUMNS',
		};
	
		// Request data from Google Sheets
		sheets.spreadsheets.values.get(request, function(err, response) {
			if (err) {			
				console.log('getColumn: ERROR: ' + err);
				console.log('request:\n\n ' + JSON.stringify( request, null, 4 ));
				return cb(err);
			} else {		
				//console.log( JSON.stringify(response, null, 4) );

				var range = response.range;						// e.g. "range": "'Class Data'!A1:Z988",
				var majorDimension = response.majorDimension;	// e.g. "majorDimension": "ROWS",
		
				// Read column
				//var cols = [];
				var entries = forceArray( response.values );
						
				// Populate spreadsheet rows			
				var col = new SpreadsheetCol( jwt_client, ss_key, worksheet_id, colNum, entries );			
			
				// Callback
				cb(null, col);	
			}
		});        
	}
  
  
	/*
	// Unimplemented
	this.addRow = function( worksheet_id, data, cb ) {
		...
	}
	*/
	
};




// Classes
var SpreadsheetWorksheet = function( spreadsheet, data ) {
	var self = this;

	self.index = data.properties.index;
	self.title = data.properties.title;
	self.rowCount = data.properties.gridProperties.rowCount;
	self.colCount = data.properties.gridProperties.columnCount;
    

	this.getRows = function(opts, cb){
		spreadsheet.getRows(self.title, opts, cb);
	}
  
	this.getColumn = function(colNum, cb) {
		spreadsheet.getColumn(self.title, colNum, cb);  
	}
    
}


// Storage class for a spreadsheet row
var SpreadsheetRow = function( auth, ss_key, worksheetId, rowIdx, data, header) {
	var self = this;
	this._auth = auth;
	this._ss_key = ss_key;
	this._worksheetId = worksheetId;
	this._rowIdx = rowIdx;  
	this._originalHeaderLabels = header;
	this._header = [];
	this.values = data;
		
	// Step 1: Automatically grow values array to be the same size as the header
	while (this.values.length < header.length) {
		this.values.push("");
	}
	
	// Step 1: Sanitize header labels
	for (var i=0; i<header.length; i++) {
		this._header.push( sanitizeHeaderStr(header[i]) );
	}			
	
	// Setter based on giving the key (header column) name
	self.setValue = function (key, value) {
		for (var i=0; i<header.length; i++) {
			if (this._header[i] == key) {
				this.values[i] = value;
				return;
			}
		}
	}
	
	// Getter baed on giving the key (header column) name
	self.getValue = function(key) {
		for (var i=0; i<this._header.length; i++) {
			if (this._header[i] == key) {				
				return this.values[i];
			}
		}
		return null;
	}
  

	// Save this row to Google Sheets
	self.save = function( cb ) {    		
	  
		// Create ValueRange GoogleSheets API v4 structure for this row	
		var A1Range = '' + this._worksheetId + '!' + (this._rowIdx+1) + ':' + (this._rowIdx+1);		// e.g. "Sheet1!A1:B2"
		var valueRange = {
			range: A1Range,
			majorDimension: 'ROWS',
			values: [this.values],
		};
	
		var request = {
			auth: this._auth,
			spreadsheetId: this._ss_key,
			range: A1Range,
			//valueInputOption: 'RAW',
			valueInputOption: 'USER_ENTERED',
			resource: valueRange
		}
	
		//console.log("Save Request: \n" + JSON.stringify( request, null, 4 ) );
	
		var sheets = google.sheets('v4');
		sheets.spreadsheets.values.update(request, function(err, response) {
			if (err) {			
				console.log('SpeadsheetRow.save(): ERROR: ' + err);			
				return cb(err);
			} else {		
				//console.log('SpeadsheetRow.save(): SUCCESS: ');
				//console.log( JSON.stringify(response, null, 4) );						
				return cb(null);	
			}
		});        
	}  
	
}


// Storage class for a spreadsheet column
var SpreadsheetCol = function( auth, ss_key, worksheetId, colIdx, data) {
	var self = this;
	this._auth = auth;
	this._ss_key = ss_key;
	this._worksheetId = worksheetId;
	this._colIdx = colIdx;  
	this._originalHeaderLabel = '';
	if (data[0].length > 0) this._originalHeaderLabel = data[0][0];
	this._header = sanitizeHeaderStr(this._originalHeaderLabel);
	this.values = data[0].slice(1);  // Main column data. (Note: slice trims off header label)

	// Save this column to Google Sheets
	self.save = function( cb ) {    
	  
		// Create ValueRange GoogleSheets API v4 structure for this row	
		var A1Range = '' + this._worksheetId + '!' + IntToA1(colIdx) + ':' + IntToA1(colIdx);	// e.g. "Sheet1!A1:B2"		
		var valueRange = {
			range: A1Range,
			majorDimension: 'COLUMNS',
			values: [ [this._originalHeaderLabel].concat(this.values) ],	// Add header label back on
		};
	
		var request = {
			auth: this._auth,
			spreadsheetId: this._ss_key,
			range: A1Range,
			valueInputOption: 'RAW',
			resource: valueRange
		}
	
		//console.log("\n\n* Save Request: \n" + JSON.stringify( request, null, 4 ) + "\n\n");
	
		var sheets = google.sheets('v4');
		sheets.spreadsheets.values.update(request, function(err, response) {
			if (err) {			
				console.log('SpeadsheetCol.save(): ERROR: ' + err);			
				return cb(err);
			} else {		
				//console.log('SpeadsheetCol.save(): SUCCESS: ');
				//console.log( JSON.stringify(response, null, 4) );						
				return cb(null);	
			}
		});        
	}  
	
}


module.exports = GoogleSpreadsheet;

/*
 * Utilties
 */ 
var forceArray = function(val) {
  if ( Array.isArray( val ) ) return val;
  if ( !val ) return [];
  return [ val ];
}

// Remove spaces and non-alphanumeric characters from a string
var sanitizeHeaderStr = function(strIn) {
	var strOut = strIn.replace(/[^0-9a-zA-Z]/gi, '').toLowerCase();
	return strOut;
}


// Convert from X/Y notation to A1 notation
var XYtoA1 = function (x, y) {
	var os = '';	
	
	// X	
	os += IntToA1(x);
	
	// Y
	os += (y+1);
	
	return os;
}

var IntToA1 = function(num) {
	var os = '';
	os += String.fromCharCode(65+num);		//TODO: This only works up to 25/Z
	return os;
}
