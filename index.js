var async = require("async");
var request = require("request");
var xml2js = require("xml2js");
var http = require("http");
var querystring = require("querystring");
var _ = require('lodash');
var GoogleAuth = require("google-auth-library");
var google = require('googleapis');


var GOOGLE_FEED_URL = "https://spreadsheets.google.com/feeds/";
var GOOGLE_AUTH_SCOPE = ['https://www.googleapis.com/auth/spreadsheets'];

var REQUIRE_AUTH_MESSAGE = 'You must authenticate to modify sheet data';

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

  var xml_parser = new xml2js.Parser({
    // options carried over from older version of xml2js
    // might want to update how the code works, but for now this is fine
    explicitArray: false,
    explicitRoot: false
  });

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
    if (!options.visibility){
      visibility = google_auth ? 'private' : 'public';
    }
    if (!options.projection){
      projection = google_auth ? 'full' : 'values';
    }
  }

  /*
  // This method is used internally to make all requests
  //## OLD: This is replaced internally with calls to the Google Sheets API
  this.makeFeedRequest = function( url_params, method, query_or_data, cb ){
    var url;
    var headers = {};
    if (!cb ) cb = function(){};
    if ( typeof(url_params) == 'string' ) {
      // used for edit / delete requests
      url = url_params;
    } else if ( Array.isArray( url_params )){
      //used for get and post requets
      url_params.push( visibility, projection );
      url = GOOGLE_FEED_URL + url_params.join("/");
    }

    async.series({
      auth: function(step) {
        if (auth_mode != 'jwt') return step();
        // check if jwt token is expired
        if (google_auth && google_auth.expires > +new Date()) return step();
        renewJwtAuth(step);
      },
      request: function(result, step) {
        if ( google_auth ) {
          if (google_auth.type === 'Bearer') {
            headers['Authorization'] = 'Bearer ' + google_auth.value;
          } else {
            headers['Authorization'] = "GoogleLogin auth=" + google_auth;
          }
        }

        headers['Gdata-Version'] = '3.0';

        if ( method == 'POST' || method == 'PUT' ) {
          headers['content-type'] = 'application/atom+xml';
        }

        if (method == 'PUT' || method == 'POST' && url.indexOf('/batch') != -1) {
          headers['If-Match'] = '*';
        }

        if ( method == 'GET' && query_or_data ) {
          var query = "?" + querystring.stringify( query_or_data );
          // replacements are needed for using structured queries on getRows
          query = query.replace(/%3E/g,'>');
          query = query.replace(/%3D/g,'=');
          query = query.replace(/%3C/g,'<');
          url += query;
        }

        request( {
          url: url,
          method: method,
          headers: headers,
          body: method == 'POST' || method == 'PUT' ? query_or_data : null
        }, function(err, response, body){
          if (err) {
            return cb( err );
          } else if( response.statusCode === 401 ) {
            return cb( new Error("Invalid authorization key."));
          } else if ( response.statusCode >= 400 ) {
            var message = _.isObject(body) ? JSON.stringify(body) : body.replace(/&quot;/g, '"');
            return cb( new Error("HTTP error "+response.statusCode+" ("+http.STATUS_CODES[response.statusCode])+") - "+message);
          } else if ( response.statusCode === 200 && response.headers['content-type'].indexOf('text/html') >= 0 ) {
            return cb( new Error("Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)"));
          }


          if ( body ){
            xml_parser.parseString(body, function(err, result){
              if ( err ) return cb( err );
              cb( null, result, body );
            });
          } else {
            if ( err ) cb( err );
            else cb( null, true );
          }
        })
      }
    });
  }
  */



  // public API methods
  //## Complete? (except for worksheets)
  this.getInfo = function( cb ){	
	sheets.spreadsheets.get({
	  auth: jwt_client,
	  spreadsheetId: ss_key,
	  fields:'',
	}, function(err, response) {
	if (err) {
		console.log('ERROR: ' + err);
		return cb(err);
	} else {
		//console.log('success');
		console.log( JSON.stringify(response, null, 4) );
		//console.log(response.sheets);

		var ss_data = {
			id: response.spreadsheetUrl,
			title: response.properties.title,
			//updated: data.updated,	// no longer used
			//author: data.author,		// no longer used
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

  
  // NOTE: worksheet IDs start at 1
  /*
  this.addWorksheet = function( opts, cb ) {
    // make opts optional
    if (typeof opts == 'function'){
      cb = opts;
      opts = {};
    }

    cb = cb || _.noop;

    if (!this.isAuthActive()) return cb(new Error(REQUIRE_AUTH_MESSAGE));

    var defaults = {
      title: 'Worksheet '+(+new Date()),  // need a unique title
      rowCount: 50,
      colCount: 20
    };

    var opts = _.extend({}, defaults, opts);

    // if column headers are set, make sure the sheet is big enough for them
    if (opts.headers && opts.headers.length > opts.colCount) {
      opts.colCount = opts.headers.length;
    }

    var data_xml = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006"><title>' +
        opts.title +
      '</title><gs:rowCount>' +
        opts.rowCount +
      '</gs:rowCount><gs:colCount>' +
        opts.colCount +
      '</gs:colCount></entry>';

    self.makeFeedRequest( ["worksheets", ss_key], 'POST', data_xml, function(err, data, xml) {
      if ( err ) return cb( err );

      var sheet = new SpreadsheetWorksheet( self, data );
      self.worksheets = self.worksheets || [];
      self.worksheets.push(sheet);
      sheet.setHeaderRow(opts.headers, function(err) {
        cb(err, sheet);
      })
    });
  }

  this.removeWorksheet = function ( sheet_id, cb ){
    if (!this.isAuthActive()) return cb(new Error(REQUIRE_AUTH_MESSAGE));
    if (sheet_id instanceof SpreadsheetWorksheet) return sheet_id.del(cb);
    self.makeFeedRequest( GOOGLE_FEED_URL + "worksheets/" + ss_key + "/private/full/" + sheet_id, 'DELETE', null, cb );
  }
  */

  //## Complete
  this.getRows = function( worksheet_id, opts, cb ){
    // The first row is used as titles/keys and is not included (## is now included)

    // opts is optional
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
			console.log('request:\n\n ' + JSON.stringify( request, null, 4 ));
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

  this.getColumn = function( worksheet_id, colNum, cb ){
    // The first row is used as titles/keys and is not included (## is now included)

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
  this.addRow = function( worksheet_id, data, cb ){
    var data_xml = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">' + "\n";
    Object.keys(data).forEach(function(key) {
      if (key != 'id' && key != 'title' && key != 'content' && key != '_links'){
        data_xml += '<gsx:'+ xmlSafeColumnName(key) + '>' + xmlSafeValue(data[key]) + '</gsx:'+ xmlSafeColumnName(key) + '>' + "\n"
      }
    });
    data_xml += '</entry>';
    self.makeFeedRequest( ["list", ss_key, worksheet_id], 'POST', data_xml, function(err, data, new_xml) {
      if (err) return cb(err);
      var entries_xml = new_xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g);
      var row = new SpreadsheetRow(self, data, entries_xml[0]);
      cb(null, row);
    });
  }
  */

  this.getCells = function (worksheet_id, opts, cb) {
    // opts is optional
    if (typeof( opts ) == 'function') {
      cb = opts;
      opts = {};
    }

    // Supported options are:
    // min-row, max-row, min-col, max-col, return-empty
    var query = _.assign({}, opts);


    self.makeFeedRequest(["cells", ss_key, worksheet_id], 'GET', query, function (err, data, xml) {
      if (err) return cb(err);
      if (data===true) {
        return cb(new Error('No response to getCells call'))
      }

      var cells = [];
      var entries = forceArray(data['entry']);
      var i = 0;
      entries.forEach(function( cell_data ){
        cells.push( new SpreadsheetCell( self, worksheet_id, cell_data ) );
      });

      cb( null, cells );
    });
  }
};

// Classes
var SpreadsheetWorksheet = function( spreadsheet, data ){
  var self = this;
  var links;

  //self.url = data.id;											//## Unused?
  //self.id = data.properties.index;							//## Unused?
  self.index = data.properties.index;							//## NEW
  self.title = data.properties.title;							//## OK
  self.rowCount = data.properties.gridProperties.rowCount;		//## OK
  self.colCount = data.properties.gridProperties.columnCount;	//## OK

  // _links unused
  /*
  self['_links'] = [];
  links = forceArray( data.link );
  links.forEach( function( link ){
    self['_links'][ link['$']['rel'] ] = link['$']['href'];
  });
  self['_links']['cells'] = self['_links']['http://schemas.google.com/spreadsheets/2006#cellsfeed'];
  self['_links']['bulkcells'] = self['_links']['cells']+'/batch';  
  
  
  function _setInfo(opts, cb) {
    cb = cb || _.noop;
    var xml = ''
      + '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006">'
      + '<title>'+(opts.title || self.title)+'</title>'
      + '<gs:rowCount>'+(opts.rowCount || self.rowCount)+'</gs:rowCount>'
      + '<gs:colCount>'+(opts.colCount || self.colCount)+'</gs:colCount>'
      + '</entry>';
    spreadsheet.makeFeedRequest(self['_links']['edit'], 'PUT', xml, function(err, response) {
      if (err) return cb(err);
      self.title = response.title;
      self.rowCount = parseInt(response['gs:rowCount']);
      self.colCount = parseInt(response['gs:colCount']);
      cb();
    });
  }

  this.resize = _setInfo;
  this.setTitle = function(title, cb) {
    _setInfo({title: title}, cb);
  }


  // just a convenience method to clear the whole sheet
  // resizes to 1 cell, clears the cell, and puts it back
  this.clear = function(cb) {
    var cols = self.colCount;
    var rows = self.colCount;
    self.resize({rowCount: 1, colCount: 1}, function(err) {
      if (err) return cb(err);
      self.getCells(function(err, cells) {
        cells[0].setValue(null, function(err) {
          if (err) return cb(err);
          self.resize({rowCount: rows, colCount: cols}, cb);
        });
      })
    });
  }
  */
  // stopped here
  
  

  this.getRows = function(opts, cb){
    spreadsheet.getRows(self.title, opts, cb);
  }
  
  this.getColumn = function(colNum, cb) {
	spreadsheet.getColumn(self.title, colNum, cb);  
  }
  
  this.getCells = function(opts, cb) {
    spreadsheet.getCells(self.title, opts, cb);
  }
  /*
  this.addRow = function(data, cb){
    spreadsheet.addRow(self.title, data, cb);
  }
  */
  this.bulkUpdateCells = function(cells, cb) {
    if ( !cb ) cb = function(){};

    var entries = cells.map(function (cell, i) {
      cell._needsSave = false;
      return "<entry>\n        <batch:id>" + cell.batchId + "</batch:id>\n        <batch:operation type=\"update\"/>\n        <id>" + self['_links']['cells']+'/'+cell.batchId + "</id>\n        <link rel=\"edit\" type=\"application/atom+xml\"\n          href=\"" + cell._links.edit + "\"/>\n        <gs:cell row=\"" + cell.row + "\" col=\"" + cell.col + "\" inputValue=\"" + cell.valueForSave + "\"/>\n      </entry>";
    });
    var data_xml = "<feed xmlns=\"http://www.w3.org/2005/Atom\"\n      xmlns:batch=\"http://schemas.google.com/gdata/batch\"\n      xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">\n      <id>" + self['_links']['cells'] + "</id>\n      " + entries.join("\n") + "\n    </feed>";

    spreadsheet.makeFeedRequest(self['_links']['bulkcells'], 'POST', data_xml, function(err, data) {
      if (err) return cb(err);

      // update all the cells
      var cells_by_batch_id = _.indexBy(cells, 'batchId');
      if (data.entry && data.entry.length) data.entry.forEach(function(cell_data) {
        cells_by_batch_id[cell_data['batch:id']].updateValuesFromResponseData(cell_data);
      });
      cb();
    });
  }
  
  /*
  this.del = function(cb){
    spreadsheet.makeFeedRequest(self['_links']['edit'], 'DELETE', null, cb);
  }
  */

  //## Not required?
  /*
  this.setHeaderRow = function(values, cb) {
    if ( !cb ) cb = function(){};
    if (!values) return cb();
    if (values.length > self.colCount){
      return cb(new Error('Sheet is not large enough to fit '+values.length+' columns. Resize the sheet first.'));
    }
    self.getCells({
      'min-row': 1,
      'max-row': 1,
      'min-col': 1,
      'max-col': self.colCount,
      'return-empty': true
    }, function(err, cells) {
      if (err) return cb(err);
      _.each(cells, function(cell) {
        cell.value = values[cell.col-1] ? values[cell.col-1] : '';
      });
      self.bulkUpdateCells(cells, cb);
    });
  }
  */
}


// Storage class for a spreadsheet row
var SpreadsheetRow = function( auth, ss_key, worksheetId, rowIdx, data, header) {
  var self = this;
  this._auth = auth;
  this._ss_key = ss_key;
  this._worksheetId = worksheetId;
  this._rowIdx = rowIdx;  
  this._header = header;
  
  // Create map
  for (var i=0; i<header.length; i++) {
	var sanitizedLabel = sanitizeHeaderStr(header[i]);
	this[sanitizedLabel] = data[i];
  }
  

  // Save this row to Google Sheets
  self.save = function( cb ) {    

	// Convert from map to flat values array
	var values = [];
	for (var i=0; i<this._header.length; i++) {
		var sanitizedLabel = sanitizeHeaderStr(this._header[i]);
		values.push( this[sanitizedLabel] );
	}		
	  
	// Create ValueRange GoogleSheets API v4 structure for this row
	//var A1Range = '' + worksheet_id + '!' + XYtoA1(0, this.rowIdx) + ':' + XYtoA1(this.values.length-1, this.rowIdx);		// e.g. "Sheet1!A1:B2"
	var A1Range = '' + this._worksheetId + '!' + (this._rowIdx+1) + ':' + (this._rowIdx+1);		// e.g. "Sheet1!A1:B2"
	var valueRange = {
		range: A1Range,
		majorDimension: 'ROWS',
		values: [values],
	};
	
	var request = {
		auth: this._auth,
		spreadsheetId: this._ss_key,
		range: A1Range,
		valueInputOption: 'RAW',
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

// Storage class for a spreadsheet row
var SpreadsheetCol = function( auth, ss_key, worksheetId, colIdx, data) {
  var self = this;
  this._auth = auth;
  this._ss_key = ss_key;
  this._worksheetId = worksheetId;
  this._colIdx = colIdx;  
  this._originalHeaderLabel = '';
  if (data[0].length > 0) this._originalHeaderLabel = data[0][0];
  this._header = sanitizeHeaderStr(this._originalHeaderLabel);
  this.values = data[0].slice(1);  // trim off header label

  // Save this row to Google Sheets
  self.save = function( cb ) {    
	  
	// Create ValueRange GoogleSheets API v4 structure for this row	
	var A1Range = '' + this._worksheetId + '!' + IntToA1(colIdx) + ':' + IntToA1(colIdx);	// e.g. "Sheet1!A1:B2"		
	var valueRange = {
		range: A1Range,
		majorDimension: 'COLUMNS',
		values: [ [this._originalHeaderLabel].concat(this.values) ],		// Add back on header label
	};
	
	var request = {
		auth: this._auth,
		spreadsheetId: this._ss_key,
		range: A1Range,
		valueInputOption: 'RAW',
		resource: valueRange
	}
	
	console.log("\n\n* Save Request: \n" + JSON.stringify( request, null, 4 ) + "\n\n");
	
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





var SpreadsheetCell = function( spreadsheet, worksheet_id, data ){
  var self = this;

  function init() {
    var links;
    self.id = data['id'];
    self.row = parseInt(data['gs:cell']['$']['row']);
    self.col = parseInt(data['gs:cell']['$']['col']);
    self.batchId = 'R'+self.row+'C'+self.col;

    self['_links'] = [];
    links = forceArray( data.link );
    links.forEach( function( link ){
      self['_links'][ link['$']['rel'] ] = link['$']['href'];
    });

    self.updateValuesFromResponseData(data);
  }

  self.updateValuesFromResponseData = function(_data) {
    // formula value
    var input_val = _data['gs:cell']['$']['inputValue'];
    // inputValue can be undefined so substr throws an error
    // still unsure how this situation happens
    if (input_val && input_val.substr(0,1) === '='){
      self._formula = input_val;
    } else {
      self._formula = undefined;
    }

    // numeric values
    if (_data['gs:cell']['$']['numericValue'] !== undefined) {
      self._numericValue = parseFloat(_data['gs:cell']['$']['numericValue']);
    } else {
      self._numericValue = undefined;
    }

    // the main "value" - its always a string
    self._value = _data['gs:cell']['_'] || '';
  }

  self.setValue = function(new_value, cb) {
    self.value = new_value;
    self.save(cb);
  };

  self._clearValue = function() {
    self._formula = undefined;
    self._numericValue = undefined;
    self._value = '';
  }

  self.__defineGetter__('value', function(){
    return self._value;
  });
  self.__defineSetter__('value', function(val){
    if (!val) return self._clearValue();

    var numeric_val = parseFloat(val);
    if (!isNaN(numeric_val)){
      self._numericValue = numeric_val;
      self._value = val.toString();
    } else {
      self._numericValue = undefined;
      self._value = val;
    }

    if (typeof val == 'string' && val.substr(0,1) === '=') {
      // use the getter to clear the value
      self.formula = val;
    } else {
      self._formula = undefined;
    }
  });

  self.__defineGetter__('formula', function() {
    return self._formula;
  });
  self.__defineSetter__('formula', function(val){
    if (!val) return self._clearValue();

    if (val.substr(0,1) !== '=') {
      throw new Error('Formulas must start with "="');
    }
    self._numericValue = undefined;
    self._value = '*SAVE TO GET NEW VALUE*';
    self._formula = val;
  });

  self.__defineGetter__('numericValue', function() {
    return self._numericValue;
  });
  self.__defineSetter__('numericValue', function(val) {
    if (val === undefined || val === null) return self._clearValue();

    if (isNaN(parseFloat(val)) || !isFinite(val)) {
      throw new Error('Invalid numeric value assignment');
    }

    self._value = val.toString();
    self._numericValue = parseFloat(val);
    self._formula = undefined;
  });

  self.__defineGetter__('valueForSave', function() {
    return xmlSafeValue(self._formula || self._value);
  });

  self.save = function(cb) {
    if ( !cb ) cb = function(){};
    self._needsSave = false;

    var edit_id = 'https://spreadsheets.google.com/feeds/cells/key/worksheetId/private/full/R'+self.row+'C'+self.col;
    var data_xml =
      '<entry><id>'+self.id+'</id>'+
      '<link rel="edit" type="application/atom+xml" href="'+self.id+'"/>'+
      '<gs:cell row="'+self.row+'" col="'+self.col+'" inputValue="'+self.valueForSave+'"/></entry>'

    data_xml = data_xml.replace('<entry>', "<entry xmlns='http://www.w3.org/2005/Atom' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>");

    spreadsheet.makeFeedRequest( self['_links']['edit'], 'PUT', data_xml, function(err, response) {
      if (err) return cb(err);
      self.updateValuesFromResponseData(response);
      cb();
    });
  }

  self.del = function(cb) {
    self.setValue('', cb);
  }

  init();
  return self;
}

module.exports = GoogleSpreadsheet;

//utils
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
	os += String.fromCharCode(65+num);		// Up to 25/Z
	return os;
}

/*
//## Remove?
var xmlSafeValue = function(val){
  if ( val == null ) return '';
  return String(val).replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/\n/g,'&#10;')
      .replace(/\r/g,'&#13;');
}
//## Remove?
var xmlSafeColumnName = function(val){
  if (!val) return '';
  return String(val).replace(/[\s_]+/g, '')
      .toLowerCase();
}
*/