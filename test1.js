var GoogleSpreadsheet = require('./index.js');
var async = require('async');
var _ = require("lodash");

// spreadsheet key is the long id in the sheets URL
var doc = new GoogleSpreadsheet('1pU6-_mZ-FcpgKvtl-ljOVTE5pvXj4squtrPG6Dj0VtE');
var creds = require('./myproject_service_account.json');


    async.waterfall([
        function setAuth(next) {
            doc.useServiceAccountAuth(creds, next);
        }, function getTsData(next) {
            doc.getInfo(function(err, info) {
                console.log("Loaded google sheet: " + info.title);
                
                next(null, info.worksheets);
            });
        }, function(sheets, next) {
			
            for (var i=0; i<sheets.length; i++) {				
				var sheet = sheets[i];
				
				// Rows
				
                sheet.getRows({offset: 1}, function(err, rows) {
                    if (err) console.log(err);
                    console.log(sheet.title + " (" + sheet.rowCount + "x" + sheet.colCount + ") --> Read " + rows.length + " rows");                    
                    
                    // Create new (blank) single table
                    var tableName = sheet.title
                    var headerRow = Object.keys(rows[0]);
                                        
                    
                    // Store each row in the tablestore                    
                    
					for (var j=0; j<rows.length; j++) {
						var row = rows[j];
                        //oneTable.addRow( row );
						console.log("Row:\n " + JSON.stringify(row, null, 4) );
						
						// Save test
						if ((row._worksheetId == "Sheet3") && (row._rowIdx == 1)) {
							console.log("Save Test");
							//row["test111"] += "1";
							// Write by index
							//row.values[0] += "1";		
							// Write by header key
							row.setValue("test111", row.getValue("test111") + "Z");
							row.save( function (err) {
								if (err) {
									console.log(err)									
								}
							});
						}
                    }
																		
                    next();
                    
                    
                });  
				
				/*
				// Column
				console.log("COLUMN:\n");
				sheet.getColumn(0, function(err, col) {
					if (err) {
						console.log(err);
					} else {
						console.log("COLUMN SUCCESS:\n");
						console.log(col);
						
						col.values[0] += "A";
						col.values[1] += "B";
						col.save( function (err) {
							if (err) {
								console.log(err)									
							}
						});
						
					}
				});
				*/						
				
            }					
			
        }
        
    ], function(err, result) {
        if (err) console.log(err);
        
    });

