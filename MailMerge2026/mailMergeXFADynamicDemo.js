//use windows environment variables %variable%
//temp files created in %LOCALAPPDATA%\Temp\acrobat_sbx
//save this to %APPDATA%\Adobe\Acrobat\Privileged\DC\Javascripts
//this creates the menu button in the add-ons tool menu
if (app["viewerVersion"] < 10) {
    app["addMenuItem"]({
      cName: "MailMergeXFA",
      cUser: "Mail Merge XFA",
      cParent: "Tools",
      cExec: "mailMergeXFAFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  } else {
    app["addToolButton"]({
      cName: "MailMergeXFA",
      cLabel: "Mail Merge XFA",
      cTooltext: "Mail Merge XFA",
      cExec: "mailMergeXFAFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  };

  //setup mailMergeXFADialog
  var mailMergeXFADialog = {
    radioValue: null,
    initialize: function (dialog) {
      dialog.load({
        "rd01": true,
      });
    },
    commit: function (dialog) {
      var oRslt = dialog.store();
      this.radioValue = this.getRadioValue(oRslt);
    },
    getRadioValue: function (results) {
      for ( var i=1; i<=3; i++) {
        if ( results["rd0"+i] ) {
        switch (i) {
        case 1:
        var cDelim = ",";
        break;
        case 2:
        var cDelim = "\t";
        break;
        }
        }
      };
      return cDelim;
    },
    description: {
      name: "Mail Merge XFA",
      elements: [{
          type: "view",
          elements: [{
              item_id: "str1",
              name: "Mail Merge XFA",
              type: "static_text",
              font: "dialog",
              bold: true,
              alignment: "align_center",
              char_width: 30,
              height: 20,
            },
            {
              type: "gap",
            },
            {
              type: "view",
              align_children: "align_row",
              elements: [{
                  item_id: "str2",
                  name: "Delimiter:",
                  type: "static_text",
                },
                {
                    type: "radio",
                    item_id: "rd01",
                    group_id: "group1",
                    name: "Comma"
                  },
                  {
                    type: "radio",
                    item_id: "rd02",
                    group_id: "group1",
                    name: "Tab"
                  },
              ],
            },
          ],
        },
        {
          type: "ok_cancel",
          alignment: "align_right",
        },
      ]
    },
  };

  // Function to get all XFA fields and create a mapping
  function getXFAFieldMap(doc) {
    console.println("\n========== STARTING FIELD MAP CREATION ==========");
    var fieldMap = {};
    
    try {
      // Access xfa through the document object
      var xfaObj = doc.xfa;
      console.println("XFA object accessed successfully");
      console.println("Number of pages: " + xfaObj.host.numPages);
      
      for (var i = 0; i < xfaObj.host.numPages; i++) {
        console.println("\n--- Scanning Page " + i + " ---");
        var oFields = xfaObj.layout.pageContent(i, "field");
        var nodesLength = oFields.length;
        console.println("Found " + nodesLength + " fields on page " + i);
        
        for (var j = 0; j < nodesLength; j++) {
          var oItem = oFields.item(j);
          var fieldName = oItem.name;
          var somPath = oItem.somExpression;
          
          console.println("  Field #" + j + ":");
          console.println("    Name: " + fieldName);
          console.println("    SOM Path: " + somPath);
          console.println("    Current Value: " + oItem.rawValue);
          
          // Store the field with its SOM path
          // If multiple fields have the same name, store them in an array
          if (!fieldMap[fieldName]) {
            fieldMap[fieldName] = [];
          }
          
          fieldMap[fieldName].push({
            somPath: somPath,
            field: oItem
          });
        }
      }
      
      console.println("\n========== FIELD MAP SUMMARY ==========");
      console.println("Total unique field names: " + Object.keys(fieldMap).length);
      for (var fname in fieldMap) {
        console.println("  '" + fname + "' has " + fieldMap[fname].length + " instance(s)");
      }
      console.println("========================================\n");
      
    } catch (e) {
      console.println("ERROR in getXFAFieldMap: " + e);
      console.println("Stack: " + e.stack);
    }
    
    return fieldMap;
  }

  function mailMergeXFAFiles(doc) {
    console.println("\n########## MAIL MERGE XFA STARTED ##########");
    console.println("Document name: " + doc.documentFileName);
    console.println("Document path: " + doc.path);
    
    if (doc["xfa"] == null) {
      console.println("ERROR: Document is not an XFA form!");
      xfaAlert("Error! This script is designed for XFA forms only.");
      return;
    };
    
    console.println("XFA form detected successfully");
    
    // Get delimiter from user
    console.println("\nShowing delimiter dialog...");
    if ("ok" == xfaDialog(mailMergeXFADialog)) {
      var delimiterValue = mailMergeXFADialog.radioValue;
      console.println("User selected delimiter: " + (delimiterValue == "," ? "Comma" : "Tab"));
    } else {
      console.println("User cancelled delimiter dialog");
      return;
    }
    
    if (delimiterValue == null || delimiterValue == "") {
      console.println("ERROR: No delimiter selected");
      return;
    };
    
    console.println("\nReading CSV file...");
    var fileDataArray = readXFACSV(delimiterValue, null);
    
    if (fileDataArray == null || fileDataArray["length"] == 0) {
      console.println("ERROR: No data in CSV file");
      xfaAlert("No items found in the input file.");
      return;
    };
    
    console.println("CSV loaded successfully. Total rows: " + fileDataArray.length);
    console.println("Header row: " + fileDataArray[0].join(", "));
    
    // Build XFA field map
    console.println("\n========== Building XFA field map ==========");
    var xfaFieldMap = getXFAFieldMap(doc);
    console.println("Field map complete. Found " + Object.keys(xfaFieldMap).length + " unique field names.");
    
    // Get CSV headers
    var csvHeaders = fileDataArray[0];
    console.println("\n========== CSV HEADERS ==========");
    for (var h = 0; h < csvHeaders.length; h++) {
      console.println("  Column " + h + ": '" + csvHeaders[h] + "'");
    }
    console.println("=================================\n");
    
    // Find which column to use for filename (look for "To" column)
    var filenameColumnIndex = -1;
    for (var i = 0; i < csvHeaders["length"]; i++) {
      if (csvHeaders[i] == "To") {
        filenameColumnIndex = i;
        console.println("Found 'To' column at index: " + i);
        break;
      }
    }
    
    if (filenameColumnIndex == -1) {
      console.println("ERROR: 'To' column not found in CSV");
      xfaAlert("Error! Could not find 'To' column in CSV file for naming files.");
      return;
    }
    
    var filesCreated = 0;
    var hasErrors = false;
    var docPath = doc["path"]["replace"](doc["documentFileName"], "");
    console.println("Output directory: " + docPath);
    
    // Loop through each row in the CSV (skip header row)
    for (var rowIndex = 1; rowIndex < fileDataArray["length"]; rowIndex++) {
      
      console.println("\n\n########## PROCESSING ROW " + rowIndex + " ##########");
      
      var filenameValue = "";
      var rowData = fileDataArray[rowIndex];
      
      console.println("Row data: " + rowData.join(" | "));
      
      // Loop through each column/header
      for (var colIndex = 0; colIndex < csvHeaders["length"]; colIndex++) {
        
        if (colIndex >= rowData["length"]) {
          console.println("  Column " + colIndex + " (" + csvHeaders[colIndex] + "): SKIPPED - no data in row");
          continue;
        }
        
        var headerName = csvHeaders[colIndex];
        var cellValue = rowData[colIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
        
        console.println("\n  --- Column " + colIndex + ": '" + headerName + "' ---");
        console.println("      Raw value: '" + rowData[colIndex] + "'");
        console.println("      Clean value: '" + cellValue + "'");
        
        // Save filename value
        if (colIndex == filenameColumnIndex) {
          filenameValue = cellValue;
          console.println("      ** This is the filename column **");
        }
        
        // Check if this header matches any XFA field name
        if (xfaFieldMap[headerName]) {
          var matchingFields = xfaFieldMap[headerName];
          console.println("      MATCH FOUND! " + matchingFields.length + " field(s) with this name");
          
          // If there's only one field with this name, use it
          if (matchingFields.length == 1) {
            try {
              console.println("      Setting field: " + matchingFields[0].somPath);
              console.println("      Old value: '" + matchingFields[0].field.rawValue + "'");
              matchingFields[0].field.rawValue = cellValue;
              console.println("      New value: '" + matchingFields[0].field.rawValue + "'");
              console.println("      SUCCESS!");
            } catch (e) {
              console.println("      ERROR setting field: " + e);
              console.println("      Stack: " + e.stack);
              hasErrors = true;
            }
          } else {
            // Multiple fields with same name - use the first one (index 0)
            try {
              console.println("      Multiple fields found, using first: " + matchingFields[0].somPath);
              console.println("      Old value: '" + matchingFields[0].field.rawValue + "'");
              matchingFields[0].field.rawValue = cellValue;
              console.println("      New value: '" + matchingFields[0].field.rawValue + "'");
              console.println("      SUCCESS!");
            } catch (e) {
              console.println("      ERROR setting field: " + e);
              console.println("      Stack: " + e.stack);
              hasErrors = true;
            }
          }
        } else {
          console.println("      WARNING: No XFA field found for this CSV header");
        }
      }
      
      // Create safe filename from "To" value
      var fileName = filenameValue["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
      if (fileName == "") {
        fileName = "output_" + rowIndex;
      }
      
      // Add .pdf extension if not present
      if (!fileName.match(/\.pdf$/i)) {
        fileName = fileName + ".pdf";
      }
      
      var savePath = docPath + fileName;
      
      console.println("\n  Saving file...");
      console.println("  Filename: " + fileName);
      console.println("  Full path: " + savePath);
      
      // Save the document
      try {
        xfaTrustedSaveAs(doc, savePath);
        filesCreated++;
        console.println("  SUCCESS! File saved.");
      } catch (e) {
        console.println("  ERROR saving file: " + e);
        console.println("  Stack: " + e.stack);
        hasErrors = true;
      }
      
      console.println("########## END ROW " + rowIndex + " ##########\n");
    }
    
    console.println("\n########## MAIL MERGE COMPLETE ##########");
    console.println("Files created: " + filesCreated);
    console.println("Errors occurred: " + hasErrors);
    console.println("##########################################\n");
    
    if (hasErrors) {
      xfaAlert("Done. " + filesCreated + " merged files were created, but there were some errors. See the Console for more details.", 1);
      console["show"]();
    } else {
      xfaAlert({cMsg:"Done. " + filesCreated + " merged files were created.", nIcon:4});
    }
  }

  function readXFACSV(delimiter, unused) {
    console.println("\n--- readXFACSV function started ---");
    console.println("Delimiter: " + delimiter);
    
    if (delimiter == null) {
      delimiter = ";";
      console.println("Delimiter was null, using default: ;");
    };
    
    var csvContent = "";
    console.println("Creating temporary document...");
    var tempDoc = xfaTrustedNewDoc();
    console.println("Temporary document created");
    
    var csvFileName = "file.csv";
    console.println("Importing data object: " + csvFileName);
    xfaTrustedImportDataObject(tempDoc, csvFileName);
    
    if (tempDoc["dataObjects"] == null || tempDoc["dataObjects"]["length"] != 1) {
      console.println("ERROR: No data objects found or wrong number of data objects");
      console.println("Data objects: " + tempDoc["dataObjects"]);
      if (tempDoc["dataObjects"]) {
        console.println("Number of data objects: " + tempDoc["dataObjects"]["length"]);
      }
      return;
    };
    
    console.println("Data object imported successfully");
    console.println("Getting data object contents...");
    var dataStream = tempDoc["getDataObjectContents"](csvFileName);
    var csvContent = util["stringFromStream"]({
      oStream: dataStream,
      oCharSet: "utf-8"
    });
    
    console.println("CSV content length: " + csvContent.length + " characters");
    console.println("First 200 characters: " + csvContent.substring(0, 200));
    
    csvContent = csvContent["replace"](/\r\n/g, "\n");
    csvContent = csvContent["replace"](/\r/g, "\n");
    tempDoc["closeDoc"](true);
    console.println("Temporary document closed");
    
    var resultArray = new Array();
    var fileDataArray = csvContent["split"]("\n");
    console.println("Split into " + fileDataArray.length + " lines");
    
    for (var i = 0; i < fileDataArray["length"]; i++) {
      var cleanLine = fileDataArray[i]["replace"](/[\r\n]/g, "")["replace"]("﻿", "");
      if (cleanLine != "") {
        if (delimiter == ",") {
          resultArray["push"](splitXFACSV(cleanLine, delimiter));
        } else {
          resultArray["push"](cleanLine["split"](delimiter));
        }
      }
    };
    
    console.println("Parsed into " + resultArray.length + " rows");
    console.println("--- readXFACSV function complete ---\n");
    
    return resultArray;
  }

  function splitXFACSV(str, delim) {
    var arr = str["split"](delim = delim || ",");
    var arrLen = arr["length"] - 1;
    for (var temp; arrLen >= 0; arrLen--) {
      if (arr[arrLen]["replace"](/"\s+$/, "\"")["charAt"](arr[arrLen]["length"] - 1) == "\"") {
        if ((temp = arr[arrLen]["replace"](/^\s+"/, "\""))["length"] > 1 && temp["charAt"](0) == "\"") {
          arr[arrLen] = arr[arrLen]["replace"](/^\s*"|"\s*$/g, "")["replace"](/""/g, "\"");
        } else {
          if (arrLen) {
            arr["splice"](arrLen - 1, 2, [arr[arrLen - 1], arr[arrLen]]["join"](delim));
          } else {
            arr = arr["shift"]()["split"](delim)["concat"](arr);
          }
        }
      } else {
        arr[arrLen]["replace"](/""/g, "\"");
      }
    };
    return arr;
  }

  xfaImportDataObject = app["trustPropagatorFunction"](function (doc, objName) {
    app["beginPriv"]();
    doc["importDataObject"]({
      cName: objName
    });
    app["endPriv"]();
  });

  xfaTrustedImportDataObject = app["trustedFunction"](function (doc, objName) {
    app["beginPriv"]();
    xfaImportDataObject(doc, objName);
    app["endPriv"]();
  });

  xfaNewDoc = app["trustPropagatorFunction"](function () {
    app["beginPriv"]();
    return app["newDoc"]();
    app["endPriv"]();
  });

  xfaTrustedNewDoc = app["trustedFunction"](function () {
    app["beginPriv"]();
    return xfaNewDoc();
    app["endPriv"]();
  });

  xfaSaveAs = app["trustPropagatorFunction"](function(doc, path){
    app["beginPriv"]();
    doc["saveAs"]({cPath: path});
    app["endPriv"]()});

  xfaTrustedSaveAs = app["trustedFunction"](function (doc, path) {
    app["beginPriv"]();
    xfaSaveAs(doc, path);
    app["endPriv"]();
  });

  xfaAlert = app.trustedFunction(function(altMessage)
  {
      app.beginPriv();
      app.alert(altMessage);
      app.endPriv();
  });

  xfaResponse = app.trustedFunction(function(responseMessage)
  {
      app.beginPriv();
      app.response(responseMessage);
      app.endPriv();
  });

  var xfaDialog = app.trustedFunction(function(dialogObject)
  {
     app.beginPriv();
     return app.execDialog(dialogObject);
     app.endPriv();
  });