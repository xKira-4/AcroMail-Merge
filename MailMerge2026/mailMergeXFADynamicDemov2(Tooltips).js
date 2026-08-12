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

  // Function to extract row number from tooltip
  function extractRowNumberFromTooltip(tooltipObj) {
    if (!tooltipObj) return null;
    
    try {
      // Try to get the value property if it's an XFA object
      var tooltip = null;
      if (typeof tooltipObj === "string") {
        tooltip = tooltipObj;
      } else if (tooltipObj.value) {
        tooltip = tooltipObj.value;
      } else {
        // Try to convert to string
        tooltip = String(tooltipObj);
      }
      
      if (!tooltip || tooltip === "[object XFAObject]") {
        return null;
      }
      
      // Look for patterns like "row 19" or "row 1" in the tooltip
      var match = tooltip.match(/row\s+(\d+)/i);
      if (match && match[1]) {
        var rowNum = parseInt(match[1], 10);
        return rowNum;
      }
    } catch (e) {
      console.println("      Error extracting tooltip: " + e);
    }
    
    return null;
  }

  // Function to get all XFA fields and create a mapping organized by page
  function getXFAFieldMap(doc) {
    console.println("\n========== STARTING FIELD MAP CREATION ==========");
    var fieldMap = {};
    var pageInfo = {}; // Track max row number per page
    
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
        
        pageInfo[i] = {
          maxRow: 0,
          fieldCount: 0
        };
        
        for (var j = 0; j < nodesLength; j++) {
          var oItem = oFields.item(j);
          var fieldName = oItem.name;
          var somPath = oItem.somExpression;
          
          // Try to get tooltip text
          var tooltip = null;
          var rowNumber = null;
          try {
            if (oItem.assist && oItem.assist.toolTip) {
              if (typeof oItem.assist.toolTip === "string") {
                tooltip = oItem.assist.toolTip;
              } else if (oItem.assist.toolTip.value) {
                tooltip = oItem.assist.toolTip.value;
              }
              rowNumber = extractRowNumberFromTooltip(tooltip);
              
              // Track the maximum row number on this page
              if (rowNumber != null && rowNumber > pageInfo[i].maxRow) {
                pageInfo[i].maxRow = rowNumber;
              }
            }
          } catch (e) {
            console.println("  Error reading tooltip for field " + j + ": " + e);
          }
          
          console.println("  Field #" + j + ":");
          console.println("    Name: " + fieldName);
          console.println("    SOM Path: " + somPath);
          console.println("    Tooltip: " + tooltip);
          console.println("    Row Number: " + rowNumber);
          console.println("    Page: " + i);
          console.println("    Current Value: " + oItem.rawValue);
          
          // Store the field with its SOM path
          if (!fieldMap[fieldName]) {
            fieldMap[fieldName] = [];
          }
          
          fieldMap[fieldName].push({
            somPath: somPath,
            field: oItem,
            rowNumber: rowNumber,
            tooltip: tooltip,
            page: i
          });
          
          if (rowNumber != null) {
            pageInfo[i].fieldCount++;
          }
        }
        
        console.println("Page " + i + " summary: Max row = " + pageInfo[i].maxRow + ", Fields with row numbers = " + pageInfo[i].fieldCount);
      }
      
      console.println("\n========== FIELD MAP SUMMARY ==========");
      console.println("Total unique field names: " + Object.keys(fieldMap).length);
      for (var fname in fieldMap) {
        console.println("  '" + fname + "' has " + fieldMap[fname].length + " instance(s)");
        for (var k = 0; k < fieldMap[fname].length; k++) {
          console.println("    Instance " + k + ": page " + fieldMap[fname][k].page + ", row " + fieldMap[fname][k].rowNumber);
        }
      }
      
      console.println("\n========== PAGE CAPACITY SUMMARY ==========");
      for (var p in pageInfo) {
        console.println("  Page " + p + ": Can hold up to " + pageInfo[p].maxRow + " rows");
      }
      console.println("========================================\n");
      
    } catch (e) {
      console.println("ERROR in getXFAFieldMap: " + e);
      console.println("Stack: " + e.stack);
    }
    
    return {
      fieldMap: fieldMap,
      pageInfo: pageInfo
    };
  }

  // Function to determine which page and row number a CSV item should go to
  function calculatePageAndRow(itemIndex, pageInfo) {
    var cumulativeRows = 0;
    var pageKeys = [];
    
    // Get sorted page keys
    for (var p in pageInfo) {
      pageKeys.push(parseInt(p));
    }
    pageKeys.sort(function(a, b) { return a - b; });
    
    // Find which page this item belongs to
    for (var i = 0; i < pageKeys.length; i++) {
      var pageNum = pageKeys[i];
      var pageCapacity = pageInfo[pageNum].maxRow;
      
      if (itemIndex < cumulativeRows + pageCapacity) {
        // This item belongs on this page
        var rowOnPage = (itemIndex - cumulativeRows) + 1; // +1 because rows start at 1
        return {
          page: pageNum,
          row: rowOnPage
        };
      }
      
      cumulativeRows += pageCapacity;
    }
    
    // If we get here, we've exceeded all pages
    return null;
  }

  // Function to group CSV rows by the "To" column
  function groupCSVByRecipient(fileDataArray, toColumnIndex) {
    console.println("\n========== GROUPING CSV ROWS BY RECIPIENT ==========");
    var groups = [];
    var currentRecipient = null;
    var currentGroup = null;
    
    // Start from row 1 (skip header at row 0)
    for (var i = 1; i < fileDataArray.length; i++) {
      var row = fileDataArray[i];
      var toValue = row[toColumnIndex] ? row[toColumnIndex].replace(/^"/, "").replace(/"$/, "").replace(/""/, "\"") : "";
      
      // If "To" column has a value, start a new group
      if (toValue !== "") {
        // Save previous group if it exists
        if (currentGroup != null) {
          groups.push(currentGroup);
          console.println("  Completed group for: " + currentRecipient + " (" + currentGroup.rows.length + " items)");
        }
        
        // Start new group
        currentRecipient = toValue;
        currentGroup = {
          recipient: currentRecipient,
          rows: [row]
        };
        console.println("  Started new group for: " + currentRecipient);
      } else {
        // Add to current group
        if (currentGroup != null) {
          currentGroup.rows.push(row);
        } else {
          console.println("  WARNING: Row " + i + " has no recipient and no current group");
        }
      }
    }
    
    // Don't forget the last group
    if (currentGroup != null) {
      groups.push(currentGroup);
      console.println("  Completed group for: " + currentRecipient + " (" + currentGroup.rows.length + " items)");
    }
    
    console.println("\nTotal groups created: " + groups.length);
    console.println("====================================================\n");
    
    return groups;
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
    var mapResult = getXFAFieldMap(doc);
    var xfaFieldMap = mapResult.fieldMap;
    var pageInfo = mapResult.pageInfo;
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
    
    // Group CSV rows by recipient
    var recipientGroups = groupCSVByRecipient(fileDataArray, filenameColumnIndex);
    
    if (recipientGroups.length == 0) {
      console.println("ERROR: No recipient groups found");
      xfaAlert("Error! No valid data groups found in CSV.");
      return;
    }
    
    var filesCreated = 0;
    var hasErrors = false;
    var docPath = doc["path"]["replace"](doc["documentFileName"], "");
    console.println("Output directory: " + docPath);
    
    // Build a list of field names that are in the CSV (these are the ones we'll clear)
    var csvFieldNames = {};
    for (var h = 0; h < csvHeaders.length; h++) {
      csvFieldNames[csvHeaders[h]] = true;
    }
    
    // Calculate total capacity across all pages
    var totalCapacity = 0;
    for (var p in pageInfo) {
      totalCapacity += pageInfo[p].maxRow;
    }
    console.println("Total form capacity: " + totalCapacity + " rows across all pages");
    
    // Process each recipient group
    for (var groupIndex = 0; groupIndex < recipientGroups.length; groupIndex++) {
      var group = recipientGroups[groupIndex];
      var recipient = group.recipient;
      var rows = group.rows;
      
      console.println("\n\n########## PROCESSING GROUP " + (groupIndex + 1) + ": " + recipient + " ##########");
      console.println("Number of items: " + rows.length);
      
      // Check if we have too many rows
      if (rows.length > totalCapacity) {
        console.println("WARNING: Group has " + rows.length + " items but form only has capacity for " + totalCapacity);
        xfaAlert("Warning: " + recipient + " has " + rows.length + " items but the form can only hold " + totalCapacity + ". Some items will be skipped.");
      }
      
      // Clear only fields that are in the CSV headers
      console.println("\nClearing CSV-related fields...");
      for (var fieldName in xfaFieldMap) {
        // Only clear if this field name is in the CSV headers (but not "To")
        if (csvFieldNames[fieldName] && fieldName != "To") {
          var fieldInstances = xfaFieldMap[fieldName];
          console.println("  Clearing field: " + fieldName + " (" + fieldInstances.length + " instances)");
          for (var fi = 0; fi < fieldInstances.length; fi++) {
            try {
              fieldInstances[fi].field.rawValue = null;
            } catch (e) {
              // Ignore errors when clearing
            }
          }
        }
      }
      
      // Set the "To" field with the recipient name (if it exists in the form)
      if (xfaFieldMap["To"]) {
        console.println("\nSetting 'To' field to: " + recipient);
        var toFields = xfaFieldMap["To"];
        // Set all instances of the "To" field
        for (var ti = 0; ti < toFields.length; ti++) {
          try {
            console.println("  Setting To field instance " + ti + ": " + toFields[ti].somPath);
            toFields[ti].field.rawValue = recipient;
            console.println("  SUCCESS!");
          } catch (e) {
            console.println("  ERROR: " + e);
            hasErrors = true;
          }
        }
      }
      
      // Process each row in this group
      for (var rowIdx = 0; rowIdx < rows.length; rowIdx++) {
        var rowData = rows[rowIdx];
        
        // Calculate which page and row this item should go to
        var location = calculatePageAndRow(rowIdx, pageInfo);
        
        if (location == null) {
          console.println("\n  --- Item " + (rowIdx + 1) + " SKIPPED (exceeds form capacity) ---");
          continue;
        }
        
        var targetPage = location.page;
        var targetRow = location.row;
        
        console.println("\n  --- Processing Item " + (rowIdx + 1) + " -> Page " + targetPage + ", Row " + targetRow + " ---");
        console.println("  Row data: " + rowData.join(" | "));
        
        // Loop through each column/header
        for (var colIndex = 0; colIndex < csvHeaders["length"]; colIndex++) {
          
          if (colIndex >= rowData["length"]) {
            continue;
          }
          
          var headerName = csvHeaders[colIndex];
          var cellValue = rowData[colIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
          
          // Skip the "To" column (we already handled it above)
          if (colIndex == filenameColumnIndex) {
            continue;
          }
          
          console.println("\n    Column: '" + headerName + "' = '" + cellValue + "'");
          
          // Check if this header matches any XFA field name
          if (xfaFieldMap[headerName]) {
            var matchingFields = xfaFieldMap[headerName];
            console.println("    MATCH FOUND! " + matchingFields.length + " field(s) with this name");
            
            // Find the field that matches this page and row number
            var targetField = null;
            
            for (var f = 0; f < matchingFields.length; f++) {
              var fieldInfo = matchingFields[f];
              
              // Match based on page and tooltip row number
              if (fieldInfo.page == targetPage && fieldInfo.rowNumber != null && fieldInfo.rowNumber == targetRow) {
                targetField = fieldInfo;
                console.println("    PAGE & ROW MATCH! Page " + targetPage + ", Row " + targetRow);
                break;
              }
            }
            
            // If no match found, try just matching the row number (old behavior)
            if (targetField == null) {
              for (var f = 0; f < matchingFields.length; f++) {
                var fieldInfo = matchingFields[f];
                if (fieldInfo.rowNumber != null && fieldInfo.rowNumber == targetRow) {
                  targetField = fieldInfo;
                  console.println("    ROW MATCH (any page): Row " + targetRow);
                  break;
                }
              }
            }
            
            // If still no match, use first field
            if (targetField == null) {
              console.println("    Using first field instance as fallback");
              targetField = matchingFields[0];
            }
            
            // Set the field value
            if (targetField) {
              try {
                console.println("    Setting field: " + targetField.somPath);
                console.println("    Old value: '" + targetField.field.rawValue + "'");
                targetField.field.rawValue = cellValue;
                console.println("    New value: '" + targetField.field.rawValue + "'");
                console.println("    SUCCESS!");
              } catch (e) {
                console.println("    ERROR setting field: " + e);
                hasErrors = true;
              }
            }
          } else {
            console.println("    WARNING: No XFA field found for this CSV header");
          }
        }
      }
      
      // Create safe filename from recipient name
      var fileName = recipient["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
      if (fileName == "") {
        fileName = "output_" + (groupIndex + 1);
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
      
      console.println("########## END GROUP " + (groupIndex + 1) + " ##########\n");
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