// INSTALLATION INSTRUCTIONS:
// - Uses Windows environment variables %variable%
// - Temp files created in %LOCALAPPDATA%\Temp\acrobat_sbx
// - Save this file to %APPDATA%\Adobe\Acrobat\Privileged\DC\Javascripts
// - This creates the menu button in the add-ons tool menu

// ============================================================================
// MENU INITIALIZATION
// ============================================================================
// Create menu item or toolbar button depending on Acrobat version
if (app.viewerVersion < 10) {
    // For older versions (< 10), add a menu item
    app.addMenuItem({
        cName: "MailMerge",
        cUser: "Mail Merge",
        cParent: "Tools",
        cExec: "mailMergeFiles(this)",
        cEnable: "event.rc = (event.target != null);"
    });
} else {
    // For newer versions (>= 10), add a toolbar button
    app.addToolButton({
        cName: "MailMerge",
        cLabel: "Mail Merge",
        cTooltext: "Mail Merge",
        cExec: "mailMergeFiles(this)",
        cEnable: "event.rc = (event.target != null);"
    });
}

// ============================================================================
// DYNAMIC DIALOG OBJECT (Not used in Demo version)
// ============================================================================
// This dialog would allow users to select which CSV columns to include
var dynamicDlg = { 
    fileHeaders: [],
    
    // Initialize dialog with default values
    initialize: function(dialog) {
        this.loadDefaults(dialog);
    },
    
    // Process dialog results when user clicks OK
    commit: function(dialog) {
        var headerResults = dialog.store();
        this.fileHeaders = this.getFileHeaders(headerResults);
    },
    
    // Set all checkboxes to checked by default
    loadDefaults: function(dialog) {
        // Count the number of checkboxes in the dynamicDlg description
        var checkboxCount = 0;
        dynamicDlg.description.elements[0].elements.forEach(element => {
            if (element.type === "check_box") {
                checkboxCount++;
            }
        });
        
        // Create an object to hold the default values for the checkboxes
        var defaults = {};
        for (var i = 0; i < checkboxCount; i++) {
            defaults["chk" + i] = true;
        }
        
        // Load the defaults into the dialog
        dialog.load(defaults);
    },
    
    // Extract selected headers from checkbox results
    getFileHeaders: function(results) {
        var headers = [];
        for (var x in results) {
            // Only include checked checkboxes
            if (x.startsWith("chk") && results[x]) {
                console.println(x);
                headers.push(x);
            }
        }
        return headers;
    },
    
    // Dialog layout definition
    description: {
        name: "Dynamic Dialog", 
        elements: [
            {
                type: "view", 
                elements: [
                    { 
                        name: "Select which headings to include:", 
                        type: "static_text", 
                    },
                ],
            },
            { 
                type: "ok_cancel",
                alignment: "align_right",
            },
        ]
    }
};

// ============================================================================
// MAIL MERGE DIALOG OBJECT
// ============================================================================
// Dialog for selecting CSV delimiter type
var mailMergeDialog = {
    radioValue: null,
    
    // Initialize dialog with comma delimiter selected by default
    initialize: function (dialog) {
        dialog.load({
            "rd01": true,  // Comma is default
        });
    },
    
    // Process dialog results and store selected delimiter
    commit: function (dialog) {
        var dialogResults = dialog.store();
        this.radioValue = this.getRadioValue(dialogResults);
    },
    
    // Convert radio button selection to delimiter character
    getRadioValue: function (results) {
        var delimiter;
        for (var i = 1; i <= 3; i++) {
            if (results["rd0" + i]) {
                switch (i) {
                    case 1:
                        delimiter = ",";  // Comma
                        break;
                    case 2:
                        delimiter = "\t";  // Tab
                        break;
                }
            }
        }
        return delimiter;
    },
    
    // Dialog layout definition
    description: {
        name: "Mail Merge",
        elements: [
            {
                type: "view",
                elements: [
                    {
                        item_id: "str1",
                        name: "Mail Merge",
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
                        elements: [
                            {
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

// ============================================================================
// MAIN MAIL MERGE FUNCTION
// ============================================================================
/**
 * Main function to perform mail merge operation
 * @param {Object} doc - The current PDF document (can be 'app' or 'this')
 */
function mailMergeFiles(doc) {
    // Check if document is XFA-based (LiveCycle Designer)
    if (doc.xfa != null) {
        myAlert("Error! This file was created using LiveCycle Designer and therefore can't be mail-merged using this script.");
        return;
    }
    
    // Show dialog to select delimiter type
    var delimiterValue;
    if ("ok" == myDialog(mailMergeDialog)) {
        delimiterValue = mailMergeDialog.radioValue;
    }
    
    // Exit if user cancelled or no delimiter selected
    if (delimiterValue == null || delimiterValue == "") {
        return;
    }
    
    // Read and parse CSV file
    var fileDataArray = readCSV(delimiterValue, null);
    if (fileDataArray == null || fileDataArray.length == 0) {
        myAlert("No items found in the input file.");
        return;
    }
    
    // COMMENTED OUT: Dynamic dialog for selecting which columns to include
    // This feature is not used in the demo version
    // for (var i = 0; i < fileDataArray[0].length; i++) {
    //     var obj = { type: "check_box", item_id: "chk" + i, name: fileDataArray[0][i] };
    //     dynamicDlg.description.elements[0].elements.push(obj);
    // }
    // if ("ok" == myDialog(dynamicDlg)) {
    //     console.println(dynamicDlg.fileHeaders);
    // }
    // for (var i = 0; i < fileDataArray[0].length; i++) {
    //     dynamicDlg.description.elements[0].elements.pop();
    // }
    
    // COMMENTED OUT: Custom file naming pattern
    // var fileNamePattern = myResponse("Enter the file-name pattern for the mail-merged files.\nYou can use a field's name in a tag, for example:\n\t\"Letter to <FirstName> <LastName>\"\n(leave empty to use the row number: 1.pdf, 2.pdf, etc.)", "Mail Merge");
    // if (fileNamePattern == null) {
    //     fileNamePattern = "";
    // }
    
    // Initialize counters and flags
    var mergedFileCount = 0;
    var hasErrors = false;
    
    // Store original document information
    var originalDocName = this.documentFileName;
    var originalPageCount = this.numPages;
    
    // Create a copy of the document
    var re = /\.pdf$/i;
    var copyDocName = this.documentFileName.replace(re, "-copy.pdf");
    var documentPath = doc.path.replace(doc.documentFileName, "");
    var copyDocPath = documentPath + copyDocName;
    myTrustedSaveAs(doc, copyDocPath);
    
    // ========================================================================
    // MAIN LOOP: Process each row in CSV file
    // ========================================================================
    // Start from row 1 (skip header row at index 0)
    for (var rowIndex = 1; rowIndex < fileDataArray.length; rowIndex++) {
        // Reset all form fields to empty
        doc.resetForm();
        
        // Loop through each column in the current row
        for (var colIndex = 0; colIndex < fileDataArray[0].length; colIndex++) {
            // Break if this row has fewer columns than the header
            if (colIndex >= fileDataArray[rowIndex].length) {
                break;
            }
            
            // Get the field name from the header row
            var fieldName = fileDataArray[0][colIndex];
            
            // Get the field object from the PDF
            var field = doc.getField(fieldName);
            if (field == null) {
                console.println("Error! Can't locate a field called: " + fieldName);
                hasErrors = true;
                continue;
            }
            
            // Get the value for this cell and clean up quotes
            var cellValue = fileDataArray[rowIndex][colIndex]
                .replace(/^"/, "")      // Remove leading quote
                .replace(/"$/, "")      // Remove trailing quote
                .replace(/""/, "\"");   // Replace double quotes with single quote
            
            // Handle different field types
            if (field.type == "button") {
                // For button fields
                if (field.buttonPosition == position.textOnly) {
                    // Set button caption if text-only
                    field.buttonSetCaption(cellValue);
                } else {
                    // Import image as button icon
                    try {
                        myTrustedButtonImportIcon(field, convertRealPathToPDFPath(cellValue));
                    } catch (e) {
                        console.println("Error! Can't apply the image \"" + cellValue + "\" as the icon of the field: " + fieldName);
                        hasErrors = true;
                        continue;
                    }
                }
            } else {
                // For all other field types, set the value directly
                try {
                    field.value = cellValue;
                } catch (e) {
                    console.println("Error! Can't apply the value \"" + cellValue + "\" to the field: " + fieldName);
                    hasErrors = true;
                    continue;
                }
            }
            
            // COMMENTED OUT: Custom file naming logic
            // if (fileNamePattern != "") {
            //     fileName = replaceTagWithText(fileName, fieldName, cellValue);
            // }
        }
        
        // COMMENTED OUT: File name sanitization
        // if (fileNamePattern == "") {
        //     fileName = rowIndex;
        // } else {
        //     fileName = fileName.replace(/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
        // }
        
        // Flatten the current page (convert form fields to static content)
        doc.flattenPages();
        
        // Insert pages from original document if needed
        // This ensures we have enough pages for all records
        if (doc.numPages < originalPageCount * (fileDataArray.length - 1)) {
            myTrustedinsert(doc, originalDocName);
        }
        
        mergedFileCount++;
    }
    
    // Save the final merged document
    myTrustedSaveAs(doc, doc.path);
    doc.closeDoc(true);
    
    // Show completion message
    if (hasErrors) {
        myAlert("Done. " + mergedFileCount + " merged files were created, but there were some errors. See the Console for more details.", 1);
        console.show();
    } else {
        myAlert({cMsg: "Done. " + mergedFileCount + " merged files were created.", nIcon: 4});
    }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Replace tagged placeholders in a string with actual values
 * @param {string} text - The text containing tags like <FieldName>
 * @param {string} tagName - The tag name to replace
 * @param {string} replacement - The value to replace the tag with
 * @returns {string} Text with tags replaced
 */
function replaceTagWithText(text, tagName, replacement) {
    var regex = new RegExp("<" + tagName + ">", "g");
    return text.replace(regex, replacement);
}

/**
 * Remove illegal characters from file path/name
 * @param {string} text - The string to sanitize
 * @returns {string} Sanitized string safe for file paths
 */
function makeStringFilePathSafe(text) {
    return text.replace(/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
}

/**
 * Read and parse a CSV file
 * @param {string} delimiter - The delimiter character (comma, tab, etc.)
 * @param {*} unused - Unused parameter (kept for compatibility)
 * @returns {Array} 2D array of CSV data
 */
function readCSV(delimiter, unused) {
    // Default to semicolon if no delimiter specified
    if (delimiter == null) {
        delimiter = ";";
    }
    
    var csvContent = "";
    
    // Create a temporary document to import the CSV file
    var tempDoc = myTrustedNewDoc();
    var csvFileName = "file.csv";
    myTrustedImportDataObject(tempDoc, csvFileName);
    
    // Verify the file was imported
    if (tempDoc.dataObjects == null || tempDoc.dataObjects.length != 1) {
        return;
    }
    
    // Read the CSV file contents
    var csvStream = tempDoc.getDataObjectContents(csvFileName);
    csvContent = util.stringFromStream({
        oStream: csvStream,
        oCharSet: "utf-8"
    });
    
    // Normalize line endings to \n
    csvContent = csvContent.replace(/\r\n/g, "\n");
    csvContent = csvContent.replace(/\r/g, "\n");
    
    // Close the temporary document
    tempDoc.closeDoc(true);
    
    // Parse CSV into 2D array
    var parsedData = new Array();
    var lines = csvContent.split("\n");
    
    for (var i = 0; i < lines.length; i++) {
        // Clean up line (remove line breaks and BOM)
        var line = lines[i].replace(/[\r\n]/g, "").replace("﻿", "");
        
        if (line != "") {
            if (delimiter == ",") {
                // Use special CSV parser for comma-delimited files (handles quoted values)
                parsedData.push(splitCSV(line, delimiter));
            } else {
                // Simple split for other delimiters
                parsedData.push(line.split(delimiter));
            }
        }
    }
    
    return parsedData;
}

/**
 * Split a CSV line while properly handling quoted values
 * @param {string} text - The CSV line to split
 * @param {string} delimiter - The delimiter character
 * @returns {Array} Array of values from the CSV line
 */
function splitCSV(text, delimiter) {
    delimiter = delimiter || ",";
    var values = text.split(delimiter);
    var lastIndex = values.length - 1;
    var temp;
    
    // Process array backwards to handle quoted values that contain delimiters
    for (var i = lastIndex; i >= 0; i--) {
        // Check if value ends with a quote
        if (values[i].replace(/"\s+$/, "\"").charAt(values[i].length - 1) == "\"") {
            temp = values[i].replace(/^\s+"/, "\"");
            
            // Check if it's a properly quoted value
            if (temp.length > 1 && temp.charAt(0) == "\"") {
                // Remove surrounding quotes and unescape double quotes
                values[i] = values[i].replace(/^\s*"|"\s*$/g, "").replace(/""/g, "\"");
            } else {
                // This quote is part of a value that was split - rejoin it
                if (i) {
                    values.splice(i - 1, 2, [values[i - 1], values[i]].join(delimiter));
                } else {
                    values = values.shift().split(delimiter).concat(values);
                }
            }
        } else {
            // Unescape double quotes in unquoted values
            values[i].replace(/""/g, "\"");
        }
    }
    
    return values;
}

/**
 * Convert Windows/Mac file paths to PDF-compatible paths
 * @param {string} path - The file path to convert
 * @returns {string} PDF-compatible path
 */
function convertRealPathToPDFPath(path) {
    if (app.platform == "WIN") {
        // Handle Windows paths
        if (path.charAt(0) == "\\") {
            // UNC path (\\server\share)
            return path.replace(/^\\/, "//").replace(/\\/g, "/");
        } else {
            // Regular path (C:\folder\file)
            return "/" + path.replace(":\\", "/").replace(/\\/g, "/");
        }
    }
    
    if (app.platform == "MAC") {
        // Handle Mac paths
        path = path.replace(/:/g, "/");
        if (/^\//.test(path) == false) {
            path = "/" + path;
        }
        return path;
    }
    
    return path;
}

// ============================================================================
// TRUSTED FUNCTION WRAPPERS
// ============================================================================
// These functions wrap privileged operations to work around Acrobat security

/**
 * Safe wrapper for buttonImportIcon (trust propagator)
 */
safeButtonImportIcon = app.trustPropagatorFunction(function (field, iconPath) {
    app.beginPriv();
    return field.buttonImportIcon(iconPath);
    app.endPriv();
});

/**
 * Trusted function to import button icon
 */
myTrustedButtonImportIcon = app.trustedFunction(function (field, iconPath) {
    app.beginPriv();
    return safeButtonImportIcon(field, iconPath);
    app.endPriv();
});

/**
 * Safe wrapper for importDataObject (trust propagator)
 */
myImportDataObject = app.trustPropagatorFunction(function (doc, objectName) {
    app.beginPriv();
    doc.importDataObject({
        cName: objectName
    });
    app.endPriv();
});

/**
 * Trusted function to import data object
 */
myTrustedImportDataObject = app.trustedFunction(function (doc, objectName) {
    app.beginPriv();
    myImportDataObject(doc, objectName);
    app.endPriv();
});

/**
 * Safe wrapper for creating new document (trust propagator)
 */
myNewDoc = app.trustPropagatorFunction(function () {
    app.beginPriv();
    return app.newDoc();
    app.endPriv();
});

/**
 * Trusted function to create new document
 */
myTrustedNewDoc = app.trustedFunction(function () {
    app.beginPriv();
    return myNewDoc();
    app.endPriv();
});

/**
 * Safe wrapper for saveAs (trust propagator)
 */
safeSaveAs = app.trustPropagatorFunction(function(doc, path) {
    app.beginPriv();
    doc.saveAs({cPath: path});
    app.endPriv();
});

/**
 * Trusted function to save document
 */
myTrustedSaveAs = app.trustedFunction(function (doc, path) {
    app.beginPriv();
    safeSaveAs(doc, path);
    app.endPriv();
});

/**
 * Trusted function to insert pages from another document
 */
myTrustedinsert = app.trustedFunction(function (doc, sourcePath) {
    app.beginPriv();
    doc.insertPages(doc.numPages - 1, sourcePath);
    app.endPriv();
});

/**
 * Trusted function to show alert dialog
 */
myAlert = app.trustedFunction(function(message) {
    app.beginPriv();
    app.alert(message);
    app.endPriv();
});

/**
 * Trusted function to show response dialog (text input)
 */
myResponse = app.trustedFunction(function(message) {
    app.beginPriv();
    app.response(message);
    app.endPriv();
});

/**
 * Trusted function to execute custom dialog
 */
var myDialog = app.trustedFunction(function(dialogObject) {
    app.beginPriv();
    return app.execDialog(dialogObject);
    app.endPriv();
});