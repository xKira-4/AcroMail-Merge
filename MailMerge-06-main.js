/**
 * ============================================================================
 * MAIL MERGE - MAIN EXECUTION
 * ============================================================================
 */

/**
 * Main function to perform mail merge operation
 */
MailMerge.execute = function(doc) {
    try {
        MailMerge.Utilities.log("Starting mail merge operation", "INFO");
        
        // Step 1: Validate document
        if (!MailMerge.PDF.validateDocument(doc)) {
            return;
        }
        
        // Step 2: Get delimiter and mapping preference from user
        var dialogResult = MailMerge.Dialogs.showDelimiterDialog();
        if (!dialogResult) {
            MailMerge.Utilities.log("User cancelled delimiter selection", "INFO");
            return;
        }
        
        var delimiter = dialogResult.delimiter;
        var showMapping = dialogResult.showMapping;
        
        // Step 3: Read CSV file
        var csvData = MailMerge.CSV.readFile(delimiter);
        if (!csvData || csvData.length === 0) {
            MailMerge.Trusted.alert("No items found in the input file.");
            return;
        }
        
        MailMerge.Utilities.log("CSV file loaded with " + (csvData.length - 1) + " data rows", "INFO");
        
        // Step 4: Determine field mapping
        var fieldMap = null;
        
        if (showMapping === true) {
            // Get PDF field names
            var pdfFields = MailMerge.PDF.getFieldNames(doc);
            if (pdfFields.length === 0) {
                MailMerge.Trusted.alert("No form fields found in the PDF.");
                return;
            }
            
            // Show field mapping dialog
            var csvHeaders = csvData[0];
            fieldMap = MailMerge.Dialogs.showFieldMappingDialog(csvHeaders, pdfFields, csvData);
            
            if (!fieldMap) {
                MailMerge.Utilities.log("User cancelled field mapping", "INFO");
                return;
            }
            
            MailMerge.Utilities.log("Using custom field mapping", "INFO");
        } else {
            // Use automatic exact matching (original behavior)
            MailMerge.Utilities.log("Using automatic exact name matching", "INFO");
        }
        
        // Step 5: Perform merge
        var result = MailMerge.PDF.performMerge(doc, csvData, fieldMap);
        
        // Step 6: Show results
        MailMerge.showResults(result);
        
        MailMerge.Utilities.log("Mail merge completed successfully", "INFO");
        
    } catch (error) {
        MailMerge.Trusted.alert("An error occurred during mail merge: " + error.message);
        MailMerge.Utilities.log("Mail Merge Error: " + error.toString(), "ERROR");
    }
};

/**
 * Display the results of the mail merge operation
 */
MailMerge.showResults = function(result) {
    if (result.hasErrors) {
        MailMerge.Trusted.alert(
            "Done. " + result.count + " merged files were created, but there were some errors. " +
            "See the Console for more details.",
            1
        );
        console.show();
    } else {
        MailMerge.Trusted.alert({
            cMsg: "Done. " + result.count + " merged files were created.",
            nIcon: 4
        });
    }
};

/**
 * Initialize the Mail Merge menu item or toolbar button
 */
MailMerge.initializeMenu = function() {
    if (app.viewerVersion < 10) {
        app.addMenuItem({
            cName: "MailMergeDEMO",
            cUser: "Mail Merge (DEMO)",
            cParent: "Tools",
            cExec: "MailMerge.execute(this)",
            cEnable: "event.rc = (event.target != null);"
        });
    } else {
        app.addToolButton({
            cName: "MailMergeDEMO",
            cLabel: "Mail Merge (DEMO)",
            cTooltext: "Mail Merge (DEMO)",
            cExec: "MailMerge.execute(this)",
            cEnable: "event.rc = (event.target != null);"
        });
    }
    
    MailMerge.Utilities.log("Mail Merge menu initialized", "INFO");
};

// Initialize on load
MailMerge.initializeMenu();

console.println("MailMerge.Main loaded - All modules ready");
