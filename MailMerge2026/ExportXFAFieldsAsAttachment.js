// ============================================================================
// ExportXFAFields.js
// Folder-level script for exporting XFA field data to CSV
// ============================================================================

// Main trusted function
var ExportXFAFieldsToCSV_Trusted = app.trustedFunction(function(doc) {
    app.beginPriv();
    
    try {
        // Use the passed document or try to get the active document
        if (!doc) {
            doc = app.activeDocs[0];
        }
        
        // Verify this is an XFA document
        if (!doc.xfa || typeof doc.xfa.host === 'undefined') {
            app.alert({
                cMsg: "This is not an XFA form.\n\nThis script only works with XFA (XML Forms Architecture) PDFs.",
                cTitle: "Not an XFA Form",
                nIcon: 0
            });
            app.endPriv();
            return;
        }
        
        // Create output array to store all data
        var outputData = [];
        
        // Add header row
        outputData.push("Field Name,SOM Path,Tooltip,Value");
        
        var fieldCount = 0;
        
        // Collect all field data using the document's xfa object
        for (var i = 0; i < doc.xfa.host.numPages; i++) {
            var oFields = doc.xfa.layout.pageContent(i, "field");
            var nodesLength = oFields.length;
            
            for (var j = 0; j < nodesLength; j++) {
                var oItem = oFields.item(j);
                var somPath = oItem.somExpression;
                
                // Get tooltip
                var tooltip = "";
                if (oItem.assist && oItem.assist.toolTip) {
                    tooltip = oItem.assist.toolTip.value;
                } else if (oItem.assist && oItem.assist.speak) {
                    tooltip = oItem.assist.speak.value;
                }
                
                // Clean up values for CSV output (escape quotes and commas)
                var fieldName = (oItem.name || "").replace(/"/g, '""');
                var fieldValue = (oItem.rawValue || "").replace(/"/g, '""');
                tooltip = tooltip.replace(/"/g, '""').replace(/\n/g, " ").replace(/\r/g, " ");
                somPath = somPath.replace(/"/g, '""');
                
                // Wrap fields in quotes to handle commas within data
                var row = '"' + fieldName + '","' + somPath + '","' + tooltip + '","' + fieldValue + '"';
                outputData.push(row);
                fieldCount++;
            }
        }
        
        if (fieldCount === 0) {
            app.alert({
                cMsg: "No fields found in this XFA form.",
                cTitle: "No Fields",
                nIcon: 3
            });
            app.endPriv();
            return;
        }
        
        // Join all rows with newlines
        var csvContent = outputData.join("\r\n");
        
        // Create CSV filename
        var pdfName = doc.documentFileName.replace(/\.pdf$/i, "");
        var timestamp = util.printd("yyyymmdd_HHMMss", new Date());
        var csvFilename = pdfName + "_fields_" + timestamp + ".csv";
        
        // Save as attachment to PDF
        doc.createDataObject({
            cName: csvFilename,
            cValue: csvContent
        });
        
        app.alert({
            cMsg: "Field data exported successfully!\n\n" +
                  "File: " + csvFilename + "\n" +
                  "Total fields: " + fieldCount + "\n\n" +
                  "The CSV has been attached to this PDF.\n\n" +
                  "To save it:\n" +
                  "1. Click the paperclip icon in the left panel\n" +
                  "2. Right-click '" + csvFilename + "'\n" +
                  "3. Select 'Save Attachment...'",
            cTitle: "Export Complete",
            nIcon: 3
        });
        
    } catch(e) {
        app.alert({
            cMsg: "Error during export:\n\n" + e.message,
            cTitle: "Export Error",
            nIcon: 0
        });
    }
    
    app.endPriv();
});

// Console-callable wrapper function
function ExportXFAFieldsToCSV() {
    var doc;
    
    // Try multiple ways to get the document
    if (typeof this.xfa !== 'undefined' && this.xfa !== null) {
        doc = this;
    } else if (app.activeDocs && app.activeDocs.length > 0) {
        doc = app.activeDocs[0];
    } else {
        app.alert({
            cMsg: "No active document found.\n\nPlease open an XFA PDF first.",
            cTitle: "No Document",
            nIcon: 0
        });
        return;
    }
    
    ExportXFAFieldsToCSV_Trusted(doc);
}

// Add menu item for easy access
try {
    app.addMenuItem({
        cName: "ExportXFAFields",
        cUser: "Export XFA Fields to CSV",
        cParent: "Tools",
        cExec: "ExportXFAFieldsToCSV.call(this);",
        cEnable: "event.rc = (this.xfa && typeof this.xfa.host !== 'undefined');"
    });
} catch(e) {
    // Menu item already exists or couldn't be added
}

// Console message to confirm script loaded
console.println("ExportXFAFields.js loaded successfully. Use ExportXFAFieldsToCSV() to export.");