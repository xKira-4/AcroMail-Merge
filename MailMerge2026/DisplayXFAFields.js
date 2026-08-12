// ============================================================================
// ExportXFAFields.js
// Folder-level script for displaying XFA field data in a dialog
// ============================================================================

// ----------------------------------------------------------------------------
// MAIN DISPLAY FUNCTION
// ----------------------------------------------------------------------------

var DisplayXFAFields_Trusted = app.trustedFunction(function(doc) {
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
        
        // Create dialog
        var dialog = {
            initialize: function(dialog) {
                dialog.load({"text": csvContent});
            },
            commit: function(dialog) {
                var results = dialog.store();
            },
            description: {
                name: "XFA Field Data Export",
                align_children: "align_left",
                width: 700,
                height: 500,
                elements: [
                    {
                        type: "view",
                        align_children: "align_left",
                        elements: [
                            {
                                type: "static_text",
                                name: "FieldCount",
                                item_id: "count",
                                alignment: "align_left",
                                font: "dialog",
                                bold: true,
                                width: 680,
                                text: "Total Fields: " + fieldCount
                            },
                            {
                                type: "static_text",
                                name: "Instructions",
                                item_id: "inst",
                                alignment: "align_fill",
                                font: "dialog",
                                width: 680,
                                char_width: 85,
                                text: "Select all text below (Ctrl+A), copy (Ctrl+C), and paste into a text editor or Excel:"
                            },
                            {
                                type: "edit_text",
                                item_id: "text",
                                alignment: "align_fill",
                                width: 680,
                                height: 380,
                                multiline: true,
                                char_width: 85,
                                readonly: false
                            }
                        ]
                    },
                    {
                        type: "ok_cancel",
                        ok_name: "Close"
                    }
                ]
            }
        };
        
        // Display the dialog
        app.execDialog(dialog);
        
    } catch(e) {
        app.alert({
            cMsg: "Error during display:\n\n" + e.message,
            cTitle: "Display Error",
            nIcon: 0
        });
    }
    
    app.endPriv();
});

// ----------------------------------------------------------------------------
// USER-CALLABLE WRAPPER FUNCTION
// ----------------------------------------------------------------------------

function DisplayXFAFields() {
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
    
    DisplayXFAFields_Trusted(doc);
}

// ----------------------------------------------------------------------------
// MENU ITEM
// ----------------------------------------------------------------------------

try {
    app.addMenuItem({
        cName: "DisplayXFAFields",
        cUser: "Display XFA Fields",
        cParent: "Tools",
        cExec: "DisplayXFAFields.call(this);",
        cEnable: "event.rc = (this.xfa && typeof this.xfa.host !== 'undefined');"
    });
} catch(e) {
    // Menu item already exists or couldn't be added
}

// ----------------------------------------------------------------------------
// INITIALIZATION
// ----------------------------------------------------------------------------

console.println("ExportXFAFields.js loaded successfully. Use DisplayXFAFields() to display field data.");