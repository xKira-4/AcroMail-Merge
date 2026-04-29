/**
 * ============================================================================
 * MAIL MERGE - DIALOGS MODULE
 * ============================================================================
 */

/**
 * Dialog for selecting CSV delimiter type and mapping option
 */
MailMerge.Dialogs.delimiterDialog = {
    radioValue: null,
    showMapping: true,
    
    initialize: function(dialog) {
        dialog.load({
            "rdDelimComma": true,   // Comma is default
            "rdDelimTab": false,
            "rdDelimComma": true,      // Show mapping is default
            "rdMapSkip": false
        });
    },
    
    commit: function(dialog) {
        var oRslt = dialog.store();
        
        // Get delimiter selection
        if (oRslt["rdDelimComma"] === true) {
            this.radioValue = ",";
        } else if (oRslt["rdDelimTab"] === true) {
            this.radioValue = "\t";
        }
        
        // Get mapping selection
        if (oRslt["rdMapShow"] === true) {
            this.showMapping = true;
        } else if (oRslt["rdMapSkip"] === true) {
            this.showMapping = false;
        }
    },
    
    description: {
        name: "Mail Merge - Options",
        elements: [
            {
                type: "view",
                elements: [
                    {
                        type: "static_text",
                        name: "Mail Merge",
                        font: "dialog",
                        bold: true,
                        alignment: "align_center",
                        char_width: 50,
                        height: 20,
                    },
                    {
                        type: "gap",
                        height: 10
                    }
                ]
            },
            
            // Delimiter section
            {
                type: "cluster",
                name: "CSV Delimiter:",
                bold: true,
                elements: [
                    {
                        type: "radio",
                        item_id: "rdDelimComma",
                        name: "Comma (,)"
                    },
                    {
                        type: "radio",
                        item_id: "rdDelimTab",
                        name: "Tab"
                    }
                ]
            },
            
            {
                type: "gap",
                height: 10
            },
            
            // Field mapping section
            {
                type: "cluster",
                name: "Field Mapping:",
                bold: true,
                elements: [
                    {
                        type: "radio",
                        item_id: "rdMapShow",
                        name: "Show field mapping dialog"
                    },
                    {
                        type: "static_text",
                        name: "   (Allows custom column-to-field mapping)",
                        font: "palette"
                    },
                    {
                        type: "radio",
                        item_id: "rdMapSkip",
                        name: "Use exact name matching"
                    },
                    {
                        type: "static_text",
                        name: "   (CSV headers must match PDF field names)",
                        font: "palette"
                    }
                ]
            },
            
            {
                type: "ok_cancel",
                alignment: "align_right"
            }
        ]
    }
};

/**
 * Show the delimiter selection dialog
 * @returns {Object|null} Object with {delimiter, showMapping} properties, or null if cancelled
 */
MailMerge.Dialogs.showDelimiterDialog = function() {
    if ("ok" === MailMerge.Trusted.execDialog(MailMerge.Dialogs.delimiterDialog)) {
        return {
            delimiter: MailMerge.Dialogs.delimiterDialog.radioValue,
            showMapping: MailMerge.Dialogs.delimiterDialog.showMapping
        };
    }
    return null;
};

/**
 * ============================================================================
 * FIELD MAPPING DIALOG - CORRECTED FORMAT
 * ============================================================================
 */

/**
 * Dialog for mapping CSV columns to PDF fields
 */
MailMerge.Dialogs.fieldMappingDialog = {
    fieldMap: {},
    csvHeaders: [],
    pdfFields: [],
    csvData: [],
    popupLists: [],  // Store the popup list objects for each column
    
    /**
     * Helper function to set selected item in a list
     */
    SetListSel: function(oList, cSel) {
        for (var i in oList) {
            oList[i] = (i === cSel) ? 1 : -1;  // Changed 0 to 1
        }
    },
    
    /**
     * Helper function to get selected item from a list
     */
    GetListSel: function(oList) {
        if (!oList) return null;
        
        for (var i in oList) {
            // Selected items have value 1 (not 0)
            if (oList[i] === 1) {
                return i;
            }
        }
        return null;
    },
    
    /**
     * Build a popup list object for a specific CSV column
     */
    buildPopupList: function(csvHeader) {
        var list = {};
        
        // First option: Skip
        list["(Skip this column)"] = -1;
        
        // Add all PDF fields
        for (var i = 0; i < this.pdfFields.length; i++) {
            list[this.pdfFields[i]] = -1;
        }
        
        // Try to find best match and select it
        var bestMatch = this.findBestMatchName(csvHeader);
        if (bestMatch) {
            this.SetListSel(list, bestMatch);
        } else {
            this.SetListSel(list, "(Skip this column)");
        }
        
        return list;
    },
    
    /**
     * Find best matching PDF field name for a CSV column
     */
    findBestMatchName: function(csvHeader) {
        var csvLower = csvHeader.toLowerCase().trim();
        
        // Try exact match (case-insensitive)
        for (var i = 0; i < this.pdfFields.length; i++) {
            if (this.pdfFields[i].toLowerCase().trim() === csvLower) {
                return this.pdfFields[i];
            }
        }
        
        // Try partial match
        for (var i = 0; i < this.pdfFields.length; i++) {
            var fieldLower = this.pdfFields[i].toLowerCase().trim();
            if (fieldLower.indexOf(csvLower) !== -1 || csvLower.indexOf(fieldLower) !== -1) {
                return this.pdfFields[i];
            }
        }
        
        return null;
    },
    
    /**
     * Initialize dialog with default mappings
     */
    initialize: function(dialog) {
        var defaults = {};
        
        // Build popup lists for each CSV column
        this.popupLists = [];
        for (var i = 0; i < this.csvHeaders.length; i++) {
            var list = this.buildPopupList(this.csvHeaders[i]);
            this.popupLists.push(list);
            defaults["map" + i] = list;
        }
        
        dialog.load(defaults);
    },
    
    /**
     * Process dialog results when user clicks OK
     */
    commit: function(dialog) {
        var results = dialog.store();
        this.fieldMap = {};
        
        for (var i = 0; i < this.csvHeaders.length; i++) {
            var csvHeader = this.csvHeaders[i];
            var selectedList = results["map" + i];
            var selectedField = this.GetListSel(selectedList);
            
            // Only add to map if not skipping
            if (selectedField && selectedField !== "(Skip this column)") {
                this.fieldMap[csvHeader] = selectedField;
            }
        }
        
        // Show the final mapping in console
        console.println("=== Field Mapping Created ===");
        for (var csv in this.fieldMap) {
            console.println("  CSV: '" + csv + "' -> PDF: '" + this.fieldMap[csv] + "'");
        }
        console.println("=============================");
    },
    
/**
 * Build preview section showing first data row
 */
buildPreviewSection: function() {
    var previewElements = [
        {
            type: "static_text",
            name: "Preview of first data row:",
            font: "dialog",
            bold: true,
            char_width: 60
        }
    ];
    
    // Check if we have data
    if (!this.csvData || this.csvData.length < 2) {
        previewElements.push({
            type: "static_text",
            name: "(No data rows found)",
            font: "palette",
            char_width: 60
        });
        return previewElements;
    }
    
    // Check if we have headers
    if (!this.csvHeaders || this.csvHeaders.length === 0) {
        previewElements.push({
            type: "static_text",
            name: "(No column headers found)",
            font: "palette",
            char_width: 60
        });
        return previewElements;
    }
    
    var firstRow = this.csvData[1];  // Row 0 is headers, row 1 is first data
    var previewText = "";
    var maxCols = Math.min(3, this.csvHeaders.length);
    
    for (var i = 0; i < maxCols; i++) {
        if (i > 0) previewText += " | ";
        
        var header = this.csvHeaders[i] || "Column" + i;
        var value = "";
        
        if (firstRow && i < firstRow.length) {
            value = firstRow[i] || "";
        }
        
        // Truncate long values
        if (value.length > 15) {
            value = value.substring(0, 15) + "...";
        }
        
        previewText += header + ": \"" + value + "\"";
    }
    
    if (this.csvHeaders.length > 3) {
        previewText += " ...";
    }
    
    previewElements.push({
        type: "static_text",
        name: previewText,
        font: "palette",
        char_width: 60
    });
    
    return previewElements;
},
    
    /**
     * Build the dialog description dynamically
     */
    buildDescription: function() {
        var elements = [
            {
                type: "view",
                elements: [
                    // Header section
                    {
                        type: "view",
                        align_children: "align_left",
                        elements: [
                            {
                                type: "static_text",
                                name: "Map CSV Columns to PDF Fields",
                                font: "dialog",
                                bold: true,
                                char_width: 60
                            },
                            {
                                type: "static_text",
                                name: "Select which PDF field each CSV column should fill:",
                                char_width: 60
                            },
                            {
                                type: "gap",
                                height: 5
                            }
                        ]
                    },
                    
                    // Preview section
                    {
                        type: "view",
                        align_children: "align_left",
                        elements: this.buildPreviewSection()
                    },
                    
                    {
                        type: "gap",
                        height: 10
                    },
                    
                    // Column headers
                    {
                        type: "view",
                        align_children: "align_row",
                        elements: [
                            {
                                type: "static_text",
                                name: "CSV Column",
                                font: "dialog",
                                bold: true,
                                char_width: 20,
                                alignment: "align_right"
                            },
                            {
                                type: "gap",
                                width: 10
                            },
                            {
                                type: "static_text",
                                name: "PDF Field",
                                font: "dialog",
                                bold: true,
                                char_width: 25
                            }
                        ]
                    },
                    
                    {
                        type: "gap",
                        height: 5
                    }
                ]
            }
        ];
        
        // Add a row for each CSV column
        var mappingRows = [];
        for (var i = 0; i < this.csvHeaders.length; i++) {
            mappingRows.push({
                type: "view",
                align_children: "align_row",
                elements: [
                    {
                        type: "static_text",
                        name: this.csvHeaders[i] + ":",
                        char_width: 20,
                        alignment: "align_right"
                    },
                    {
                        type: "gap",
                        width: 10
                    },
                    {
                        type: "popup",
                        item_id: "map" + i,
                        char_width: 25
                        // Note: No popup_items here - it's set via initialize()
                    }
                ]
            });
        }
        
        // Add all mapping rows to the dialog
        elements[0].elements.push({
            type: "view",
            align_children: "align_left",
            elements: mappingRows
        });
        
        // Add OK/Cancel buttons
        elements.push({
            type: "ok_cancel",
            alignment: "align_right"
        });
        
        this.description = {
            name: "Field Mapping",
            elements: elements
        };
    },
    
    description: {}
};

/**
 * Show the field mapping dialog
 */
MailMerge.Dialogs.showFieldMappingDialog = function(csvHeaders, pdfFields, csvData) {
    // Store the data in the dialog object
    MailMerge.Dialogs.fieldMappingDialog.csvHeaders = csvHeaders;
    MailMerge.Dialogs.fieldMappingDialog.pdfFields = pdfFields;
    MailMerge.Dialogs.fieldMappingDialog.csvData = csvData;
    
    // Build the dialog dynamically based on CSV columns
    MailMerge.Dialogs.fieldMappingDialog.buildDescription();
    
    // Show the dialog
    if ("ok" === MailMerge.Trusted.execDialog(MailMerge.Dialogs.fieldMappingDialog)) {
        return MailMerge.Dialogs.fieldMappingDialog.fieldMap;
    }
    
    return null;
};

/**
 * Show the field mapping dialog
 * @param {Array} csvHeaders - Array of CSV column names
 * @param {Array} pdfFields - Array of PDF field names
 * @returns {Object|null} Field mapping object or null if cancelled
 */
MailMerge.Dialogs.showFieldMappingDialog = function(csvHeaders, pdfFields) {
    // Store the data in the dialog object
    MailMerge.Dialogs.fieldMappingDialog.csvHeaders = csvHeaders;
    MailMerge.Dialogs.fieldMappingDialog.pdfFields = pdfFields;
    
    // Build the dialog dynamically based on CSV columns
    MailMerge.Dialogs.fieldMappingDialog.buildDescription();
    
    // Show the dialog
    if ("ok" === MailMerge.Trusted.execDialog(MailMerge.Dialogs.fieldMappingDialog)) {
        return MailMerge.Dialogs.fieldMappingDialog.fieldMap;
    }
    
    return null;
};

console.println("MailMerge.Dialogs loaded");
