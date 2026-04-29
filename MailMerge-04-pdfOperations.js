/**
 * ============================================================================
 * MAIL MERGE - PDF OPERATIONS MODULE
 * ============================================================================
 */

/**
 * Validate that the document can be used for mail merge
 */
MailMerge.PDF.validateDocument = function(doc) {
    try {
        // Check if document is XFA-based (LiveCycle Designer)
        if (doc.xfa != null) {
            MailMerge.Trusted.alert(
                "Error! This file was created using LiveCycle Designer and " +
                "therefore can't be mail-merged using this script."
            );
            return false;
        }
        
        // Check if document has form fields
        var fieldCount = 0;
        try {
            fieldCount = doc.numFields;
        } catch (e) {
            // If numFields throws an error, try another method
            try {
                var testField = doc.getField(doc.getNthFieldName(0));
                if (testField) {
                    fieldCount = 1;
                }
            } catch (e2) {
                fieldCount = 0;
            }
        }
        
        if (fieldCount === 0) {
            MailMerge.Trusted.alert(
                "Error! This PDF does not contain any form fields to merge data into.\n\n" +
                "Please ensure your PDF has form fields with names matching your CSV column headers."
            );
            return false;
        }
        
        return true;
        
    } catch (error) {
        MailMerge.Utilities.log("Validation error: " + error.toString(), "ERROR");
        MailMerge.Trusted.alert(
            "Error validating document: " + error.message
        );
        return false;
    }
};

/**
 * Get all field names from the document
 * @param {Object} doc - The PDF document
 * @returns {Array} Array of field names
 */
MailMerge.PDF.getFieldNames = function(doc) {
    var fieldNames = [];
    
    try {
        var fieldCount = doc.numFields;
        
        for (var i = 0; i < fieldCount; i++) {
            var fieldName = doc.getNthFieldName(i);
            if (fieldName) {
                fieldNames.push(fieldName);
            }
        }
        
        // Sort alphabetically for easier browsing
        fieldNames.sort(function(a, b) {
            return a.toLowerCase().localeCompare(b.toLowerCase());
        });
        
    } catch (error) {
        MailMerge.Utilities.log("Error getting field names: " + error.toString(), "ERROR");
    }
    
    return fieldNames;
};

/**
 * Perform the mail merge operation
 * @param {Object} doc - The PDF document
 * @param {Array} csvData - The parsed CSV data
 * @param {Object} fieldMap - Mapping of CSV columns to PDF fields (optional)
 * @returns {Object} Result object with count and error information
 */
/**
 * Perform the mail merge operation
 */
MailMerge.PDF.performMerge = function(doc, csvData, fieldMap) {
    var result = {
        count: 0,
        hasErrors: false,
        errors: []
    };
    
    try {
        var originalDocName = doc.documentFileName;
        var originalPageCount = doc.numPages;
        
        var copyPath = MailMerge.PDF.createBackup(doc);
        var headers = csvData[0];
        
        for (var rowIndex = 1; rowIndex < csvData.length; rowIndex++) {
            try {
                doc.resetForm();
                
                var rowResult = MailMerge.PDF.fillFields(
                    doc,
                    headers,
                    csvData[rowIndex],
                    fieldMap
                );
                
                if (rowResult.hasErrors) {
                    result.hasErrors = true;
                    result.errors = result.errors.concat(rowResult.errors);
                }
                
                doc.flattenPages();
                
                if (doc.numPages < originalPageCount * (csvData.length - 1)) {
                    MailMerge.Trusted.insertPages(doc, originalDocName);
                }
                
                result.count++;
                
            } catch (error) {
                result.hasErrors = true;
                result.errors.push("Row " + rowIndex + ": " + error.message);
                MailMerge.Utilities.log(
                    "Error processing row " + rowIndex + ": " + error.toString(),
                    "ERROR"
                );
            }
        }
        
        MailMerge.Trusted.saveAs(doc, doc.path);
        doc.closeDoc(true);
        
    } catch (error) {
        result.hasErrors = true;
        result.errors.push("Fatal error: " + error.message);
        MailMerge.Utilities.log("Fatal merge error: " + error.toString(), "ERROR");
    }
    
    return result;
};

/**
 * Fill PDF form fields with data from a CSV row
 * @param {Object} doc - The PDF document
 * @param {Array} headers - Column headers from CSV
 * @param {Array} rowData - Data values for current row
 * @param {Object} fieldMap - Mapping of CSV columns to PDF fields (optional)
 * @returns {Object} Result with hasErrors flag and errors array
 */
/**
 * Fill PDF form fields with data from a CSV row
 */
MailMerge.PDF.fillFields = function(doc, headers, rowData, fieldMap) {
    var result = {
        hasErrors: false,
        errors: []
    };
    
    for (var colIndex = 0; colIndex < headers.length; colIndex++) {
        if (colIndex >= rowData.length) {
            break;
        }
        
        var csvColumnName = headers[colIndex];
        var cellValue = MailMerge.Utilities.cleanCSVValue(rowData[colIndex]);
        
        // Determine which PDF field to use
        var pdfFieldName;
        if (fieldMap && fieldMap[csvColumnName]) {
            // Use custom mapping
            pdfFieldName = fieldMap[csvColumnName];
        } else {
            // Use CSV column name as field name (original behavior)
            pdfFieldName = csvColumnName;
        }
        
        // Skip if this column is not mapped to any field
        if (!pdfFieldName) {
            continue;
        }
        
        try {
            var field = doc.getField(pdfFieldName);
            if (!field) {
                result.hasErrors = true;
                result.errors.push("Field not found: " + pdfFieldName);
                continue;
            }
            
            MailMerge.PDF.setFieldValue(field, cellValue, pdfFieldName);
            
        } catch (error) {
            result.hasErrors = true;
            result.errors.push(
                "Error setting field '" + pdfFieldName + "' to '" + cellValue + "': " + error.message
            );
        }
    }
    
    return result;
};

/**
 * Set a field's value based on its type
 */
MailMerge.PDF.setFieldValue = function(field, value, fieldName) {
    if (field.type === "button") {
        if (field.buttonPosition === position.textOnly) {
            field.buttonSetCaption(value);
        } else {
            var imagePath = MailMerge.Utilities.convertPathToPDFFormat(value);
            MailMerge.Trusted.buttonImportIcon(field, imagePath);
        }
    } else {
        field.value = value;
    }
};

/**
 * Create a backup copy of the document
 */
MailMerge.PDF.createBackup = function(doc) {
    var originalName = doc.documentFileName;
    var copyName = originalName.replace(/\.pdf$/i, "-copy.pdf");
    var documentPath = doc.path.replace(doc.documentFileName, "");
    var copyPath = documentPath + copyName;
    
    MailMerge.Trusted.saveAs(doc, copyPath);
    
    return copyPath;
};

console.println("MailMerge.PDF loaded");
