/**
 * ============================================================================
 * MAIL MERGE - CSV PARSER MODULE
 * ============================================================================
 */

/**
 * Read and parse a CSV file
 */
MailMerge.CSV.readFile = function(delimiter) {
    try {
        delimiter = delimiter || MailMerge.config.defaultDelimiter;
        
        var tempDoc = MailMerge.Trusted.newDoc();
        if (!tempDoc) {
            throw new Error("Failed to create temporary document");
        }
        
        MailMerge.Trusted.importDataObject(tempDoc, MailMerge.config.csvFileName);
        
        if (!tempDoc.dataObjects || tempDoc.dataObjects.length !== 1) {
            tempDoc.closeDoc(true);
            throw new Error("Failed to import CSV file");
        }
        
        var csvStream = tempDoc.getDataObjectContents(MailMerge.config.csvFileName);
        var csvContent = util.stringFromStream({
            oStream: csvStream,
            oCharSet: MailMerge.config.encoding
        });
        
        tempDoc.closeDoc(true);
        
        return MailMerge.CSV.parse(csvContent, delimiter);
        
    } catch (error) {
        MailMerge.Utilities.log("CSV Read Error: " + error.toString(), "ERROR");
        return null;
    }
};

/**
 * Parse CSV content into a 2D array
 */
MailMerge.CSV.parse = function(content, delimiter) {
    content = MailMerge.Utilities.normalizeLineEndings(content);
    
    var parsedData = [];
    var lines = content.split("\n");
    
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].replace(/[\r\n]/g, "").replace(/^\uFEFF/, "");
        
        if (line !== "") {
            if (delimiter === ",") {
                parsedData.push(MailMerge.CSV.splitCSVLine(line, delimiter));
            } else {
                parsedData.push(line.split(delimiter));
            }
        }
    }
    
    return parsedData;
};

/**
 * Split a CSV line while properly handling quoted values
 */
MailMerge.CSV.splitCSVLine = function(line, delimiter) {
    delimiter = delimiter || ",";
    var values = line.split(delimiter);
    var lastIndex = values.length - 1;
    
    for (var i = lastIndex; i >= 0; i--) {
        var value = values[i];
        var trimmedValue = value.replace(/"\s+$/, "\"");
        
        if (trimmedValue.charAt(trimmedValue.length - 1) === "\"") {
            var startTrimmed = value.replace(/^\s+"/, "\"");
            
            if (startTrimmed.length > 1 && startTrimmed.charAt(0) === "\"") {
                values[i] = value
                    .replace(/^\s*"|"\s*$/g, "")
                    .replace(/""/g, "\"");
            } else {
                if (i > 0) {
                    values.splice(i - 1, 2, [values[i - 1], values[i]].join(delimiter));
                } else {
                    values = values.shift().split(delimiter).concat(values);
                }
            }
        } else {
            values[i] = value.replace(/""/g, "\"");
        }
    }
    
    return values;
};

/**
 * Validate CSV data structure
 */
MailMerge.CSV.validate = function(csvData) {
    var result = {
        isValid: true,
        errors: []
    };
    
    if (!csvData || csvData.length === 0) {
        result.isValid = false;
        result.errors.push("CSV file is empty");
        return result;
    }
    
    if (csvData.length < 2) {
        result.isValid = false;
        result.errors.push("CSV file must contain at least a header row and one data row");
        return result;
    }
    
    var headerCount = csvData[0].length;
    if (headerCount === 0) {
        result.isValid = false;
        result.errors.push("CSV header row is empty");
        return result;
    }
    
    for (var i = 0; i < headerCount; i++) {
        if (!csvData[0][i] || csvData[0][i].trim() === "") {
            result.errors.push("Column " + (i + 1) + " has an empty header");
        }
    }
    
    for (var row = 1; row < csvData.length; row++) {
        if (csvData[row].length !== headerCount) {
            result.errors.push(
                "Row " + (row + 1) + " has " + csvData[row].length + 
                " columns, expected " + headerCount
            );
        }
    }
    
    if (result.errors.length > 0) {
        result.isValid = false;
    }
    
    return result;
};

console.println("MailMerge.CSV loaded");
