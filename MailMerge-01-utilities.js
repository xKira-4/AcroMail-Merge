/**
 * ============================================================================
 * MAIL MERGE - UTILITIES MODULE
 * ============================================================================
 */

/**
 * Normalize line endings to \n
 */
MailMerge.Utilities.normalizeLineEndings = function(text) {
    return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
};

/**
 * Clean CSV value by removing quotes and unescaping
 */
MailMerge.Utilities.cleanCSVValue = function(value) {
    return value
        .replace(/^"/, "")      // Remove leading quote
        .replace(/"$/, "")      // Remove trailing quote
        .replace(/""/g, "\"");  // Unescape double quotes
};

/**
 * Remove illegal characters from file path/name
 */
MailMerge.Utilities.makeFilePathSafe = function(text) {
    return text.replace(/[\\\/\:\*\?\"\<\>\|\,\n\r]/g, "");
};

/**
 * Replace tagged placeholders in a string with actual values
 */
MailMerge.Utilities.replaceTag = function(text, tagName, replacement) {
    var regex = new RegExp("<" + tagName + ">", "g");
    return text.replace(regex, replacement);
};

/**
 * Convert Windows/Mac file paths to PDF-compatible paths
 */
MailMerge.Utilities.convertPathToPDFFormat = function(path) {
    if (app.platform === "WIN") {
        if (path.charAt(0) === "\\") {
            // UNC path (\\server\share)
            return path.replace(/^\\/, "//").replace(/\\/g, "/");
        } else {
            // Regular path (C:\folder\file)
            return "/" + path.replace(":\\", "/").replace(/\\/g, "/");
        }
    }
    
    if (app.platform === "MAC") {
        path = path.replace(/:/g, "/");
        if (!/^\//.test(path)) {
            path = "/" + path;
        }
        return path;
    }
    
    return path;
};

/**
 * Check if a string is empty or whitespace
 */
MailMerge.Utilities.isEmpty = function(str) {
    return !str || str.trim() === "";
};

/**
 * Log a message to the console with timestamp
 */
MailMerge.Utilities.log = function(message, level) {
    level = level || "INFO";
    var timestamp = new Date().toISOString();
    console.println("[" + timestamp + "] [" + level + "] " + message);
};

console.println("MailMerge.Utilities loaded");
