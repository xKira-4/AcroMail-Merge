/**
 * ============================================================================
 * MAIL MERGE - INITIALIZATION
 * ============================================================================
 * This file MUST load first (hence the 00 prefix)
 * Creates the namespace and configuration
 */

// Create the main namespace
var MailMerge = {
    // Configuration settings
    config: {
        version: "1.0",
        demoMode: true,
        defaultDelimiter: ",",
        csvFileName: "file.csv",
        encoding: "utf-8"
    },
    
    // Placeholder objects for modules (will be populated by other files)
    Utilities: {},
    Trusted: {},
    CSV: {},
    PDF: {},
    Dialogs: {}
};

console.println("MailMerge namespace initialized");
