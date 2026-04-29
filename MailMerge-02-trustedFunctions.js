/**
 * ============================================================================
 * MAIL MERGE - TRUSTED FUNCTIONS MODULE
 * ============================================================================
 */

/**
 * Show an alert dialog
 */
MailMerge.Trusted.alert = app.trustedFunction(function(message, icon) {
    app.beginPriv();
    if (typeof message === "object") {
        app.alert(message);
    } else {
        app.alert(message, icon);
    }
    app.endPriv();
});

/**
 * Show a response dialog (text input)
 */
MailMerge.Trusted.response = app.trustedFunction(function(message, title) {
    app.beginPriv();
    var result = app.response(message, title);
    app.endPriv();
    return result;
});

/**
 * Execute a custom dialog
 */
MailMerge.Trusted.execDialog = app.trustedFunction(function(dialogObject) {
    app.beginPriv();
    var result = app.execDialog(dialogObject);
    app.endPriv();
    return result;
});

// New Document - Trust Propagator
var MailMerge_safeNewDoc = app.trustPropagatorFunction(function() {
    app.beginPriv();
    var doc = app.newDoc();
    app.endPriv();
    return doc;
});

/**
 * Create a new document
 */
MailMerge.Trusted.newDoc = app.trustedFunction(function() {
    app.beginPriv();
    var doc = MailMerge_safeNewDoc();
    app.endPriv();
    return doc;
});

// Import Data Object - Trust Propagator
var MailMerge_safeImportDataObject = app.trustPropagatorFunction(function(doc, objectName) {
    app.beginPriv();
    doc.importDataObject({
        cName: objectName
    });
    app.endPriv();
});

/**
 * Import a data object
 */
MailMerge.Trusted.importDataObject = app.trustedFunction(function(doc, objectName) {
    app.beginPriv();
    MailMerge_safeImportDataObject(doc, objectName);
    app.endPriv();
});

// Save As - Trust Propagator
var MailMerge_safeSaveAs = app.trustPropagatorFunction(function(doc, path) {
    app.beginPriv();
    doc.saveAs({cPath: path});
    app.endPriv();
});

/**
 * Save document
 */
MailMerge.Trusted.saveAs = app.trustedFunction(function(doc, path) {
    app.beginPriv();
    MailMerge_safeSaveAs(doc, path);
    app.endPriv();
});

/**
 * Insert pages from another document
 */
MailMerge.Trusted.insertPages = app.trustedFunction(function(doc, sourcePath) {
    app.beginPriv();
    doc.insertPages(doc.numPages - 1, sourcePath);
    app.endPriv();
});

// Button Import Icon - Trust Propagator
var MailMerge_safeButtonImportIcon = app.trustPropagatorFunction(function(field, iconPath) {
    app.beginPriv();
    var result = field.buttonImportIcon(iconPath);
    app.endPriv();
    return result;
});

/**
 * Import button icon
 */
MailMerge.Trusted.buttonImportIcon = app.trustedFunction(function(field, iconPath) {
    app.beginPriv();
    var result = MailMerge_safeButtonImportIcon(field, iconPath);
    app.endPriv();
    return result;
});

console.println("MailMerge.Trusted loaded");
