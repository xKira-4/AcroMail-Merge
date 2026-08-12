// This script must be saved in the JavaScripts folder as a trusted function
// Location: C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Javascripts\
// or: %appdata%\Adobe\Acrobat\DC\JavaScripts\

importXMLToXFABatch = app.trustedFunction(function() {
    app.beginPriv();
    
    try {
        var xmlDirectory = "C:/Users/1513280110.MIL/OneDrive - US Army/Desktop/Temp/Auto 2062/Generated_XMLs/";
        var outputDirectory = "/c/Users/1513280110.MIL/OneDrive - US Army/Desktop/Temp/Auto 2062/Generated_XMLs/";
        
        // Prompt user to select XML files
        var selectedFiles = app.browseForDoc({
            bOpen: true,
            cFilenameFilter: ["XML Files (*.xml)|*.xml"],
            bAllowMultiple: true
        });
        
        if (selectedFiles == null) {
            app.alert("No files selected. Operation cancelled.");
            return;
        }
        
        // Handle single file selection (returns string) vs multiple (returns array)
        var fileArray = [];
        if (typeof selectedFiles === "string") {
            fileArray.push(selectedFiles);
        } else {
            fileArray = selectedFiles;
        }
        
        console.println("Processing " + fileArray.length + " XML file(s)");
        
        var successCount = 0;
        var errorCount = 0;
        var errorMessages = [];
        
        for (var i = 0; i < fileArray.length; i++) {
            try {
                var xmlFilePath = fileArray[i];
                console.println("Processing file " + (i + 1) + " of " + fileArray.length + ": " + xmlFilePath);
                
                // Extract filename without path and extension
                var fileName = xmlFilePath.substring(xmlFilePath.lastIndexOf("/") + 1);
                if (fileName.indexOf("\\") > -1) {
                    fileName = fileName.substring(fileName.lastIndexOf("\\") + 1);
                }
                
                // Remove "_data.xml" suffix to get base name
                var baseName = fileName.replace("_data.xml", "").replace(".xml", "");
                
                console.println("Base name: " + baseName);
                
                // Import the XML data into the form
                xfa.host.importData(xmlFilePath);
                
                // Save the form with the base name
                var newFileName = outputDirectory + baseName + ".pdf";
                this.saveAs(newFileName);
                
                console.println("Saved: " + newFileName);
                successCount++;
                
            } catch (fileError) {
                console.println("Error processing file " + (i + 1) + ": " + fileError);
                errorMessages.push("File " + (i + 1) + ": " + fileError);
                errorCount++;
            }
        }
        
        // Show completion message
        var message = "Processing complete!\n\n";
        message += "Successful: " + successCount + "\n";
        message += "Errors: " + errorCount;
        
        if (errorCount > 0) {
            message += "\n\nError details:\n" + errorMessages.join("\n");
        }
        
        app.alert(message);
        
    } catch (error) {
        console.println("Error in batch process: " + error);
        app.alert("Error: " + error);
    }
    
    app.endPriv();
});