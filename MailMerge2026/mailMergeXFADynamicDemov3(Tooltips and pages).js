//use windows environment variables %variable%
//temp files created in %LOCALAPPDATA%\Temp\acrobat_sbx
//save this to %APPDATA%\Adobe\Acrobat\Privileged\DC\Javascripts
if (app["viewerVersion"] < 10) {
    app["addMenuItem"]({
      cName: "MailMergeXFA",
      cUser: "Mail Merge XFA",
      cParent: "Tools",
      cExec: "mailMergeXFAFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  } else {
    app["addToolButton"]({
      cName: "MailMergeXFA",
      cLabel: "Mail Merge XFA",
      cTooltext: "Mail Merge XFA",
      cExec: "mailMergeXFAFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  };

  var mailMergeXFADialog = {
    radioValue: null,
    initialize: function (dialog) {
      dialog.load({"rd01": true});
    },
    commit: function (dialog) {
      var oRslt = dialog.store();
      this.radioValue = this.getRadioValue(oRslt);
    },
    getRadioValue: function (results) {
      for (var i=1; i<=3; i++) {
        if (results["rd0"+i]) {
          switch (i) {
            case 1: return ",";
            case 2: return "\t";
          }
        }
      }
      return null;
    },
    description: {
      name: "Mail Merge XFA",
      elements: [{
          type: "view",
          elements: [{
              item_id: "str1",
              name: "Mail Merge XFA",
              type: "static_text",
              font: "dialog",
              bold: true,
              alignment: "align_center",
              char_width: 30,
              height: 20
            },
            {type: "gap"},
            {
              type: "view",
              align_children: "align_row",
              elements: [
                {item_id: "str2", name: "Delimiter:", type: "static_text"},
                {type: "radio", item_id: "rd01", group_id: "group1", name: "Comma"},
                {type: "radio", item_id: "rd02", group_id: "group1", name: "Tab"}
              ]
            }
          ]
        },
        {type: "ok_cancel", alignment: "align_right"}
      ]
    }
  };

  function addPageInstance(doc) {
    try {
      var page2Node = doc.xfa.form.resolveNode('form1._Page2_Flowed');
      if (page2Node) {
        page2Node.addInstance(1);
        if (doc.xfa.host.version < 8) {
          doc.xfa.form.recalculate(1);
        }
        return true;
      }
    } catch (e) {}
    return false;
  }

  function removeExcessPages(doc, targetCount) {
    try {
      var page2Node = doc.xfa.form.resolveNode('form1._Page2_Flowed');
      if (!page2Node) return false;
      
      var currentCount = page2Node.count;
      if (targetCount < 1) targetCount = 1;
      
      while (currentCount > targetCount) {
        page2Node.removeInstance(currentCount - 1);
        currentCount--;
        if (doc.xfa.host.version < 8) {
          doc.xfa.form.recalculate(1);
        }
      }
      return true;
    } catch (e) {}
    return false;
  }

  function extractRowNumberFromTooltip(tooltipObj) {
    if (!tooltipObj) return null;
    try {
      var tooltip = (typeof tooltipObj === "string") ? tooltipObj : (tooltipObj.value || null);
      if (!tooltip || tooltip === "[object XFAObject]") return null;
      var match = tooltip.match(/row\s+(\d+)/i);
      return match ? parseInt(match[1], 10) : null;
    } catch (e) {}
    return null;
  }

  function scanNodeForFields(node, sectionId, fieldMap, sectionInfo) {
    try {
      if (node.className === "field") {
        var fieldName = node.name;
        var rowNumber = null;
        
        try {
          if (node.assist && node.assist.toolTip) {
            var tooltip = node.assist.toolTip;
            rowNumber = extractRowNumberFromTooltip(
              (typeof tooltip === "string") ? tooltip : (tooltip.value || null)
            );
          }
        } catch (e) {}
        
        if (rowNumber != null) {
          if (!sectionInfo[sectionId]) {
            sectionInfo[sectionId] = {maxRow: 0};
          }
          if (rowNumber > sectionInfo[sectionId].maxRow) {
            sectionInfo[sectionId].maxRow = rowNumber;
          }
        }
        
        if (!fieldMap[fieldName]) {
          fieldMap[fieldName] = [];
        }
        
        fieldMap[fieldName].push({
          field: node,
          rowNumber: rowNumber,
          section: sectionId
        });
      }
      
      if (node.nodes) {
        var nodeCount = node.nodes.length;
        for (var i = 0; i < nodeCount; i++) {
          scanNodeForFields(node.nodes.item(i), sectionId, fieldMap, sectionInfo);
        }
      }
    } catch (e) {}
  }

  function getXFAFieldMapDirect(doc) {
    var fieldMap = {};
    var sectionInfo = {};
    
    try {
      var xfaForm = doc.xfa.form;
      
      try {
        var page1Node = xfaForm.resolveNode('form1.Page1');
        if (page1Node) {
          scanNodeForFields(page1Node, "Page1", fieldMap, sectionInfo);
        }
      } catch (e) {}
      
      var page2Node = xfaForm.resolveNode('form1._Page2_Flowed');
      if (page2Node) {
        var instanceCount = page2Node.count;
        for (var i = 0; i < instanceCount; i++) {
          try {
            var instanceNode = xfaForm.resolveNode('form1.Page2_Flowed[' + i + ']');
            if (instanceNode) {
              scanNodeForFields(instanceNode, i, fieldMap, sectionInfo);
            }
          } catch (e) {}
        }
      }
    } catch (e) {}
    
    return {fieldMap: fieldMap, sectionInfo: sectionInfo};
  }

  function calculateSectionAndRow(itemIndex, sectionInfo) {
    var sections = [];
    for (var s in sectionInfo) {
      sections.push({id: s, maxRow: sectionInfo[s].maxRow});
    }
    
    sections.sort(function(a, b) {
      if (a.id === "Page1") return -1;
      if (b.id === "Page1") return 1;
      return a.id - b.id;
    });
    
    var cumulative = 0;
    for (var i = 0; i < sections.length; i++) {
      var capacity = sections[i].maxRow;
      if (itemIndex < cumulative + capacity) {
        return {
          section: sections[i].id,
          row: (itemIndex - cumulative) + 1
        };
      }
      cumulative += capacity;
    }
    return null;
  }

  function calculateInstancesNeeded(totalItems, sectionInfo) {
    var page1Capacity = sectionInfo["Page1"] ? sectionInfo["Page1"].maxRow : 0;
    var page2Capacity = sectionInfo[0] ? sectionInfo[0].maxRow : 0;
    
    if (page2Capacity == 0) return 1;
    
    var itemsAfterPage1 = totalItems - page1Capacity;
    return (itemsAfterPage1 <= 0) ? 1 : Math.ceil(itemsAfterPage1 / page2Capacity);
  }

  function groupCSVByRecipient(fileDataArray, toColumnIndex) {
    var groups = [];
    var currentRecipient = null;
    var currentGroup = null;
    
    for (var i = 1; i < fileDataArray.length; i++) {
      var row = fileDataArray[i];
      var toValue = row[toColumnIndex] ? 
        row[toColumnIndex].replace(/^"/, "").replace(/"$/, "").replace(/""/, "\"") : "";
      
      if (toValue !== "") {
        if (currentGroup != null) {
          groups.push(currentGroup);
        }
        currentRecipient = toValue;
        currentGroup = {recipient: currentRecipient, rows: [row]};
      } else if (currentGroup != null) {
        currentGroup.rows.push(row);
      }
    }
    
    if (currentGroup != null) {
      groups.push(currentGroup);
    }
    
    return groups;
  }

  function mailMergeXFAFiles(doc) {
    if (doc["xfa"] == null) {
      xfaAlert("Error! This script is designed for XFA forms only.");
      return;
    }
    
    if ("ok" != xfaDialog(mailMergeXFADialog)) return;
    
    var delimiterValue = mailMergeXFADialog.radioValue;
    if (!delimiterValue) return;
    
    var fileDataArray = readXFACSV(delimiterValue, null);
    if (!fileDataArray || fileDataArray["length"] == 0) {
      xfaAlert("No items found in the input file.");
      return;
    }
    
    var csvHeaders = fileDataArray[0];
    
    var filenameColumnIndex = -1;
    for (var i = 0; i < csvHeaders["length"]; i++) {
      if (csvHeaders[i] == "To") {
        filenameColumnIndex = i;
        break;
      }
    }
    
    if (filenameColumnIndex == -1) {
      xfaAlert("Error! Could not find 'To' column in CSV file.");
      return;
    }
    
    var recipientGroups = groupCSVByRecipient(fileDataArray, filenameColumnIndex);
    if (recipientGroups.length == 0) {
      xfaAlert("Error! No valid data groups found in CSV.");
      return;
    }
    
    var filesCreated = 0;
    var hasErrors = false;
    var docPath = doc["path"]["replace"](doc["documentFileName"], "");
    
    var csvFieldNames = {};
    for (var h = 0; h < csvHeaders.length; h++) {
      csvFieldNames[csvHeaders[h]] = true;
    }
    
    for (var groupIndex = 0; groupIndex < recipientGroups.length; groupIndex++) {
      var group = recipientGroups[groupIndex];
      var recipient = group.recipient;
      var rows = group.rows;
      
      var mapResult = getXFAFieldMapDirect(doc);
      var xfaFieldMap = mapResult.fieldMap;
      var sectionInfo = mapResult.sectionInfo;
      
      var instancesNeeded = calculateInstancesNeeded(rows.length, sectionInfo);
      var page2Node = doc.xfa.form.resolveNode('form1._Page2_Flowed');
      var currentInstances = page2Node ? page2Node.count : 1;
      
      if (instancesNeeded > currentInstances) {
        var pagesToAdd = instancesNeeded - currentInstances;
        for (var p = 0; p < pagesToAdd; p++) {
          if (!addPageInstance(doc)) {
            hasErrors = true;
            break;
          }
        }
        mapResult = getXFAFieldMapDirect(doc);
        xfaFieldMap = mapResult.fieldMap;
        sectionInfo = mapResult.sectionInfo;
      } else if (instancesNeeded < currentInstances) {
        removeExcessPages(doc, instancesNeeded);
        mapResult = getXFAFieldMapDirect(doc);
        xfaFieldMap = mapResult.fieldMap;
        sectionInfo = mapResult.sectionInfo;
      }
      
      for (var fieldName in xfaFieldMap) {
        if (csvFieldNames[fieldName] && fieldName != "To") {
          var fieldInstances = xfaFieldMap[fieldName];
          for (var fi = 0; fi < fieldInstances.length; fi++) {
            try {
              fieldInstances[fi].field.rawValue = null;
            } catch (e) {}
          }
        }
      }
      
      if (xfaFieldMap["To"]) {
        var toFields = xfaFieldMap["To"];
        for (var ti = 0; ti < toFields.length; ti++) {
          try {
            toFields[ti].field.rawValue = recipient;
          } catch (e) {
            hasErrors = true;
          }
        }
      }
      
      for (var rowIdx = 0; rowIdx < rows.length; rowIdx++) {
        var rowData = rows[rowIdx];
        var location = calculateSectionAndRow(rowIdx, sectionInfo);
        
        if (!location) continue;
        
        var targetSection = location.section;
        var targetRow = location.row;
        
        for (var colIndex = 0; colIndex < csvHeaders["length"]; colIndex++) {
          if (colIndex >= rowData["length"] || colIndex == filenameColumnIndex) continue;
          
          var headerName = csvHeaders[colIndex];
          var cellValue = rowData[colIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
          
          if (xfaFieldMap[headerName]) {
            var matchingFields = xfaFieldMap[headerName];
            var targetField = null;
            
            for (var f = 0; f < matchingFields.length; f++) {
              var fieldInfo = matchingFields[f];
              if (fieldInfo.section == targetSection && fieldInfo.rowNumber == targetRow) {
                targetField = fieldInfo;
                break;
              }
            }
            
            if (!targetField && matchingFields.length > 0) {
              targetField = matchingFields[0];
            }
            
            if (targetField) {
              try {
                targetField.field.rawValue = cellValue;
              } catch (e) {
                hasErrors = true;
              }
            }
          }
        }
      }
      
      var fileName = recipient["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
      if (!fileName) {
        fileName = "output_" + (groupIndex + 1);
      }
      
      if (!fileName.match(/\.pdf$/i)) {
        fileName = fileName + ".pdf";
      }
      
      var savePath = docPath + fileName;
      
      try {
        xfaTrustedSaveAs(doc, savePath);
        filesCreated++;
      } catch (e) {
        hasErrors = true;
      }
    }
    
    if (hasErrors) {
      xfaAlert("Done. " + filesCreated + " files created with some errors.");
    } else {
      xfaAlert({cMsg:"Done. " + filesCreated + " merged files were created.", nIcon:4});
    }
  }

  function readXFACSV(delimiter, unused) {
    if (!delimiter) delimiter = ";";
    
    var tempDoc = xfaTrustedNewDoc();
    var csvFileName = "file.csv";
    xfaTrustedImportDataObject(tempDoc, csvFileName);
    
    if (!tempDoc["dataObjects"] || tempDoc["dataObjects"]["length"] != 1) {
      return null;
    }
    
    var dataStream = tempDoc["getDataObjectContents"](csvFileName);
    var csvContent = util["stringFromStream"]({
      oStream: dataStream,
      oCharSet: "utf-8"
    });
    
    csvContent = csvContent["replace"](/\r\n/g, "\n")["replace"](/\r/g, "\n");
    tempDoc["closeDoc"](true);
    
    var resultArray = [];
    var fileDataArray = csvContent["split"]("\n");
    
    for (var i = 0; i < fileDataArray["length"]; i++) {
      var cleanLine = fileDataArray[i]["replace"](/[\r\n]/g, "")["replace"]("﻿", "");
      if (cleanLine) {
        resultArray["push"](
          (delimiter == ",") ? splitXFACSV(cleanLine, delimiter) : cleanLine["split"](delimiter)
        );
      }
    }
    
    return resultArray;
  }

  function splitXFACSV(str, delim) {
    var arr = str["split"](delim || ",");
    var arrLen = arr["length"] - 1;
    for (var temp; arrLen >= 0; arrLen--) {
      if (arr[arrLen]["replace"](/"\s+$/, "\"")["charAt"](arr[arrLen]["length"] - 1) == "\"") {
        if ((temp = arr[arrLen]["replace"](/^\s+"/, "\""))["length"] > 1 && temp["charAt"](0) == "\"") {
          arr[arrLen] = arr[arrLen]["replace"](/^\s*"|"\s*$/g, "")["replace"](/""/g, "\"");
        } else {
          if (arrLen) {
            arr["splice"](arrLen - 1, 2, [arr[arrLen - 1], arr[arrLen]]["join"](delim));
          } else {
            arr = arr["shift"]()["split"](delim)["concat"](arr);
          }
        }
      } else {
        arr[arrLen]["replace"](/""/g, "\"");
      }
    }
    return arr;
  }

  xfaImportDataObject = app["trustPropagatorFunction"](function (doc, objName) {
    app["beginPriv"]();
    doc["importDataObject"]({cName: objName});
    app["endPriv"]();
  });

  xfaTrustedImportDataObject = app["trustedFunction"](function (doc, objName) {
    app["beginPriv"]();
    xfaImportDataObject(doc, objName);
    app["endPriv"]();
  });

  xfaNewDoc = app["trustPropagatorFunction"](function () {
    app["beginPriv"]();
    return app["newDoc"]();
    app["endPriv"]();
  });

  xfaTrustedNewDoc = app["trustedFunction"](function () {
    app["beginPriv"]();
    return xfaNewDoc();
    app["endPriv"]();
  });

  xfaSaveAs = app["trustPropagatorFunction"](function(doc, path){
    app["beginPriv"]();
    doc["saveAs"]({cPath: path});
    app["endPriv"]();
  });

  xfaTrustedSaveAs = app["trustedFunction"](function (doc, path) {
    app["beginPriv"]();
    xfaSaveAs(doc, path);
    app["endPriv"]();
  });

  xfaAlert = app.trustedFunction(function(altMessage) {
    app.beginPriv();
    app.alert(altMessage);
    app.endPriv();
  });

  xfaResponse = app.trustedFunction(function(responseMessage) {
    app.beginPriv();
    app.response(responseMessage);
    app.endPriv();
  });

  var xfaDialog = app.trustedFunction(function(dialogObject) {
    app.beginPriv();
    return app.execDialog(dialogObject);
    app.endPriv();
  });