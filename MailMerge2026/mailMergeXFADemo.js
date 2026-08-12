//use windows environment variables %variable%
//temp files created in %LOCALAPPDATA%\Temp\acrobat_sbx
//save this to %APPDATA%\Adobe\Acrobat\Privileged\DC\Javascripts
//this creates the menu button in the add-ons tool menu
if (app["viewerVersion"] < 10) {
    app["addMenuItem"]({
      cName: "MailMergeDEMO",
      cUser: "Mail Merge XFA (DEMO)",
      cParent: "Tools",
      cExec: "mailMergeFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  } else {
    app["addToolButton"]({
      cName: "MailMergeDEMO",
      cLabel: "Mail Merge XFA (DEMO)",
      cTooltext: "Mail Merge XFA (DEMO)",
      cExec: "mailMergeFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  };

  //setup mailMergeDialog
  var mailMergeDialog = {
    radioValue: null,
    initialize: function (dialog) {
      dialog.load({
        "rd01": true,
      });
    },
    commit: function (dialog) {
      var oRslt = dialog.store();
      this.radioValue = this.getRadioValue(oRslt);
    },
    getRadioValue: function (results) {
      for ( var i=1; i<=3; i++) {
        if ( results["rd0"+i] ) {
        switch (i) {
        case 1:
        var cDelim = ",";
        break;
        case 2:
        var cDelim = "\t";
        break;
        }
        }
      };
      return cDelim;
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
              height: 20,
            },
            {
              type: "gap",
            },
            {
              type: "view",
              align_children: "align_row",
              elements: [{
                  item_id: "str2",
                  name: "Delimiter:",
                  type: "static_text",
                },
                {
                    type: "radio",
                    item_id: "rd01",
                    group_id: "group1",
                    name: "Comma"
                  },
                  {
                    type: "radio",
                    item_id: "rd02",
                    group_id: "group1",
                    name: "Tab"
                  },
              ],
            },
          ],
        },
        {
          type: "ok_cancel",
          alignment: "align_right",
        },
      ]
    },
  };

function mailMergeFiles(_0x8ec0x2) {
    if (_0x8ec0x2["xfa"] == null) {
      myAlert("Error! This script is designed for XFA forms only.");
      return;
    };
    
    // Get delimiter from user
    if ("ok" == myDialog(mailMergeDialog)) {
      var delimiterValue = mailMergeDialog.radioValue;
    };
    if (delimiterValue == null || delimiterValue == "") {
      return;
    };
    
    var fileDataArray = readCSV(delimiterValue, null);
    if (fileDataArray == null || fileDataArray["length"] == 0) {
      myAlert("No items found in the input file.");
      return;
    };
    
    // Find column indices for all required fields
    var toColumnIndex = -1;
    var itemDescColumnIndex = -1;
    var endItemMatNumColumnIndex = -1;
    
    for (var i = 0; i < fileDataArray[0]["length"]; i++) {
      if (fileDataArray[0][i] == "To") {
        toColumnIndex = i;
      }
      if (fileDataArray[0][i] == "Item_Description_Column_C") {
        itemDescColumnIndex = i;
      }
      if (fileDataArray[0][i] == "End_Item_Material_Number") {
        endItemMatNumColumnIndex = i;
      }
    }
    
    if (toColumnIndex == -1) {
      myAlert("Error! Could not find 'To' column in CSV file.");
      return;
    }
    
    if (itemDescColumnIndex == -1) {
      myAlert("Error! Could not find 'Item_Description_Column_C' column in CSV file.");
      return;
    }
    
    if (endItemMatNumColumnIndex == -1) {
      myAlert("Error! Could not find 'End_Item_Material_Number' column in CSV file.");
      return;
    }
    
    var _0x8ec0x7 = 0;
    var _0x8ec0x8 = false;
    var _0x8ec0xf = _0x8ec0x2["path"]["replace"](_0x8ec0x2["documentFileName"], "");
    
    // Loop through each row in the CSV (skip header row)
    for (var _0x8ec0x9 = 1; _0x8ec0x9 < fileDataArray["length"]; _0x8ec0x9++) {
      
      // Get the values for this row
      var toValue = "";
      var itemDescValue = "";
      var endItemMatNumValue = "";
      
      if (toColumnIndex < fileDataArray[_0x8ec0x9]["length"]) {
        toValue = fileDataArray[_0x8ec0x9][toColumnIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
      }
      
      if (itemDescColumnIndex < fileDataArray[_0x8ec0x9]["length"]) {
        itemDescValue = fileDataArray[_0x8ec0x9][itemDescColumnIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
      }
      
      if (endItemMatNumColumnIndex < fileDataArray[_0x8ec0x9]["length"]) {
        endItemMatNumValue = fileDataArray[_0x8ec0x9][endItemMatNumColumnIndex]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
      }
      
      // Set XFA field values using the document's xfa object
      try {
        var toField = _0x8ec0x2.xfa.resolveNode("xfa.form.form1.Page1.To[0]");
        if (toField != null) {
          toField.rawValue = toValue;
        } else {
          console["println"]("Error! Can't find To field");
          _0x8ec0x8 = true;
        }
      } catch (e) {
        console["println"]("Error! Can't set value for To field: " + e);
        _0x8ec0x8 = true;
      }
      
      try {
        var itemDescField = _0x8ec0x2.xfa.resolveNode("xfa.form.form1.Page1.Item_Description_Column_C[1]");
        if (itemDescField != null) {
          itemDescField.rawValue = itemDescValue;
        } else {
          console["println"]("Error! Can't find Item_Description_Column_C field");
          _0x8ec0x8 = true;
        }
      } catch (e) {
        console["println"]("Error! Can't set value for Item_Description_Column_C field: " + e);
        _0x8ec0x8 = true;
      }
      
      try {
        var endItemMatNumField = _0x8ec0x2.xfa.resolveNode("xfa.form.form1.Page1.End_Item_Material_Number[0]");
        if (endItemMatNumField != null) {
          endItemMatNumField.rawValue = endItemMatNumValue;
        } else {
          console["println"]("Error! Can't find End_Item_Material_Number field");
          _0x8ec0x8 = true;
        }
      } catch (e) {
        console["println"]("Error! Can't set value for End_Item_Material_Number field: " + e);
        _0x8ec0x8 = true;
      }
      
      // Create safe filename from "To" value
      var fileName = toValue["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
      if (fileName == "") {
        fileName = "output_" + _0x8ec0x9;
      }
      
      // Add .pdf extension if not present
      if (!fileName.match(/\.pdf$/i)) {
        fileName = fileName + ".pdf";
      }
      
      var savePath = _0x8ec0xf + fileName;
      
      // Save the document
      try {
        myTrustedSaveAs(_0x8ec0x2, savePath);
        _0x8ec0x7++;
      } catch (e) {
        console["println"]("Error! Can't save file: " + savePath + " - " + e);
        _0x8ec0x8 = true;
      }
    }
    
    if (_0x8ec0x8) {
      myAlert("Done. " + _0x8ec0x7 + " merged files were created, but there were some errors. See the Console for more details.", 1);
      console["show"]();
    } else {
      myAlert({cMsg:"Done. " + _0x8ec0x7 + " merged files were created.", nIcon:4});
    }
  }

  function readCSV(_0x8ec0x18, _0x8ec0x19) {
    if (_0x8ec0x18 == null) {
      _0x8ec0x18 = ";";
    };
    var _0x8ec0x1a = "";
    var _0x8ec0x1b = myTrustedNewDoc();
    var _0x8ec0x1c = "file.csv";
    myTrustedImportDataObject(_0x8ec0x1b, _0x8ec0x1c);
    if (_0x8ec0x1b["dataObjects"] == null || _0x8ec0x1b["dataObjects"]["length"] != 1) {
      return;
    };
    var _0x8ec0x1d = _0x8ec0x1b["getDataObjectContents"](_0x8ec0x1c);
    var _0x8ec0x1a = util["stringFromStream"]({
      oStream: _0x8ec0x1d,
      oCharSet: "utf-8"
    });
    _0x8ec0x1a = _0x8ec0x1a["replace"](/\r\n/g, "\n");
    _0x8ec0x1a = _0x8ec0x1a["replace"](/\r/g, "\n");
    _0x8ec0x1b["closeDoc"](true);
    var _0x8ec0x1e = new Array();
    var fileDataArray = _0x8ec0x1a["split"]("\n");
    for (var _0x8ec0x9 = 0; _0x8ec0x9 < fileDataArray["length"]; _0x8ec0x9++) {
      var _0x8ec0x1f = fileDataArray[_0x8ec0x9]["replace"](/[\r\n]/g, "")["replace"]("﻿", "");
      if (_0x8ec0x1f != "") {
        if (_0x8ec0x18 == ",") {
          _0x8ec0x1e["push"](splitCSV(_0x8ec0x1f, _0x8ec0x18));
        } else {
          _0x8ec0x1e["push"](_0x8ec0x1f["split"](_0x8ec0x18));
        }
      }
    };
    return _0x8ec0x1e;
  }

  function splitCSV(_0x8ec0x12, _0x8ec0x21) {
    var _0x8ec0x22 = _0x8ec0x12["split"](_0x8ec0x21 = _0x8ec0x21 || ",");
    var _0x8ec0x23 = _0x8ec0x22["length"] - 1;
    for (var _0x8ec0x24; _0x8ec0x23 >= 0; _0x8ec0x23--) {
      if (_0x8ec0x22[_0x8ec0x23]["replace"](/"\s+$/, "\"")["charAt"](_0x8ec0x22[_0x8ec0x23]["length"] - 1) == "\"") {
        if ((_0x8ec0x24 = _0x8ec0x22[_0x8ec0x23]["replace"](/^\s+"/, "\""))["length"] > 1 && _0x8ec0x24["charAt"](0) == "\"") {
          _0x8ec0x22[_0x8ec0x23] = _0x8ec0x22[_0x8ec0x23]["replace"](/^\s*"|"\s*$/g, "")["replace"](/""/g, "\"");
        } else {
          if (_0x8ec0x23) {
            _0x8ec0x22["splice"](_0x8ec0x23 - 1, 2, [_0x8ec0x22[_0x8ec0x23 - 1], _0x8ec0x22[_0x8ec0x23]]["join"](_0x8ec0x21));
          } else {
            _0x8ec0x22 = _0x8ec0x22["shift"]()["split"](_0x8ec0x21)["concat"](_0x8ec0x22);
          }
        }
      } else {
        _0x8ec0x22[_0x8ec0x23]["replace"](/""/g, "\"");
      }
    };
    return _0x8ec0x22;
  }

  myImportDataObject = app["trustPropagatorFunction"](function (_0x8ec0x2, _0x8ec0x27) {
    app["beginPriv"]();
    _0x8ec0x2["importDataObject"]({
      cName: _0x8ec0x27
    });
    app["endPriv"]();
  });

  myTrustedImportDataObject = app["trustedFunction"](function (_0x8ec0x2, _0x8ec0x27) {
    app["beginPriv"]();
    myImportDataObject(_0x8ec0x2, _0x8ec0x27);
    app["endPriv"]();
  });

  myNewDoc = app["trustPropagatorFunction"](function () {
    app["beginPriv"]();
    return app["newDoc"]();
    app["endPriv"]();
  });

  myTrustedNewDoc = app["trustedFunction"](function () {
    app["beginPriv"]();
    return myNewDoc();
    app["endPriv"]();
  });

  safeSaveAs = app["trustPropagatorFunction"](function(_0x8ec0x2,_0x8ec0x28){
    app["beginPriv"]();
    _0x8ec0x2["saveAs"]({cPath:_0x8ec0x28});
    app["endPriv"]()});

  myTrustedSaveAs = app["trustedFunction"](function (_0x8ec0x2, _0x8ec0x28) {
    app["beginPriv"]();
    safeSaveAs(_0x8ec0x2, _0x8ec0x28);
    app["endPriv"]();
  });

  myAlert = app.trustedFunction(function(altMessage)
  {
      app.beginPriv();
      app.alert(altMessage);
      app.endPriv();
  });

  myResponse = app.trustedFunction(function(responseMessage)
  {
      app.beginPriv();
      app.response(responseMessage);
      app.endPriv();
  });

  var myDialog = app.trustedFunction(function(dialogObject)
  {
     app.beginPriv();
     return app.execDialog(dialogObject);
     app.endPriv();
  });