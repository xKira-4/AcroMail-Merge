//use windows environment variables %variable%
//temp files created in %LOCALAPPDATA%\Temp\acrobat_sbx
//save this to %APPDATA%\Adobe\Acrobat\Privileged\DC\Javascripts
//this creates the menu button in the add-ons tool menu
if (app["viewerVersion"] < 10) {
    app["addMenuItem"]({
      cName: "MailMergeDEMO",
      cUser: "Mail Merge (DEMO)",
      cParent: "Tools",
      cExec: "mailMergeFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  } else {
    app["addToolButton"]({
      cName: "MailMergeDEMO",
      cLabel: "Mail Merge (DEMO)",
      cTooltext: "Mail Merge (DEMO)",
      cExec: "mailMergeFiles(this)",
      cEnable: "event.rc = (event.target != null);"
    });
  };
  var dynamicDlg = { 
    fileHeaders: [],
    initialize: function(dialog) {
      this.loadDefaults(dialog);
    },
    commit: function(dialog) {
        var headerResults = dialog.store();
        this.fileHeaders = this.getFileHeaders(headerResults);
    },
    loadDefaults: function(dialog) {
      // Count the number of checkboxes in the dynamicDlg description
      var checkboxCount = 0;
      dynamicDlg.description.elements[0].elements.forEach(element => {
          if (element.type === "check_box") {
              checkboxCount++;
          }
      });
       // Create an object to hold the default values for the checkboxes
      var defaults = {};
      for (var i = 0; i < checkboxCount; i++) {
          defaults["chk" + i] = true;
      }
      // Load the defaults into the dialog
      dialog.load(defaults);
  },
    getFileHeaders: function(results) {
        var headers = [];
        for ( var x in results) {
          if (x.startsWith("chk")) {
          }
          if (x.startsWith("chk") && results[x]) {
            console.println(x);
            headers.push(x);
          }
        }
        return headers;
    },
    description: {
        name: "Dynamic Dialog", 
        elements: [ {
            type: "view", 
            elements: [
                { name: "Select which headings to include:", 
                    type: "static_text", 
                },
            ],
        },
        { 
            type: "ok_cancel",
            alignment: "align_right",
        },
    ]
    }
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
      this.radioValue = this.getRadioValue(oRslt); // Get the radio button value. Return it with mailMergeDialog.radioValue
      //console.println("You selected radio button " + radioValue);
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
      name: "Mail Merge",
      elements: [{
          type: "view",
          elements: [{
              item_id: "str1",
              name: "Mail Merge",
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
  // _0x8ec0x2 is equal to either app or this oth work when plugged into the javascript
  function mailMergeFiles(_0x8ec0x2) {
    if (_0x8ec0x2["xfa"] != null) {
      myAlert("Error! This file was created using LiveCycle Designer and therefore can't be mail-merged using this script.");
      return;
    }
    ;
    // the default delimiter is a tab
    // the user can change the delimiter to a comma or a semicolon
    // comma delimited csv files worked fine during testing
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
    // for loop to add the headers to the dynamic dialog
    var headerObj = {};
    for (var i = 0; i < fileDataArray[0].length; i++) {
      var obj = { type: "check_box", item_id: "chk" + i, name: fileDataArray[0][i] };
      headerObj["chk" + i] = fileDataArray[0][i];
      dynamicDlg.description.elements[0].elements.push(obj);
    };
    console.println(headerObj);
    // call the dynamic dialog
    if ("ok" == myDialog(dynamicDlg)) {
      // get the values of the checkboxes
      console.println(dynamicDlg.fileHeaders);
    };
    // for loop to remove the check boxes from the dynamic dialog
    for (var i = 0; i < fileDataArray[0].length; i++) {
      dynamicDlg.description.elements[0].elements.pop();
    };
    // var fileNamePattern = myResponse("Enter the file-name pattern for the mail-merged files.\nYou can use a field's name in a tag, for example:\n\t\"Letter to <FirstName> <LastName>\"\n(leave empty to use the row number: 1.pdf, 2.pdf, etc.)", "Mail Merge");
    // if (fileNamePattern == null) {
    //   fileNamePattern = "";
    // };
    var _0x8ec0x7 = 0;
    var _0x8ec0x8 = false;
    var ogDoc = this.documentFileName;
    var ogPageNum = this.numPages;
    var re = /\.pdf$/i;
    var copyDoc = this.documentFileName.replace(re,"-copy.pdf");
    var _0x8ec0xf = _0x8ec0x2["path"]["replace"](_0x8ec0x2["documentFileName"], "");
    var _0x8ec0x10 = _0x8ec0xf + copyDoc;
    myTrustedSaveAs(_0x8ec0x2, _0x8ec0x10);
    //this is the loop that goes through the csv file and sets the values of the fields
    //it also saves each row in the csv as its own pdf
    for (var _0x8ec0x9 = 1; _0x8ec0x9 < fileDataArray["length"]; _0x8ec0x9++) {
      // var fileName = fileNamePattern;
      _0x8ec0x2["resetForm"]();
      for (var _0x8ec0xb = 0; _0x8ec0xb < fileDataArray[0]["length"]; _0x8ec0xb++) {
        if (_0x8ec0xb >= fileDataArray[_0x8ec0x9]["length"]) {
          break;
        }
        ;
        var columnHeaders = fileDataArray[0][_0x8ec0xb];
        var _0x8ec0xd = _0x8ec0x2["getField"](columnHeaders);
        if (_0x8ec0xd == null) {
          console["println"]("Error! Can't locate a field called: " + columnHeaders);
          _0x8ec0x8 = true;
          continue;
        }
        ;
        var rowValues = fileDataArray[_0x8ec0x9][_0x8ec0xb]["replace"](/^"/, "")["replace"](/"$/, "")["replace"](/""/, "\"");
        if (_0x8ec0xd["type"] == "button") {
          if (_0x8ec0xd["buttonPosition"] == position["textOnly"]) {
            _0x8ec0xd["buttonSetCaption"](rowValues);
          } else {
            try {
              myTrustedButtonImportIcon(_0x8ec0xd, convertRealPathToPDFPath(rowValues));
            } catch (e) {
              console["println"]("Error! Can't apply the image \"" + rowValues + "\" as the icon of the field: " + columnHeaders);
              _0x8ec0x8 = true;
              continue;
            }
          }
        } else {
          try {
            _0x8ec0xd["value"] = rowValues; //this is where the value is being set to the fields
          } catch (e) {
            console["println"]("Error! Can't apply the value \"" + rowValues + "\" to the field: " + columnHeaders);
            _0x8ec0x8 = true;
            continue;
          }
        }
        ;
        // if (fileNamePattern != "") {
        //   fileName = replaceTagWithText(fileName, columnHeaders, rowValues);
        // }
      };
      // if (fileNamePattern == "") {
      //   fileName = _0x8ec0x9; //this is where the file name is set to the number of documents created so far i.e. 1..2..3..4
      // } else {
      //   fileName = fileName["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, ""); //this removes all illegal characters from the file name
      // };
      _0x8ec0x2.flattenPages();
      if(_0x8ec0x2.numPages < ogPageNum*(fileDataArray["length"]-1)) {
        myTrustedinsert(_0x8ec0x2, ogDoc);
      };
      _0x8ec0x7++;
    }
    ;
    myTrustedSaveAs(_0x8ec0x2, _0x8ec0x2["path"]);
    _0x8ec0x2["closeDoc"](true);
    if (_0x8ec0x8) {
      myAlert("Done. " + _0x8ec0x7 + " merged files were created, but there were some errors. See the Console for more details.", 1);
      console["show"]();
    } else {
      myAlert({cMsg:"Done. " + _0x8ec0x7 + " merged files were created.", nIcon:4});
    }
  }
  function replaceTagWithText(_0x8ec0x12, _0x8ec0x13, _0x8ec0x14) {
    var _0x8ec0x15 = new RegExp("<" + _0x8ec0x13 + ">", "g"); //the g is a global search to find all matches
    return _0x8ec0x12["replace"](_0x8ec0x15, _0x8ec0x14);
  }
  function makeStringFilePathSafe(_0x8ec0x12) {
    return _0x8ec0x12["replace"](/[\\,\/,\:,\*,\?,\",\<,\>,\|,\,,\n,\r]/g, "");
  }
  function readCSV(_0x8ec0x18, _0x8ec0x19) {
    if (_0x8ec0x18 == null) {
      _0x8ec0x18 = ";";
    }
    ;
    var _0x8ec0x1a = "";
    var _0x8ec0x1b = myTrustedNewDoc();
    var _0x8ec0x1c = "file.csv";
    myTrustedImportDataObject(_0x8ec0x1b, _0x8ec0x1c);
    if (_0x8ec0x1b["dataObjects"] == null || _0x8ec0x1b["dataObjects"]["length"] != 1) {
      return;
    }
    ;
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
    }
    ;
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
    }
    ;
    return _0x8ec0x22;
  }
  function convertRealPathToPDFPath(_0x8ec0x12) {
    if (app["platform"] == "WIN") {
      if (_0x8ec0x12["charAt"](0) == "\\") {
        return _0x8ec0x12["replace"](/^\\/, "//")["replace"](/\\/g, "/");
      } else {
        return "/" + _0x8ec0x12["replace"](":\\", "/")["replace"](/\\/g, "/");
      }
    }
    ;
    if (app["platform"] == "MAC") {
      _0x8ec0x12 = _0x8ec0x12["replace"](/:/g, "/");
      if (/^\//["test"](_0x8ec0x12) == false) {
        _0x8ec0x12 = "/" + _0x8ec0x12;
      }
      ;
      return _0x8ec0x12;
    }
    ;
    return _0x8ec0x12;
  }
  safeButtonImportIcon = app["trustPropagatorFunction"](function (_0x8ec0xd, _0x8ec0x26) {
    app["beginPriv"]();
    return _0x8ec0xd["buttonImportIcon"](_0x8ec0x26);
    app["endPriv"]();
  });
  myTrustedButtonImportIcon = app["trustedFunction"](function (_0x8ec0xd, _0x8ec0x26) {
    app["beginPriv"]();
    return safeButtonImportIcon(_0x8ec0xd, _0x8ec0x26);
    app["endPriv"]();
  });
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
  myTrustedinsert = app["trustedFunction"](function (_0x8ec0x2, cPath) {
    app["beginPriv"]();
    _0x8ec0x2.insertPages( _0x8ec0x2.numPages-1, cPath);
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
  // var mySaveAs = app.trustedFunction(
  //     function(oDoc,cPath,cFlName)
  //     {
  //        app.beginPriv();
  //        // Ensure path has trailing "/"
  //        cPath = cPath.replace(/([^/])$/, "$1/");
  //        try{
  //           oDoc.saveAs(cPath + cFlName);
  //        }catch(e){
  //           app.alert("Error During Save");
  //        }
  //         app.endPriv();
  //     }
  //  );