var app = new Vue({
    el: "#app",
    data() {
        return {
            wb: "",
            tableHeader: [],
            tableTbody: [],
            isAddPrefix: false,
            fileArr: [],
            prefixContent:"",
            dataSourceOption: 1,
            exportOption: 0,
        }
    },
    methods: {
        exportExcel() {
            if (this.fileArr.length == 0) {
                alert("Please upload files.");
                return;
            }

            for ( idx = 0 ; idx<this.fileArr.length ; idx++ ) {
                var category;
                var divisions;
                var arr = this.fileArr[idx].name.split("__");
                var formNum = parseInt(arr[0]);
                switch (formNum) {
                    case 9:
                        category = (this.dataSourceOption == 1) ? dCate9 : cate9;
                        divisions = (this.dataSourceOption == 1) ? null : submenu9;
                        console.log(divisions);
                        break;
                    case 10:
                        category = (this.dataSourceOption == 1) ? dCate10 : cate10;
                        divisions = (this.dataSourceOption == 1) ? null : submenu10;
                        break;
                    case 11:
                        category = (this.dataSourceOption == 1) ? dCate11 : cate11;
                        divisions = (this.dataSourceOption == 1) ? null : submenu11;
                        break;
                    case 12:
                        category = (this.dataSourceOption == 1) ? dCate12 : cate12;
                        divisions = null;
                        break;
                    case 13:
                        category = (this.dataSourceOption == 1) ? dCate13 : cate13;
                        divisions = (this.dataSourceOption == 1) ? null : submenu13;
                        break;
                    case 14:
                        category = cate14;
                        break;
                    default:
                        alert("Cannot find this category.");
                        return;
                }
                /* Form a new file name */
                var filename = (this.isAddPrefix)? this.prefixContent+this.fileArr[idx].name : this.fileArr[idx].name;
                if (this.dataSourceOption == 1) {
                  this.file2Xce_drupal(this.fileArr[idx], filename, category, divisions, this.exportOption, this.dataSourceOption).then(tabJson => {
                      console.log(tabJson);
                      if (tabJson["data"] && tabJson["data"].length > 0) {
                          /* this line is only needed if you are not adding a script tag reference */
                          if(typeof XLSX == "undefined") XLSX = require("xlsx");

                          /* make the worksheet */
                          var ws = XLSX.utils.json_to_sheet(tabJson["data"], {skipHeader: 1});

                          /* add to workbook */
                          var wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, ws, "Form");

                          /* generate an XLSX file */
                          XLSX.writeFile(wb, tabJson["filename"]);
                      }
                  });
                } else {
                  this.file2Xce(this.fileArr[idx], filename, category, divisions, this.exportOption, this.dataSourceOption).then(tabJson => {
                      console.log(tabJson);
                      if (tabJson["data"] && tabJson["data"].length > 0) {
                          /* this line is only needed if you are not adding a script tag reference */
                          if(typeof XLSX == "undefined") XLSX = require("xlsx");

                          /* make the worksheet */
                          var ws = XLSX.utils.json_to_sheet(tabJson["data"], {skipHeader: 1});

                          /* add to workbook */
                          var wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, ws, "Form");

                          /* generate an XLSX file */
                          XLSX.writeFile(wb, tabJson["filename"]);
                      }
                  });
              }
            }
        },
        importExcel(obj) {
            if (!obj.files) {
                return;
            }
            let file = obj.files[0],
                types = file.name.split(".").pop(),
                fileType = ["xlsx", "xlc", "xlm", "xls", "xlt", "xlw", "csv"].some(item => item === types);
            if (!fileType) {
                alert("Cannot process this format.");
                return;
            }
            this.file2Xce(file).then(tabJson => {
                if (tabJson && tabJson.length > 0) {
                    this.tableHeader = Object.keys(tabJson[0]);
                    this.tableTbody = tabJson;
                }
            });
        },
        file2Xce_drupal(file, name, cate, divisions, exportOption, dataSourceOption) {
          return new Promise(function (resolve, reject) {
              let reader = new FileReader();
              reader.onload = function (e) {
                  let data = e.target.result;
                  this.wb = XLSX.read(data, {
                      type: "binary",
                      cellDates:true
                  });

                  var startIndex = 2; // The excel files come with two exmpty rows in the beginning.
                  var ws = this.wb.Sheets[this.wb.SheetNames[0]];
                  var range = XLSX.utils.decode_range(ws["!ref"]);
                  var colNum, rowNum;
                  // Get year range and the "Is the data available?"
                  var mode = (name.indexOf("range") > 0) ? "range" : "yearly";
                  var yearRange = {"begin": null, "end": null};
                  var divisionRange = {"begin": null, "end": null};
                  var isAvailableCol, categoryCol, codeCol, titleCol;
                  var bookIdx = {};
                  var divisionColNum = {"category":1, "division1": 1, "division2": 2, "division3": 3, "division4": 4, "subdivision":5};
                  for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
                      var cell = ws[ XLSX.utils.encode_cell({r: startIndex, c: colNum}) ];
                      if (cell.w.length == 4 && !isNaN(cell.w)) {
                          if (yearRange.begin == null && yearRange.end == null) {
                              yearRange.begin = colNum;
                          } else if (yearRange.begin != null) {
                              yearRange.end = colNum;
                          }
                      } else if (cell.w == "Is the data available?") {
                          isAvailableCol = colNum;
                      } else if (cell.w == "Start Year") {
                          yearRange.begin = colNum;
                      } else if (cell.w == "Data") {
                          yearRange.end = colNum;
                      } else if (cell.w == "Category") {
                          categoryCol = colNum;
                      } else if (cell.w == "村志代码 Gazetteer Code") {
                          codeCol = colNum;
                      } else if (cell.w == "村志书名 Gazetteer Title") {
                          titleCol = colNum;
                      } else if (cell.w.indexOf("ivision") > 0) {
                        if (!divisionRange.begin || colNum < divisionRange.begin) {
                          divisionRange.begin = colNum;
                        }
                        if (!divisionRange.end || colNum > divisionRange.end) {
                          divisionRange.end = colNum;
                        }
                      } else if (cell.w == "Unit") {
                        divisionRange.end = colNum;
                      }
                  }

                  // Create Header
                  var content = new Array();
                  var header = new Array();
                  for(colNum  = codeCol; colNum <= yearRange.end; colNum++) {
                      header.push(ws[ XLSX.utils.encode_cell({r: startIndex, c: colNum}) ]);
                  }

                  header.splice(3, 0, "Is available?");
                  // header.splice(4, 0, "Category");
                  content.push(header);

                  // Rename category and get contents of code, title, category
                  var groups = {};
                  for (rowNum = startIndex + 1; rowNum <= range.e.r; rowNum++) {
                      var cellCode = ws[ XLSX.utils.encode_cell({r: rowNum, c: codeCol})].w;
                      var cellTitle = ws[ XLSX.utils.encode_cell({r: rowNum, c: titleCol})].w;
                      var cellCategory = ws[ XLSX.utils.encode_cell({r: rowNum, c: categoryCol})].w;
                      var cellAvailable = ws[ XLSX.utils.encode_cell({r: rowNum, c: isAvailableCol})];

                      var divisionSet = [cellCategory];
                      for (cc = divisionRange.begin; cc <= divisionRange.end ; cc++) {
                        var celldivision = ws[ XLSX.utils.encode_cell({r: rowNum, c: cc})];
                        if (celldivision != undefined && celldivision.x != "Subdivision") {
                          divisionSet.push(celldivision.w);
                        } else {
                          divisionSet.push("");
                        }
                      }

                      if (groups[cellCode] == null) {
                          groups[cellCode] = new Array();
                      }
                      groups[cellCode].push({"row": rowNum, "title": cellTitle, "category": cellCategory, "tag": divisionSet.join("^"), "isAvailable": (cellAvailable == undefined)?"Yes":cellAvailable.w});

                      if (bookIdx[cellCode] == null) {
                          bookIdx[cellCode] = cellTitle;
                      }
                  }

                  // console.log("group: "+JSON.stringify(groups));
                  // console.log("Books: "+JSON.stringify(bookIdx));

                  // Compare with the default category list remove unwanted rows and insert non-existed row from default category
                  for (var code in groups) {
                    var tempCate = Array.from(cate);
                    var delArr = new Array();
                    for (var i = 0 ; i < groups[code].length ; i++) {
                      var obj = groups[code][i];
                      var index = tempCate.indexOf(obj["tag"]);
                      // console.log("index: "+index+", tag: "+obj["tag"]);
                      if (index >= 0 && obj["isAvailable"] == "Yes") {
                          tempCate.splice(index, 1);
                      } else if (index == -1 && cate.indexOf(obj["tag"]) >= 0 && mode == "range") {
                          // TOOD: double check here
                          if (ws[ XLSX.utils.encode_cell({r: obj["row"], c: yearRange.begin})] == undefined) {
                              delArr.push(i);
                          }
                      } else {
                          delArr.push(i);
                      }
                    }
                    // console.log("delete:"+ delArr);
                    // Remove duplicated row
                    // if (delArr.length > 0 ) {
                    //     for (var i = delArr.length-1; i >= 0 ; i--) {
                    //         console.log(delArr[i]);
                    //         delete groups[code][delArr[i]];
                    //     }
                    // }
                    if (tempCate.length > 0) {
                        for (var j = 0; j < tempCate.length; j++) {

                          groups[code].push({"row": -1, "title": bookIdx[code], "category": tempCate[j]});
                        }
                    }
                  }
                  // console.log(JSON.stringify(groups));

                  // Generating the excel content
                  // console.log("divisions: "+JSON.stringify(divisionRange));
                  for (var code in groups) {
                      var counter = 0;
                      var rowContent = new Array();
                      var isAvailable = false;
                      for (obj of groups[code]) {
                          if (obj == undefined) continue;
                          rowContent.push(code);
                          rowContent.push(obj["title"]);
                          segs = obj['category'].split("^");
                          rowContent.push(segs[0]);

                          // Division Data
                          if (obj["row"] > 0) {
                            for (var i = divisionRange.begin ; i <= divisionRange.end ; i++) {
                                // rowContent.push(divisions[obj["category"]][i]);
                                // console.log(divisions[obj["category"]][i]);
                                var divisionCell = ws[ XLSX.utils.encode_cell({r: obj["row"], c: i}) ];
                                if (divisionCell != undefined) {
                                  rowContent.push(divisionCell.w);
                                } else {
                                  rowContent.push("");
                                }

                            }
                          } else {
                            for (var i = 0; i<divisionRange.end - divisionRange.begin + 1 ; i++) {
                              rowContent.push(segs[i]);
                            }
                          }

                          // // Year information
                          for(colNum = yearRange.begin; colNum <= yearRange.end; colNum++) {
                              if (obj["row"] > -1) {
                                  var cell = ws[ XLSX.utils.encode_cell({r: obj["row"], c: colNum}) ];
                                  if (cell === undefined || isNaN(cell.w)) {
                                      rowContent.push("n/a");
                                  } else {
                                      if (!isAvailable) {isAvailable = true;}
                                      rowContent.push(cell.w);
                                  }
                              }
                              else {
                                  rowContent.push("n/a");
                              }
                          }

                          counter++;
                          if (exportOption == 1 && !isAvailable) {

                          } else {
                              rowContent.splice(3,0, (isAvailable)?"Yes":"No");
                              content.push(Array.from(rowContent));
                          }
                          rowContent.length = 0;
                          isAvailable = false;
                      }
                  }
                  var result = {"data": content, "filename": name};
                  resolve(result);
              };
              reader.readAsBinaryString(file);
          });
        },
        file2Xce(file, name, cate, divisions, exportOption, dataSourceOption) {
            return new Promise(function (resolve, reject) {
                let reader = new FileReader();
                reader.onload = function (e) {
                    let data = e.target.result;
                    this.wb = XLSX.read(data, {
                        type: "binary",
                        cellDates:true
                    });

                    var startIndex = 2; // The excel files come with two exmpty rows in the beginning.
                    var ws = this.wb.Sheets[this.wb.SheetNames[0]];
                    var range = XLSX.utils.decode_range(ws["!ref"]);
                    var colNum, rowNum;
                    // Get year range and the "Is the data available?"
                    var mode = (name.indexOf("range") > 0) ? "range" : "yearly";
                    var yearRange = {"begin": null, "end": null};
                    var isAvailableCol, categoryCol, codeCol, titleCol;
                    var bookIdx = {};
                    var divisionColNum = {"category":1, "division1": 1, "division2": 2, "division3": 3, "division4": 4, "subdivision":5};
                    for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
                        var cell = ws[ XLSX.utils.encode_cell({r: startIndex, c: colNum}) ];
                        console.log(cell);
                        if (cell.w.length == 4 && !isNaN(cell.w)) {
                            if (yearRange.begin == null && yearRange.end == null) {
                                yearRange.begin = colNum;
                            } else if (yearRange.begin != null) {
                                yearRange.end = colNum;
                            }
                        } else if (cell.w == "Is the data available?") {
                            isAvailableCol = colNum;
                        } else if (cell.w == "Start Year") {
                            yearRange.begin = colNum;
                        } else if (cell.w == "Data") {
                            yearRange.end = colNum;
                        } else if (cell.w == "Category") {
                            categoryCol = colNum;
                        } else if (cell.w == "村志代码 Gazetteer Code") {
                            codeCol = colNum;
                        } else if (cell.w == "村志书名 Gazetteer Title") {
                            titleCol = colNum;
                        }
                    }

                    // Create Header
                    var content = new Array();
                    var header = new Array();
                    for(colNum  = codeCol; colNum <= yearRange.end; colNum++) {
                        header.push(ws[ XLSX.utils.encode_cell({r: startIndex, c: colNum}) ]);
                    }
                    // var divisionHead = ["category", "division1", "divisio2", "division3","division4", "subdivision"];
                    // for (var i = 1; i <= divisionHead.length; i++) {
                    //     ws[ XLSX.utils.encode_cell({r: 0, c: yearRange.end+1})].w = divisionHead[i];
                    //     var cell = ws[ XLSX.utils.encode_cell({r: 0, c: yearRange.end+1})];
                    //     cell.w = divisionHead[i];
                    //     header.push(cell);
                    // }
                    header.splice(3, 0, "Is available?");
                    header.splice(4, 0, "Category");
                    console.log(header);
                    content.push(header);

                    // Rename category and get contents of code, title, category
                    var groups = {};
                    for (rowNum = startIndex + 1; rowNum <= range.e.r; rowNum++) {
                        var cellCode = ws[ XLSX.utils.encode_cell({r: rowNum, c: codeCol})].w;
                        var cellTitle = ws[ XLSX.utils.encode_cell({r: rowNum, c: titleCol})].w;
                        var cellCategory = ws[ XLSX.utils.encode_cell({r: rowNum, c: categoryCol})].w;
                        var cellAvailable = ws[ XLSX.utils.encode_cell({r: rowNum, c: isAvailableCol})];
                        // Rename the category
                        if (cellCategory == "农转非 (人数) Agricultural to Non-Agricultural Hukou (Change of Residency Status) (number of people)")
                            cellCategory = "农转非 (人数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of people)";
                        else if (cellCategory == "耕地面积 (平方米) Cultivated Area (square meters)")
                            cellCategory = "耕地面积 (公顷) Cultivated Area (hectares)";
                        else if (cellCategory == "上环 Use of Interuterine Device (IUD)")
                            cellCategory = "上环 Use of Intrauterine Device (IUD)";
                        else if (cellCategory == "用电量 - 每人 (度) Village Power Consumption - per person (kilowatt hours)")
                            cellCategory = "生活用电量 - 每人 (度) Electricity Consumption - per person (kilowatt hours)";
                        else if (cellCategory == "用电量 - 每户 (度) Village Power Consumption - per household (kilowatt hours)")
                            cellCategory = "生活用电量 - 每户 (度) Electricity Consumption - per household (kilowatt hours)";
                        else if (cellCategory == "用电量 - 全村 (度) Village Power Consumption - Village (kilowatt-hours)")
                            cellCategory = "生活用电量 - 全村 (度) Electricity Consumption - village (kilowatt-hours)";
                        else if (cellCategory == "计划生育参与率 Family Planning Program Participation Rate (%)")
                            cellCategory = "计划生育率 (%) Planned Pregnancy Rate (%)";

                        if (groups[cellCode] == null) {
                            groups[cellCode] = new Array();
                        }
                        groups[cellCode].push({"row": rowNum, "title": cellTitle, "category": cellCategory, "isAvailable": (cellAvailable == undefined)?"Yes":cellAvailable.w});

                        if (bookIdx[cellCode] == null) {
                            bookIdx[cellCode] = cellTitle;
                        }
                    }

                    // Compare with the default category list remove unwanted rows and insert non-existed row from default category
                    for (var code in groups) {
                        var tempCate = Array.from(cate);
                        var delArr = new Array();
                        for (var i = 0; i < groups[code].length; i++) {
                            var obj = groups[code][i];
                            var index = tempCate.indexOf(obj["category"]);
                            if (index >= 0 && obj["isAvailable"] == "Yes") {
                                tempCate.splice(index, 1);
                            } else if (index == -1 && cate.indexOf(obj["category"]) >= 0 && mode == "range") {
                                // TOOD: double check here
                                if (ws[ XLSX.utils.encode_cell({r: obj["row"], c: yearRange.begin})] == undefined) {
                                    delArr.push(i);
                                }
                            } else {
                                delArr.push(i);
                            }
                        }
                        console.log("delete:"+ delArr);
                        if (delArr.length > 0 ) {
                            for (var i = delArr.length-1; i >= 0 ; i--) {
                                console.log(delArr[i]);
                                delete groups[code][delArr[i]];
                            }
                        }
                        if (tempCate.length > 0) {
                            for (var j = 0; j < tempCate.length; j++) {
                                groups[code].push({"row": -1, "title": bookIdx[code], "category": tempCate[j]});
                            }
                        }
                    }

                    // Generating the excel content
                    for (var code in groups) {
                        var counter = 0;
                        console.log(groups[code]);
                        var rowContent = new Array();
                        var isAvailable = false;
                        for (obj of groups[code]) {
                            if (obj == undefined) continue;
                            rowContent.push(code);
                            rowContent.push(obj["title"]);
                            rowContent.push(obj["category"]);

                            // Division Data
                            if (divisions != null) {
                                for (var i = 0 ; i < divisions[obj["category"]].length ; i++) {
                                    rowContent.push(divisions[obj["category"]][i]);
                                    console.log(divisions[obj["category"]][i]);
                                }
                            }

                            while(rowContent.length < 8) {
                              rowContent.push("");
                            }
                            // Year information
                            for(colNum = yearRange.begin+2; colNum <= yearRange.end+2; colNum++) {
                                if (obj["row"] > -1) {
                                    var cell = ws[ XLSX.utils.encode_cell({r: obj["row"], c: colNum}) ];
                                    if (cell === undefined || isNaN(cell.w)) {
                                        rowContent.push("n/a");
                                    } else {
                                        if (!isAvailable) {isAvailable = true;}
                                        rowContent.push(cell.w);
                                    }
                                }
                                else {
                                    rowContent.push("n/a");
                                }
                            }

                            counter++;
                            if (exportOption == 1 && !isAvailable) {

                            } else {
                                rowContent.splice(3,0, (isAvailable)?"Yes":"No");
                                content.push(Array.from(rowContent));
                            }
                            rowContent.length = 0;
                            isAvailable = false;
                        }
                    }
                    var result = {"data": content, "filename": name};
                    resolve(result);
                };
                reader.readAsBinaryString(file);
            });
        },
        loadTextFromFile(ev) {
            const file = ev.target.files[0];
            // console.log("before:"+ev.target.files.length);
            var notValidFile = "";
            for (var i = 0 ; i < ev.target.files.length ; i++) {
                let file = ev.target.files[i],
                    types = file.name.split(".").pop(),
                    fileType = ["xlsx", "xlc", "xlm", "xls", "xlt", "xlw", "csv"].some(item => item === types);
                if (!fileType) {
                    notValidFile = notValidFile + "\n" + file.name;
                } else
                    this.fileArr.push(ev.target.files[i]);
            }
            if (notValidFile != "") {
                alert(notValidFile+"\n cannot be converted.");
            }
        },
        clearFileArr() {
            this.fileArr = [];
        },
        deleteFile(idx) {
            this.fileArr.splice(idx, 1);
        }
    }
})
