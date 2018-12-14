var app = new Vue({
    el: "#app",
    data() {
        return {
            wb: '',
            tableHeader: [],
            tableTbody: [],
            isAddPrefix: false,
            prefixContent:"",
            cate9: ['户数 Number of Households','男性人口 Male Population','女性人口 Female Population','流动人口/暂住人口 Migratory/Temporary Population','出生人数 Number of Births','自然出生率 Birth Rate (‰)','死亡人数 Number of Deaths','死亡率 Death Rate (‰)','迁入 - 男 Migration In - Male','迁入 - 女 Migration In - Female','迁入 - 总 Migration In - Total','迁出 - 男 Migration Out - Male','迁出 - 女 Migration Out - Female','迁出 - 总 Migration Out - Total','知识青年迁出 Educated Youth - Migration Out','农转非 (人数) Agricultural to Non-Agricultural Hukou (Change of Residency Status) (number of people)','自然增长率 Natural Population Growth Rate (‰)','农转非 (户数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of households)','总人口 Total Population','知识青年迁入 Educated Youth - Migration In']
        }
    },
    methods: {
        exportExcel() {
            if (files.length == 0) {
                alert("Please upload files.");
                return;
            }
            for ( idx = 0 ; idx<files.length ; idx++ ) {
                /* Form a new file name */
                var filename = (this.isAddPrefix)? this.prefixContent+files[idx].name : files[idx].name;
                this.file2Xce(files[idx], filename, this.cate9).then(tabJson => {
                    if (tabJson['data'] && tabJson['data'].length > 0) {
                        /* this line is only needed if you are not adding a script tag reference */
                        if(typeof XLSX == 'undefined') XLSX = require('xlsx');

                        /* make the worksheet */
                        var ws = XLSX.utils.json_to_sheet(tabJson['data'], {skipHeader: 1});

                        /* add to workbook */
                        var wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Form");

                        /* generate an XLSX file */
                        XLSX.writeFile(wb, tabJson['filename']);
                    }
                });
            }
        },
        importExcel(obj) {
            if (!obj.files) {
                return;
            }
            let file = obj.files[0],
                types = file.name.split('.').pop(),
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
        file2Xce(file, name, cate) {
            return new Promise(function (resolve, reject) {
                let reader = new FileReader();
                reader.onload = function (e) {
                    let data = e.target.result;
                    this.wb = XLSX.read(data, {
                        type: 'binary',
                        cellDates:true
                    });

                    var ws = this.wb.Sheets[this.wb.SheetNames[0]];
                    var range = XLSX.utils.decode_range(ws['!ref']);
                    var colNum, rowNum;

                    // Get year range and the 'Is the data available?'
                    var yearRange = {'begin': null, 'end': null};
                    var isAvailableCol, categoryCol, codeCol;
                    for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
                        var cell = ws[ XLSX.utils.encode_cell({r: 0, c: colNum}) ];
                        if (cell.w.length == 4 && !isNaN(cell.w)) {
                            if (yearRange.begin == null && yearRange.end == null) {
                                yearRange.begin = colNum;
                            } else if (yearRange.begin != null) {
                                yearRange.end = colNum;
                            }
                        } else if (cell.w == 'Is the data available?') {
                            isAvailableCol = colNum;
                        } else if (cell.w == 'Start Year') {
                            yearRange.begin = colNum;
                        } else if (cell.w == 'Data') {
                            yearRange.end = colNum;
                        } else if (cell.w == 'Category') {
                            categoryCol = colNum;
                        } else if (cell.w == '村志代码 Gazetteer Code') {
                            codeCol = colNum;
                        }
                    }

                    var groups = new Array(), group = new Array();
                    var tempCate = Array.from(cate);
                    for (rowNum = 1; rowNum <= range.e.r; rowNum++) {
                         var cellCategory = ws[ XLSX.utils.encode_cell({r: rowNum, c: categoryCol})].w;
                         var cellCode = ws[ XLSX.utils.encode_cell({r: rowNum, c: codeCol})].w;
                         var cur_code = cellCode;
                         var index = tempCate.indexOf(cellCategory);
                         console.log("code: "+cellCode+" category: " +cellCategory+ " index: "+index);
                         if (index >= 0) {
                             var isRefNextRow = false;
                             var obj = { "title":cellCategory, "row": rowNum};
                             group.push(obj);
                             tempCate.splice(index, 1);
                             if (rowNum+1<=range.e.r)
                                isRefNextRow = (cur_code ==  ws[ XLSX.utils.encode_cell({r: rowNum+1, c: codeCol})].w)?false:true;
                             if (rowNum == range.e.r || isRefNextRow) {
                                 for ( i = 0 ; i < tempCate.length ; i++ ) {
                                     var obj = { "title":tempCate[i], "row": -1};
                                     group.push(obj);
                                 }
                                 groups.push(group);
                                 tempCate = Array.from(cate);
                                 group = [];
                             }
                         }
                    }
                    console.log(groups);
                    // // Insert new cells into table
                    // if (yearRange.begin != null) {
                    //     var id;
                    //     for(rowNum = 1; rowNum <= range.e.r; rowNum++) {
                    //         var isAvailable = false;
                    //         for(colNum  = yearRange.begin; colNum <= yearRange.end; colNum++) {
                    //             var cell = ws[ XLSX.utils.encode_cell({r: rowNum, c: colNum}) ];
                    //             if (cell === undefined) {
                    //                 var str;
                    //                 if (colNum <= range.e.c) {
                    //                     str = 'n/a';
                    //                 }
                    //                 var cell = {t:'s', v: str, r: '<t>'+str+'</t>', w: str};
                    //                 ws[ XLSX.utils.encode_cell({r: rowNum, c: colNum}) ] = cell;
                    //             } else {
                    //                 if ( colNum >= yearRange.begin && colNum <= yearRange.end && !isAvailable) {isAvailable = true;}
                    //                 continue;
                    //             }
                    //         }
                    //         var str = (isAvailable) ? 'Yes': 'No';
                    //         var cell = {t:'s', v: str, r: '<t>'+str+'</t>', w: str};
                    //         ws[ XLSX.utils.encode_cell({r: rowNum, c: isAvailableCol}) ] = cell;
                    //     }
                    // }
                    var result = {'data':XLSX.utils.sheet_to_json(ws, {header: 1, raw: false, blankCell: false}), 'filename': name};
                    resolve(result);
                };
                reader.readAsBinaryString(file);
            });
        }
    }
})
