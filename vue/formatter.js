var app = new Vue({
    el: "#app",
    data() {
        return {
            wb: '',
            tableHeader: [],
            tableTbody: []
        }
    },
    methods: {
        exportExcel() {
            console.log(files);
            for ( idx = 0 ; idx<files.length ; idx++ ) {
                this.file2Xce(files[idx]).then(tabJson => {
                    if (tabJson && tabJson.length > 0) {
                        // this.tableHeader = Object.keys(tabJson[0]);
                        // this.tableTbody = tabJson;
                        /* this line is only needed if you are not adding a script tag reference */
                        if(typeof XLSX == 'undefined') XLSX = require('xlsx');

                        /* make the worksheet */
                        var ws = XLSX.utils.json_to_sheet(tabJson);

                        /* add to workbook */
                        var wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Form");

                        /* generate an XLSX file */
                        XLSX.writeFile(wb, "sheetjs.xlsx");
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
        file2Xce(file) {
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

                    // Get year range
                    var yearRange = {'begin': null, 'end': null};
                    for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
                        var cell = ws[ XLSX.utils.encode_cell({r: 0, c: colNum}) ];
                        if (cell.w.length == 4 && !isNaN(cell.w)) {
                            if (yearRange.begin == null && yearRange.end == null) {
                                yearRange.begin = colNum;
                            } else if (yearRange.begin != null) {
                                yearRange.end = colNum;
                            }
                        }
                    }

                    // Insert new cells into table
                    if (yearRange.begin != null) {
                        for(rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
                            var isAvailable = false;
                            for(colNum  =range.s.c; colNum <= range.e.c; colNum++) {
                                var cell = ws[ XLSX.utils.encode_cell({r: rowNum, c: colNum}) ];
                                if (cell === undefined) {
                                    var str;
                                    if (colNum < range.e.c) {
                                        str = 'n/a';
                                    } else {
                                        str = (isAvailable) ? 'Yes': 'No';
                                    }
                                    var cell = {t:'s', v: str, r: '<t>'+str+'</t>', w: str};
                                    ws[ XLSX.utils.encode_cell({r: rowNum, c: colNum}) ] = cell;
                                } else {
                                    if (isNaN(cell.w) && colNum >= yearRange.begin && colNum <= yearRange.end && !isAvailable) isAvailable = true;
                                    continue;
                                }
                            }
                        }
                    }
                    resolve(XLSX.utils.sheet_to_json(ws, {header: 1, raw: false, blankCell: false}));
                };
                reader.readAsBinaryString(file);
            });
        }
    }
})
