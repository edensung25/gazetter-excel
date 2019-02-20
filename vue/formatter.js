var app = new Vue({
    el: "#app",
    data() {
        return {
            wb: '',
            tableHeader: [],
            tableTbody: [],
            isAddPrefix: false,
            fileArr: [],
            prefixContent:"",
            cate9: ['户数 Number of Households',
                    '男性人口 Male Population',
                    '女性人口 Female Population',
                    '总人口 Total Population',
                    '流动人口/暂住人口 Migratory/Temporary Population',
                    '出生人数 Number of Births',
                    '自然出生率 Birth Rate (‰)',
                    '死亡人数 Number of Deaths',
                    '死亡率 Death Rate (‰)',
                    '自然增长率 Natural Population Growth Rate (‰)',
                    '迁入 (户数) Migration In (number of households)',
                    '迁出 (户数) Migration Out (number of households)',
                    '知识青年迁入 Educated Youth - Migration In',
                    '知识青年迁出 Educated Youth - Migration Out',
                    '农转非 (人数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of people)',
                    '农转非 (户数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of households)',
                    '迁入 (人数) Migration In (number of people)',
                    '迁出 (人数) Migration Out (number of people)',
                    '视力残疾 Blindness',
                    '听力语言残疾 Hearing and Speech Disabilities',
                    '肢体残疾 Amputation and/or Paralysis',
                    '精神残疾 Mental Disabilities',
                    '智力残疾 Intellectual Disabilities',
                    '残疾人总数 Total Disabled Population'],
            cate10: ['入伍 Military Enlistment',
                     '村民纠纷 Number of Civil Mediations',
                     '刑事案件 Number of Reported Crimes',
                     '共产党员 - 男 CCP Membership - Male',
                     '共产党员 - 女 CCP Membership - Female',
                     '共产党员 - 总 CCP Membership - Total',
                     '共产党员 - 少数民族 CCP Membership - Ethnic Minorities',
                     '新党员 - 男 New CCP Membership - Male',
                     '新党员 - 女 New CCP Membership - Female',
                     '新党员 - 总 New CCP Membership - Total',
                     '新党员 - 少数民族 New CCP Membership - Ethnic Minorities',
                     '阶级成分 - 地主 (户数) Class Status - Landlord (number of households)',
                     '阶级成分 - 富农 (户数) Class Status - Rich Peasant (number of households)',
                     '阶级成分 - 中农 (户数) Class Status - Middle Peasant (number of households)',
                     '阶级成分 - 贫下中农 (户数) Class Status - Poor and Lower Middle Peasant (number of households)'],
            cate11: ['总产值 (万元) Gross Output Value (10K yuan)',
                     '集体经济收入 (元) Collective Economic Income (yuan)',
                     '集体经济收入 (万元) Collective Economic Income (10K yuan)',
                     '耕地面积 (亩) Cultivated Area (mu)',
                     '耕地面积 (公顷) Cultivated Area (hectares)',
                     '粮食总产量 (万斤) Total Grain Output (10K pounds)',
                     '粮食总产量 (吨) Total Grain Output (tons)',
                     '粮食总产量 (万吨) Total Grain Output (10K tons)',
                     '生活用电量 - 每人 (度) Electricity Consumption - per person (kilowatt hours)',
                     '生活用电量 - 每户 (度) Electricity Consumption - per household (kilowatt hours)',
                     '生活用电量 - 全村 (度) Electricity Consumption - village (kilowatt-hours)',
                     '电价 (元每度) Electricity Price - General (yuan per kilowatt-hour)',
                     '电价 - 生活用电 (元每度) Electricity Price - Domestic (yuan per kilowatt-hour)',
                     '电价 - 农业用电 (元每度) Electricity Price - Agricultural (yuan per kilowatt-hour)',
                     '电价 - 工业用电 (元每度) Electricity Price - Industrial (yuan per kilowatt-hour)',
                     '水价 (元每立方米) Water Price (yuan per cubic meter)',
                     '人均收入 (元) Per Capita Income (yuan)',
                     '人均居住面积 (平方米) Per Capita Living Space (square meters)',
                     '工业用电量 Industrial Electricity Consumption',
                     '农业用电量 Agricultural Electricity Consumption'],
            cate12: ['计划生育参与率 Family Planning Program Participation Rate (%)',
                     '领取独生子女证 (人数) Certified Commitment to One Child Policy (number of people)',
                     '育龄妇女人口 Number of Women of Childbearing Age',
                     '男性结扎 Vasectomies',
                     '女性结扎 Tubal Ligations',
                     '上环 Use of Intrauterine Device (IUD)',
                     '绝育手术 Sterilization Surgeries',
                     '节育率 (%) Rate of Contraception (%)',
                     '结扎总数 Total Number of Vasectomies and Tubal Ligations',
                     '人工流产 Abortions'],
            cate13: ['在校生 - 小学 Students in School - Elementary School',
                     '在校生 - 初中 Students in School - Junior High School',
                     '在校生 - 高中 Students in School - High School',
                     '新入学生 - 大学 Initial Student Enrollment - College/University',
                     '老师- 小学 Teachers - Elementary School',
                     '老师- 初中 Teachers - Junior High School',
                     '老师- 高中 Teachers - High School',
                     '受教育程度 - 文盲 Illiterate',
                     '受教育程度 - 小学 Highest Level of Education - Elementary School',
                     '受教育程度 - 初中 Highest Level of Education - Junior High School',
                     '受教育程度 - 中专高中 Highest Level of Education - High School',
                     '受教育程度 - 大专以上 Highest Level of Education - College/University or Higher'],
            cate14: ['视力残疾 Blindness',
                     '听力语言残疾 Hearing and Speech Disabilities',
                     '肢体残疾 Amputation and/or Paralysis',
                     '精神残疾 Mental Disabilities',
                     '智力残疾 Intellectual Disabilities',
                     '残疾人总数 Total Disabled Population'],
            submenu9: {
                '户数 Number of Households': ['户数 Number of Household'],
                '男性人口 Male Population': ['人口 Population','男性人口 Male Population'],
                '女性人口 Female Population': ['人口 Population','女性人口 Female Population'],
                '总人口 Total Population': ['人口 Population','总人口 Total Population'],
                '出生人数 Number of Births': ['出生人数 Number of Births'],
                '自然出生率 Birth Rate (‰)': ['自然出生率 Birth Rate (‰)'],
                '死亡人数 Number of Deaths': ['死亡人数 Number of Deaths'],
                '死亡率 Death Rate (‰)': ['死亡率 Death Rate (‰)'],
                '自然增长率 Natural Population Growth Rate (‰)':[ '自然增长率 Natural Population Growth Rate (‰)'],
                '流动人口/暂住人口 Migratory/Temporary Population': ['流动人口/暂住人口 Migratory/Temporary Population'],
                '迁入 (户数) Migration In (number of households)': ['迁入 Migration In', ,'','户数 number of households'],
                '迁出 (户数) Migration Out (number of households)': ['迁出 Migration Out', ,'','户数 number of households'],
                '知识青年迁入 Educated Youth - Migration In': ['迁入 Migration In', ,'','知识青年 Educated Youth'],
                '知识青年迁出 Educated Youth - Migration Out': ['迁入 Migration In', ,'','知识青年 Educated Youth'],
                '农转非 (人数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of people)':['农转非 Agricultural to Non-Agricultural Hukou / Change of Residency Status','','','人数 number of people'],
                '农转非 (户数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of households)':['农转非 Agricultural to Non-Agricultural Hukou / Change of Residency Status','','','户数 number of households'],
                '迁入 (人数) Migration In (number of people)': ['迁入 Migration In', ,'','人数 number of people'],
                '迁出 (人数) Migration Out (number of people)': ['迁出 Migration Out', ,'','人数 number of people'],
                '迁出 (人数) Migration Out (number of people)': ['迁出 Migration Out', '','人数 number of people'],
                '视力残疾 Blindness': ['残疾人数 Disabled Population', '','','','视力残疾 Blindness'],
                '听力语言残疾 Hearing and Speech Disabilities': ['残疾人数 Disabled Population', '','','','听力语言残疾 Hearing and Speech Disabilities'],
                '肢体残疾 Amputation and/or Paralysis': ['残疾人数 Disabled Population', '','','','肢体残疾 Amputation and/or Paralysis'],
                '精神残疾 Mental Disabilities': ['残疾人数 Disabled Population', '','','', '精神残疾 Mental Disabilities'],
                '智力残疾 Intellectual Disabilities': ['残疾人数 Disabled Population','','','', '智力残疾 Intellectual Disabilities'],
                '残疾人总数 Total Disabled Population': ['残疾人数 Disabled Population','','','', '残疾人总数 Total Disabled Population'],
            },
            submenu10:{
                '入伍 Military Enlistment': ['入伍 Military Enlistment'],
                '村民纠纷 Number of Civil Mediations': ['村民纠纷 Number of Civil Mediations'],
                '刑事案件 Number of Reported Crimes': ['刑事案件 Number of Reported Crimes'],
                '共产党员 - 男 CCP Membership - Male': ['共产党员 CCP Membership','男 Male'],
                '共产党员 - 女 CCP Membership - Female': ['共产党员 CCP Membership','女 Female'],
                '共产党员 - 总 CCP Membership - Total': ['共产党员 CCP Membership','总 Total'],
                '共产党员 - 少数民族 CCP Membership - Ethnic Minorities': ['共产党员 CCP Membership','少数民族 Ethnic Minorities'],
                '新党员 - 男 New CCP Membership - Male': ['新党员 New CCP Membership','男 Male'],
                '新党员 - 女 New CCP Membership - Female': ['新党员 New CCP Membership','女 Female'],
                '新党员 - 总 New CCP Membership - Total': ['新党员 New CCP Membership','总 Total'],
                '新党员 - 少数民族 New CCP Membership - Ethnic Minorities': ['新党员 New CCP Membership','少数民族 Ethnic Minorities'],
                '阶级成分 - 地主 (户数) Class Status - Landlord (number of households)': ['阶级成分 Class Status','','地主 Landlord'],
                '阶级成分 - 富农 (户数) Class Status - Rich Peasant (number of households)': ['阶级成分 Class Status','','富农 Rich Peasant'],
                '阶级成分 - 中农 (户数) Class Status - Middle Peasant (number of households)': ['阶级成分 Class Status','','中农 Middle Peasant'],
                '阶级成分 - 贫下中农 (户数) Class Status - Poor and Lower Middle Peasant (number of households)': ['阶级成分 Class Status','','贫下中农 Poor and Lower Middle Peasant'],
            },
            submenu11:{
                '总产值 (万元) Gross Output Value (10K yuan)':['总产值 Gross Output Value','','','','','万元 10K yuan'],
                '集体经济收入 (元) Collective Economic Income (yuan)':['集体经济收入 Collective Economic Income','','','','','元 yuan'],
                '集体经济收入 (万元) Collective Economic Income (10K yuan)':['集体经济收入 Collective Economic Income','','','','','万元 10K yuan'],
                '耕地面积 (亩) Cultivated Area (mu)':['耕地面积 Cultivated Area','','','','','亩 mu'],
                '耕地面积 (公顷) Cultivated Area (hectares)':['耕地面积 Cultivated Area','','','','','公顷 hectares'],
                '粮食总产量 (万斤) Total Grain Output (10K pounds)':['粮食总产量 Total Grain Output','','','','','万斤 10K jin'],
                '粮食总产量 (吨) Total Grain Output (tons)':['粮食总产量 Total Grain Output','','','','','吨 tons'],
                '粮食总产量 (万吨) Total Grain Output (10K tons)':['粮食总产量 Total Grain Output','','','','','万吨 10K tons'],
                '生活用电量 - 每人 (度) Electricity Consumption - per person (kilowatt hours)':['用电量 Electricity Consumption','','','生活 Household','每人 per person','度 kilowatt hours'],
                '生活用电量 - 每户 (度) Electricity Consumption - per household (kilowatt hours)':['用电量 Electricity Consumption','','','生活 Household','每户 per household','度 kilowatt hours'],
                '生活用电量 - 全村 (度) Electricity Consumption - village (kilowatt-hours)':['用电量 Electricity Consumption','','','生活 Household','全村 village','度 kilowatt hours'],
                '电价 (元每度) Electricity Price - General (yuan per kilowatt-hour)':['电价 Electricity Price','','','','','元 yuan'],
                '电价 - 生活用电 (元每度) Electricity Price - Domestic (yuan per kilowatt-hour)':['电价 Electricity Price','','','生活 Household','','元 yuan'],
                '电价 - 农业用电 (元每度) Electricity Price - Agricultural (yuan per kilowatt-hour)':['电价 Electricity Price','','','农业 Agricultural','','元 yuan'],
                '电价 - 工业用电 (元每度) Electricity Price - Industrial (yuan per kilowatt-hour)':['电价 Electricity Price','','','工业 Industrial','','元 yuan'],
                '水价 (元每立方米) Water Price (yuan per cubic meter)':['水价 Water Price','','','','','元每立方米 yuan per cubic meter'],
                '人均收入 (元) Per Capita Income (yuan)':['人均收入 Per Capita Income','','','','','元 yuan'],
                '人均居住面积 (平方米) Per Capita Living Space (square meters)':['人均居住面积 Per Capita Living Space','','','','','立方米 cubic meters'],
                '工业用电量 Industrial Electricity Consumption':['用电量 Electricity Consumption','','','工业 Industrial','','度 kilowatt hours'],
                '农业用电量 Agricultural Electricity Consumption':['用电量 Electricity Consumption','','','农业 Agricultural','','度 kilowatt hours']
            },
            submenu13:{
                '在校生 - 小学 Students in School - Elementary School': ['在校生 Students in School','小学 Elementary School'],
                '在校生 - 初中 Students in School - Junior High School': ['在校生 Students in School','初中 Junior High School'],
                '在校生 - 高中 Students in School - High School': ['在校生 Students in School','初中 Junior High School'],
                '新入学生 - 大学 Initial Student Enrollment - College/University': ['新入学生 - 大学 Initial Student Enrollment - College/University'],
                '老师- 小学 Teachers - Elementary School': ['老师 Teachers','小学 Elementary School'],
                '老师- 初中 Teachers - Junior High School': ['老师 Teachers','初中 Junior High School'],
                '老师- 高中 Teachers - High School': ['老师 Teachers','初中 Junior High School'],
                '受教育程度 - 文盲 Illiterate': ['受教育程度 Highest Level of Education', '', '文盲 Illiterate', '', '','人数 number of people'],
                '受教育程度 - 小学 Highest Level of Education - Elementary School': ['受教育程度 Highest Level of Education', '', '小学 Elementary School', '', '','人数 number of people'],
                '受教育程度 - 初中 Highest Level of Education - Junior High School': ['受教育程度 Highest Level of Education', '', '初中 Junior High School', '', '','人数 number of people'],
                '受教育程度 - 中专高中 Highest Level of Education - High School': ['受教育程度 Highest Level of Education', '', '中专高中 High School', '', '','人数 number of people'],
                '受教育程度 - 大专以上 Highest Level of Education - College/University or Higher': ['受教育程度 Highest Level of Education', '', '大专以上 College/University or Higher', '', '','人数 number of people'],
            },
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
                console.log(this.fileArr[idx].name.substring(5, this.fileArr[idx].name.indexOf('.')));
                var formNum = parseInt(this.fileArr[idx].name.substring(5, this.fileArr[idx].name.indexOf('.')));
                switch (formNum) {
                    case 9:
                        category = this.cate9;
                        divisions = this.submenu9;
                        console.log(divisions);
                        break;
                    case 10:
                        category = this.cate10;
                        divisions = this.submenu10;
                        break;
                    case 11:
                        category = this.cate11;
                        divisions = this.submenu11;
                        break;
                    case 12:
                        category = this.cate12;
                        break;
                    case 13:
                        category = this.cate13;
                        divisions = this.submenu13;
                        break;
                    case 14:
                        category = this.cate14;
                        break;
                    default:
                        alert('Cannot find this category.');
                        return;
                }
                /* Form a new file name */
                var filename = (this.isAddPrefix)? this.prefixContent+this.fileArr[idx].name : this.fileArr[idx].name;
                this.file2Xce(this.fileArr[idx], filename, category, divisions).then(tabJson => {
                    console.log(tabJson);
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
        file2Xce(file, name, cate, divisions) {
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
                    console.log(cate);
                    // Get year range and the 'Is the data available?'
                    var mode = (name.indexOf('Range') > 0) ? 'range' : 'yearly';
                    var yearRange = {'begin': null, 'end': null};
                    var isAvailableCol, categoryCol, codeCol, titleCol;
                    var bookIdx = {};
                    var divisionColNum = {'category':1, 'division1': 1, 'division2': 2, 'division3': 3, 'division4': 4, 'subdivision':5};
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
                        } else if (cell.w == '村志书名 Gazetteer Title') {
                            titleCol = colNum;
                        }
                    }

                    // Create Header
                    var content = new Array();
                    var header = new Array();
                    for(colNum  = codeCol; colNum <= yearRange.end; colNum++) {
                        header.push(ws[ XLSX.utils.encode_cell({r: 0, c: colNum}) ]);
                    }
                    // var divisionHead = ['category', 'division1', 'divisio2', 'division3','division4', 'subdivision'];
                    // for (var i = 1; i <= divisionHead.length; i++) {
                    //     ws[ XLSX.utils.encode_cell({r: 0, c: yearRange.end+1})].w = divisionHead[i];
                    //     var cell = ws[ XLSX.utils.encode_cell({r: 0, c: yearRange.end+1})];
                    //     cell.w = divisionHead[i];
                    //     header.push(cell);
                    // }
                    content.push(header);

                    // Rename category and get contents of code, title, category
                    var groups = {};
                    for (rowNum = 1; rowNum <= range.e.r; rowNum++) {
                        var cellCode = ws[ XLSX.utils.encode_cell({r: rowNum, c: codeCol})].w;
                        var cellTitle = ws[ XLSX.utils.encode_cell({r: rowNum, c: titleCol})].w;
                        var cellCategory = ws[ XLSX.utils.encode_cell({r: rowNum, c: categoryCol})].w;
                        // Rename the category
                        if (cellCategory == '农转非 (人数) Agricultural to Non-Agricultural Hukou (Change of Residency Status) (number of people)')
                            cellCategory = '农转非 (人数) Agricultural to Non-Agricultural Hukou / Change of Residency Status (number of people)';
                        else if (cellCategory == '耕地面积 (平方米) Cultivated Area (square meters)')
                            cellCategory = '耕地面积 (公顷) Cultivated Area (hectares)';
                        else if (cellCategory == '上环 Use of Interuterine Device (IUD)')
                            cellCategory = '上环 Use of Intrauterine Device (IUD)';
                        else if (cellCategory == '用电量 - 每人 (度) Village Power Consumption - per person (kilowatt hours)')
                            cellCategory = '生活用电量 - 每人 (度) Electricity Consumption - per person (kilowatt hours)';
                        else if (cellCategory == '用电量 - 每户 (度) Village Power Consumption - per household (kilowatt hours)')
                            cellCategory = '生活用电量 - 每户 (度) Electricity Consumption - per household (kilowatt hours)';
                        else if (cellCategory == '用电量 - 全村 (度) Village Power Consumption - Village (kilowatt-hours)')
                            cellCategory = '生活用电量 - 全村 (度) Electricity Consumption - village (kilowatt-hours)';

                        if (groups[cellCode] == null) {
                            groups[cellCode] = new Array();
                        }
                        groups[cellCode].push({'row': rowNum, 'title': cellTitle, 'category': cellCategory});
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
                            var index = tempCate.indexOf(obj['category']);
                            if (index >= 0) {
                                tempCate.splice(index, 1);
                            } else if (index == -1 && cate.indexOf(obj['category']) >= 0 && mode == 'range') {
                                // TOOD: double check here
                                if (ws[ XLSX.utils.encode_cell({r: obj['row'], c: yearRange.begin})] == undefined) {
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
                                groups[code].push({'row': -1, 'title': bookIdx[code], 'category': tempCate[j]});
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
                            rowContent.push(obj['title']);
                            rowContent.push(obj['category']);

                            for(colNum = yearRange.begin; colNum <= yearRange.end; colNum++) {
                                if (obj['row'] > -1) {
                                    var cell = ws[ XLSX.utils.encode_cell({r: obj['row'], c: colNum}) ];
                                    if (cell === undefined) {
                                        rowContent.push('n/a');
                                    } else {
                                        if (!isAvailable) {isAvailable = true;}
                                        rowContent.push(cell.w);
                                    }
                                } else {
                                    rowContent.push('n/a');
                                }
                            }

                            // Division Data
                            if (divisions != null) {
                                for (var i = 0 ; i < divisions[obj['category']].length ; i++) {
                                    rowContent.push(divisions[obj['category']][i]);
                                    console.log(divisions[obj['category']][i]);
                                }
                            }

                            counter++;
                            rowContent.splice(3,0, (isAvailable)?'Yes':'No');
                            content.push(Array.from(rowContent));
                            rowContent.length = 0;
                            isAvailable = false;
                        }
                        if (counter > cate.length)
                            console.log("code: " +code+ ", counter: "+counter);
                    }
                    console.log(groups.length);
                    var result = {'data': content, 'filename': name};
                    resolve(result);
                };
                reader.readAsBinaryString(file);
            });
        },
        loadTextFromFile(ev) {
            const file = ev.target.files[0];
            console.log("before:"+ev.target.files.length);
            var notValidFile = "";
            for (var i = 0 ; i < ev.target.files.length ; i++) {
                let file = ev.target.files[i],
                    types = file.name.split('.').pop(),
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
