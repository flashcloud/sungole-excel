/**
 * Created by sungoleimac2 on 17/9/18.
 */
var path = require('path');
var fs = require('fs');
var process = require('process');
var execSync = require('child_process').execSync;

var node_xj = require('xls-to-json');
var ejsExcel = require("ejsExcel");

var utils = require('../lib/utils.js');

function getTemplateFileSuffix() {
    return '-template';
}

function isTemplateFile(filename) {
    return path.basename(filename).toLowerCase().indexOf(getTemplateFileSuffix()) > 0;
}

function getPrintableFile(xlsTemplateFile) {
    return path.join(path.dirname(xlsTemplateFile), path.basename(xlsTemplateFile, path.extname(xlsTemplateFile)).replace(getTemplateFileSuffix(), '') + "-printable.xlsx");
}

function getPrintResultJsonFile(xlsTemplateFile) {
    return path.join(path.dirname(xlsTemplateFile), path.basename(xlsTemplateFile, path.extname(xlsTemplateFile)) + '-ret.json');
}

function  getDataModJsonFile(xlsTemplateFile) {
    return path.join(path.dirname(xlsTemplateFile), path.basename(xlsTemplateFile, path.extname(xlsTemplateFile)) + '.json');
}

function exportToPrintableXls(dataFiels, xlsTemplateFile, isCheckXlsFile, success, error) {
    //对多个数据文件,按文件名进行排序, 以便在做模板文件时,取数据不会无头绪.
    //例如:对于有主从表的单据.把主表xls文件命名为“1-xx.xls”,明细表命名为“2-xx.xls”
    dataFiels = dataFiels.sort();

    var workTempRootPath = path.dirname(xlsTemplateFile);
    var outputJSONFilePath = path.join(workTempRootPath, 'json');
    var printableXLSPath = path.join(workTempRootPath, 'printableXLS');

    var errmsg = "";

    try {
        //检查文件的合法性
        var invalidFilsMsg = [];
        dataFiels.forEach(function (file, index) {
            var extname = path.extname(file);
            var filename = path.basename(file);
            //console.log('文件类型: ' + getFileMimeType(file.path).mimeType);

            var invalidFile = false;
            if (extname == null)
                invalidFile = true;
            else if (!(extname.toLowerCase() == ".xls" || extname.toLowerCase() == ".xlsx"))
                invalidFile = true;
            if (invalidFile) {
                errmsg = filename + '必须是Excel文件';
                invalidFilsMsg.push(errmsg);
            }
            else {
                if (isCheckXlsFile === true) {
                    //var workbook = xlsx.readFile(file.path); //不能使用此方法检测xls文件的有效性,它会导致应用Crash!!
                    var stdout
                    var cmdstr = 'node "' + path.join(global.appRoot, 'lib', 'utilCommand.js') + '"' + ' isXlsFile -f ' + file;
                    try {
                        stdout = execSync(cmdstr, {encoding: 'utf8'});
                    } catch (e) {
                        errmsg = '执行文件有效性验证时出错. 原因:' + e;
                        invalidFilsMsg.push(errmsg);
                    }
                    process.stdout.write(stdout);
                    if (stdout.indexOf('ok') > -1) {
                        ;
                    } else {
                        errmsg = '无效或上传不正确的' + filename + '文件';
                        invalidFilsMsg.push(errmsg);
                    }
                }
            }
        });
        if (invalidFilsMsg.length > 0) {
            error(invalidFilsMsg.join(', '));
            return false;
        }
    } catch (e) {
        error('解析xls文件失败:' + e);
        return false;
    }
    errmsg = "";

    //如果模板文件是xls,将其转换为xlsx格式
    /*通过下面的方式转换,将原xls转换成的xlsx文件是不一致的
    if (isTemplateFile(xlsTemplateFile) && path.extname(xlsTemplateFile).toLowerCase() == '.xls') {
         var ret = convertTemplateToXlsx();
         if (ret !== true) {
             errmsg = '将模板文件由xls转换成xlsx格式失败: ' + ret;
             console.log(errmsg);
             reset();
             return res.status(400).send(errmsg);
         }
     }*/

    //2. 将此原始XLS文件转成JSON存为JSON文件,存放到JSON输出文件夹
    var hasConverted = 0 ;
    var jsonDataFiles = [];
    var errs = [];
    var sourceJSonData =  {};
    var extractJson = function(err, itemsJSONData, outputFilePath) {
        hasConverted++;
        sourceJSonData[path.basename(outputFilePath, path.extname(outputFilePath))] = itemsJSONData;
        if (err)
            errs.push('转换JSON失败: ' + err);

        if (hasConverted == dataFiels.length) {
            //检查是否有错误发生
            if (errs.length > 0) {
                error(errs.join(','));
                return;
            }
            //加载转换的 JSON 文件,写入到最后的 XLSX 文件中
            var jsonDatas = [];
            var dataModObj= JSON.parse(fs.readFileSync(getDataModJsonFile(xlsTemplateFile)));
            var pagedRows = dataModObj.pagedRows;
            var reqBlankRow = dataModObj.reqBlankRow == 1 ? true : false;
            var dataSources = dataModObj.ds;
            for (var ii = 0; ii < dataSources.length; ii++) {
                var dataSource = dataSources[ii];
                var masterDSName = dataSource.name;
                if (masterDSName !== "" ) {
                    var masterDataFile = dataFiels.filter(function (el) { return path.basename(el) == masterDSName})[0];
                    var detailDSs = dataSource.det;
                    var hasDetailData = (!(detailDSs.length == 1 && detailDSs[0].name === ""));
                    //提取主数据
                    var dataXlsFilePrefix = path.basename(masterDataFile, path.extname(masterDataFile));
                    var masterItems = sourceJSonData[dataXlsFilePrefix];
                    //JSON 文件是已经在上面转换并保存到目录了.
                    if (masterItems.length > 0) {
                        //读取子从数据
                        if (hasDetailData) {
                            var emptyDetailData = true;
                            //如果有子从数据
                            for (var jj = 0; jj < masterItems.length; jj++) {
                                var masterItem = masterItems[jj];
                                //在每一条主数据记录中添加一属性“dets”,将其从属的从数据添加在此属性中
                                masterItem.dets = [];
                                for (var kk = 0; kk < detailDSs.length; kk++) {
                                    var detailDS = detailDSs[kk];
                                    var detailDSName = detailDS.name;
                                    var detailDataFile = dataFiels.filter(function (el) {
                                        return path.basename(el) == detailDSName
                                    })[0];
                                    var detailDataXlsFilePrefix = path.basename(detailDataFile, path.extname(detailDataFile));
                                    var detailItems = sourceJSonData[detailDataXlsFilePrefix];
                                    if (detailItems.length > 0) {
                                        emptyDetailData = false;
                                        var finalDetailItems = pagingData(reqBlankRow, pagedRows, detailItems);
                                        masterItem.dets.push(finalDetailItems);
                                        //请不要更改此处的数据生成格式,否则所有的EXCEL模板的数据获取方式全部要更新
                                        jsonDatas.push(masterItem);
                                    }
                                }
                            }
                            //请不要更改此处的数据生成格式,否则所有的EXCEL模板的数据获取方式全部要更新
                            if (emptyDetailData) {
                                if (jsonDatas.length == 0)
                                    jsonDatas = masterItems;
                                else
                                    masterItems.forEach(function(item) {
                                        jsonDatas.push(item);
                                    });
                            }
                        } else {
                            //请不要更改此处的数据生成格式,否则所有的EXCEL模板的数据获取方式全部要更新
                            if (jsonDatas.length == 0)
                                jsonDatas.push(masterItems);
                            else
                                masterItems.forEach(function(item) {
                                    jsonDatas.push(item);
                                });
                        }
                    }
                }
            }

            //3. 在转换完所有的数据EXCEL文件后,将JSON数据再转换成可打印格式的XLS文件
            translateToPrintableXls(xlsTemplateFile, jsonDatas, success, error);
        }
    };

    dataFiels.forEach(function (file, index) {
        var sourceXlsFilePrefix = path.basename(file, path.extname(file));
        var outputJsonFile = path.join(workTempRootPath, sourceXlsFilePrefix + ".json");
        console.log(outputJsonFile);
        jsonDataFiles.push(outputJsonFile);
        try {
            node_xj({
                input: file,  // input xls
                output: outputJsonFile
            }, extractJson);
        } catch (e) {
            console.log(e);
            error('将文件 ' + file.name + ' 转换成JSON失败: ' + e);
            return;
        }
    });
}

//根据 JSON 和模板,生成最张的可打印 XLSX 文件
function translateToPrintableXls(xlsTemplateFile, jsonDatas, success, error) {
    try {
        var xlsTemplateFileObj = fs.readFileSync(xlsTemplateFile);
        ejsExcel
            .renderExcel(xlsTemplateFileObj, jsonDatas)
            .then(function (printableXLSBuf) {
                var printableXLSFile = getPrintableFile(xlsTemplateFile);
                fs.writeFileSync(printableXLSFile, printableXLSBuf);
                //fs.writeFileSync(path.join(path.dirname(xlsTemplateFile), 'alldata-debug.json'), JSON.stringify(jsonDatas, null, 4)); //测试用,打印出所有的 Json数据
                success(printableXLSFile, "生成可打印的xlsx文件成功: " + printableXLSFile);
            })
            .catch(function (e) {
                fs.writeFileSync(path.join(path.dirname(xlsTemplateFile), 'alldata-debug.json'), JSON.stringify(jsonDatas, null, 4)); //测试用,打印出所有的 Json数据
                error('生成可打印的xlsx文件失败:' + e);
                return;
            });
    } catch (e) {
        fs.writeFileSync(path.join(path.dirname(xlsTemplateFile), 'alldata-debug.json'), JSON.stringify(jsonDatas, null, 4)); //测试用,打印出所有的 Json数据
        error('生成xlsx文件失败:' + e);
        return;
    }
}

function pagingData(reqBlankRow, pagedRows, detailItems) {
    var reqPageing = false
    //需要分页
    if (pagedRows == 0) {
        pagedRows = detailItems.length;
    } else
        reqPageing = true;
    var finalDetailItems = utils.splitArray(detailItems, pagedRows);

    //在分页的情况下,是否需要补充空白行
    if (reqPageing && reqBlankRow) {
        var lastPageData = finalDetailItems[finalDetailItems.length - 1];
        if (lastPageData.length < pagedRows) {
            var lastPageLastRow = lastPageData[lastPageData.length - 1];
            var blankRowsLen = pagedRows - lastPageData.length;
            for (var k = 1; k <= blankRowsLen; k++) {
                var blankRow = utils.deepCopy(lastPageLastRow);
                for (var key in blankRow)
                    blankRow[key] = "";
                lastPageData.push(blankRow);
            }
        }
    }

    return finalDetailItems;
}

/**
 * 如果模板文件是xls,将其转换为xlsx格式
 * @returns {*}
 * deprecate: 转换后的xlsx与原文件有出入,并不一致,要命的是,用ejsexcel生成报表时,会报:too many length or distance symbols
 */
function convertTemplateToXlsx(xlsTemplateFilePath) {
    try {
        var templateWorkbook = xlsx.readFileSync(xlsTemplateFilePath);
        var newTemplateFile = path.join(path.dirname(xlsTemplateFilePath), path.basename(xlsTemplateFilePath, path.extname(xlsTemplateFilePath)) + '.xlsx');
        xlsx.writeFileSync(templateWorkbook, newTemplateFile, {file: 'xlsx'});
        fs.unlinkSync(xlsTemplateFilePath);
        xlsTemplateFilePath = newTemplateFile;
        return true;
    } catch (e) {
        return e
    }
}

//获取文件真实类型
function getFileMimeType(filePath) {
    var buffer = new Buffer(8);
    var fd = fs.openSync(filePath, 'r');
    fs.readSync(fd, buffer, 0, 8, 0);
    var newBuf = buffer.slice(0, 4);
    var head_1 = newBuf[0].toString(16);
    var head_2 = newBuf[1].toString(16);
    var head_3 = newBuf[2].toString(16);
    var head_4 = newBuf[3].toString(16);
    var typeCode = head_1 + head_2 + head_3 + head_4;
    var filetype = '';
    var mimetype;
    switch (typeCode) {
        case 'ffd8ffe1':
            filetype = 'jpg';
            mimetype = ['image/jpeg', 'image/pjpeg'];
            break;
        case '47494638':
            filetype = 'gif';
            mimetype = 'image/gif';
            break;
        case '89504e47':
            filetype = 'png';
            mimetype = ['image/png', 'image/x-png'];
            break;
        case '504b34':
            filetype = 'zip';
            mimetype = ['application/x-zip', 'application/zip', 'application/x-zip-compressed'];
            break;
        case '2f2aae5':
            filetype = 'js';
            mimetype = 'application/x-javascript';
            break;
        case '2f2ae585':
            filetype = 'css';
            mimetype = 'text/css';
            break;
        case '5b7bda':
            filetype = 'json';
            mimetype = ['application/json', 'text/json'];
            break;
        case '3c212d2d':
            filetype = 'ejs';
            mimetype = 'text/html';
            break;
        case 'd0cf11e0':
            filetype = 'xls';
            mimetype = 'application/vnd.ms-excel';
            break;
        default:
            filetype = 'unknown';
            break;
    }
    fs.closeSync(fd);
    return {
        fileType: filetype,
        mimeType: mimetype
    }
}

module.exports = {
    exportToPrintableXls: exportToPrintableXls,
    getPrintableFile: getPrintableFile,
    getPrintResultJsonFile: getPrintResultJsonFile,
    getTemplateFileSuffix: getTemplateFileSuffix,
    isTemplateFile: isTemplateFile
};