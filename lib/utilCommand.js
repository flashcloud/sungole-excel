#! /usr/bin/env node

var process = require('process');
var path = require('path');
var fs = require('fs');

var program = require('commander');
var xlsx = require('xlsx');

var xlsToPrinter = require('../lib/xlsToPrinter');
var utils = require('../lib/utils.js');

function list(val) {
    return val.split(',');
}

program
    .version('0.0.1')
    .usage('isXlsFile -f c:/demo.xls')
    .option('-f, --file [value]', 'xls(xlsx) file', '')
    .option('-l, --list [items]', 'xls data files', list)
    .option('-t, --template [value]', 'printable of template xlsx file', '')

program
    .command('isXlsFile')
    .description('file is valid of xls(xlsx) type. usage: ' +
        '1. enter app root: npm link; ' +
        '2. input command: grsoftsvrutil isXlsFile -f c:/demo.xls' +
        'if valid type, console output "yes", else output "no"')
    .action(function () {
        isXlsFile(program.file);
    });

program
    .command('xlsToPrint')
    .description('file is valid of xls(xlsx) type. usage: ' +
        '1. enter app root: npm link; ' +
        '2. input command: grsoftsvrutil isXlsFile -f c:/demo.xls' +
        'if valid type, console output "yes", else output "no"')
    .action(function () {
        xlsToPrint(program.list, program.template);
    });

program.parse(process.argv);

/**
 * 用命令行的方式检查指定的文件file是否是合法的xls文件
 * @param file
 */
function isXlsFile(file) {
    if (!fs.existsSync(file)) return;

    try {
        var workbook = xlsx.readFileSync(file);
        workbook = null;
        process.stdout.write('ok');
    } catch (e) {
        //无效或不是XL文件,不作任何运作;
        process.stdout.write('no');
    }
}

/**
 * 使用命令行的方式生成可打印的xlsx文件
 * @param workPath xls数据文件及模板文件所在路径
 * @param xlsDataFiles
 * @param xlsTemplateFile
 * @return 控制台输出
 */
function xlsToPrint(xlsDataFiles, xlsTemplateFile) {
    var retjson = {flag:false, file: '', err: '', msg: ''};
    var retjsonfile = xlsToPrinter.getPrintResultJsonFile(xlsTemplateFile);
    xlsToPrinter.exportToPrintableXls(xlsDataFiles, xlsTemplateFile, false,
        function (printableXLSFile, msg) {
            //导出成功
            retjson.flag = true;
            retjson.printablefile = printableXLSFile;
            retjson.msg = msg;
            var data = JSON.stringify(retjson);

            //console.log('start sleep......')
            //utils.sleep(10000);
            //console.log('end sleep......')

            fs.writeFileSync(retjsonfile, data);
        }, function (err) {
            //导出失败后
            retjson.flag = false;
            retjson.err = err;
            var data = JSON.stringify(retjson);
            fs.writeFileSync(retjsonfile, data);
        }
    );
}