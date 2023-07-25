#! /usr/bin/env node

var program = require('commander'),
    XlsTranstor = require('../bin/transXls.js');

program
    .version('0.0.1')
    .usage('transxls YL -d c:/YL.xls -t c:/YL-template.xlsx -o c:/downloadXls/YL.xlsx')
    .option('-d, --datafile [value]', 'xls data file', '')
    .option('-t, --temploatefile [value]', 'xls template file', '')
    .option('-o, --outputfile [value]', 'output printable file', '')

program
    .command('transxls <name>')
    .description('Convert normal xls to printable xls. usage: ' +
        '1. 进入本应用根目录,然后执行: npm link; ' +
        '2. 输入命令: grsoftutil transxls YL -d c:/YL.xls -t c:/YL-template.xlsx -o c:/downloadXls/YL.xlsx')
    .action(function (name) {
        console.log(' datafile - %s ', program.datafile);
        console.log(' temploatefile - %s ', program.temploatefile);
        console.log(' outputfile - %s ', program.outputfile);
        XlsTranstor.convert(name, program.datafile, program.temploatefile, program.outputfile);
    });

program.parse(process.argv);
