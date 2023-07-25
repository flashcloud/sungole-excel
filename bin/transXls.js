'use strict';

var http = require('http');
var path = require('path');
var fs = require('fs');

var xlsTranslator = {
    convert: function (name, datafile, temploatefile, outputfile) {
        var files = [
            {urlKey: name, urlValue: datafile},
            {urlkey: name + "template", urlValue: temploatefile}
        ];
        var options = {
            host: "192.168.31.5",
            port: "8732",
            method: "POST",
            path: "/xlsToPrint"
        }

        var req = http.request(options, function (res) {
            console.log('STATUS: ' + res.statusCode);
            console.log('HEADERS: ' + JSON.stringify(res.headers));
            console.log('  ');

            var downloadData = '';
            res.setEncoding('binary');
            //res.setEncoding("utf8");

            res.on("data", function (chunk) {
                downloadData += chunk;
            })

            res.on("end", function () {
                var foptions = "binary";
                if (res.statusCode != 200) {
                    outputfile = path.normalize(path.dirname(outputfile) + path.sep + path.basename(outputfile, path.extname(outputfile)) + '-err' + '.txt');
                    //foptions = 'utf-8'; //加了此代码,反而是乱码
                }
                fs.writeFile(outputfile, downloadData, foptions, function(err) {
                    if (err)
                        console.log("WRITE FILE ERROR: " + err);
                    else {
                        //fs.close(outputfile);
                        console.log("WRITE OUTPUTFILE TO:" + outputfile);
                    }
                });
            })

        })

        req.on('error', function (e) {
            console.log('problem with request:' + e.message);
            console.log(e);
        })
        postFile(files, req);
        console.log("post file done");
    },
};

function postFile(fileKeyValue, req) {
    var boundaryKey = Math.random().toString(16);
    var enddata = '\r\n----' + boundaryKey + '--';

    var files = new Array();
    for (var i = 0; i < fileKeyValue.length; i++) {
        var content = "\r\n----" + boundaryKey + "\r\n" + "Content-Type: application/octet-stream\r\n" + "Content-Disposition: form-data; name=\"" + fileKeyValue[i].urlKey + "\"; filename=\"" + path.basename(fileKeyValue[i].urlValue) + "\"\r\n" + "Content-Transfer-Encoding: binary\r\n\r\n";
        var contentBinary = new Buffer(content, 'utf-8');//当编码为ascii时，中文会乱码。
        files.push({contentBinary: contentBinary, filePath: fileKeyValue[i].urlValue});
    }
    var contentLength = 0;
    for (var i = 0; i < files.length; i++) {
        var stat = fs.statSync(files[i].filePath);
        contentLength += files[i].contentBinary.length;
        contentLength += stat.size;
    }

    req.setHeader('Content-Type', 'multipart/form-data; boundary=--' + boundaryKey);
    req.setHeader('Content-Length', contentLength + Buffer.byteLength(enddata));

    // 将参数发出
    var fileindex = 0;
    var doOneFile = function () {
        req.write(files[fileindex].contentBinary);
        var fileStream = fs.createReadStream(files[fileindex].filePath, {bufferSize: 4 * 1024});
        fileStream.pipe(req, {end: false});
        fileStream.on('end', function () {
            fileindex++;
            if (fileindex == files.length) {
                req.end(enddata);
            } else {
                doOneFile();
            }
        });
    };
    if (fileindex == files.length) {
        req.end(enddata);
    } else {
        doOneFile();
    }
}

module.exports = xlsTranslator;