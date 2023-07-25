/**
 * Created by sungoleimac2 on 18/1/13.
 */

var fs = require('fs');
var path = require('path');

function delFolder(folderPath) {
    emptyDir(folderPath);
    rmEmptyDir(folderPath);
    fs.unlinkSync(folderPath);
}

var emptyDir = function (fileUrl) {
    var files = fs.readdirSync(fileUrl);//读取该文件夹
    files.forEach(function (file) {
        var stats = fs.statSync(fileUrl + '/' + file);
        if (stats.isDirectory()) {
            emptyDir(fileUrl + '/' + file);
        } else {
            fs.unlinkSync(fileUrl + '/' + file);
        }
    });
};

//删除所有的空文件夹
var rmEmptyDir = function (fileUrl) {
    var files = fs.readdirSync(fileUrl);
    if (files.length > 0) {
        var tempFile = 0;
        files.forEach(function (fileName) {
            tempFile++;
            rmEmptyDir(fileUrl + '/' + fileName);
        });
        if (tempFile == files.length) {//删除母文件夹下的所有字空文件夹后，将母文件夹也删除
            fs.rmdirSync(fileUrl);
        }
    } else {
        fs.rmdirSync(fileUrl);
    }
};

// Return a list of files of the specified fileTypes in the provided dir,
// with the file path relative to the given dir
// dir: path of the directory you want to search the files for
// fileTypes: array of file types you are search files, ex: ['.txt', '.jpg']
//eg:
//print the txt files in the current directory
//getFilesFromDir("./", [".txt"]).map(console.log);
function getFilesFromDir(dir, fileTypes) {
    var filesToReturn = [];

    function walkDir(currentPath) {
        var files = fs.readdirSync(currentPath);
        for (var i in files) {
            var curFile = path.join(currentPath, files[i]);
            if (fs.statSync(curFile).isFile() && fileTypes.indexOf(path.extname(curFile)) != -1) {
                filesToReturn.push(curFile.replace(dir, ''));
            } else if (fs.statSync(curFile).isDirectory()) {
                walkDir(curFile);
            }
        }
    };
    walkDir(dir);
    return filesToReturn;
}

//分割数组
//len: 每个数组的长度
function splitArray(arr, len) {
    var a_len = arr.length;
    var result = [];
    for (var i = 0; i < a_len; i += len) {
        result.push(arr.slice(i, i + len));
    }
    return result;
}

function sleep(delay) {
    var start = new Date().getTime();
    while (new Date().getTime() < start + delay);
}

function deepCopy(p, c) {
    var c = c || {};
    for (var i in p) {
        if (typeof p[i] === 'object') {
            c[i] = (p[i].constructor === Array) ? [] : {};
            deepCopy(p[i], c[i]);
        } else {
            c[i] = p[i];
        }
    }
    return c;
}

module.exports = {
    delFolder: delFolder,
    getFilesFromDir: getFilesFromDir,
    splitArray: splitArray,
    sleep: sleep,
    deepCopy: deepCopy
};