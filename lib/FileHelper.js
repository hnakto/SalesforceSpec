/**
 * * FileHelper
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/24/2015
 */

var Promise = require('bluebird');
var fs = require('fs');
var _ = require('underscore');
var Log = require('log')
var log = new Log(Log.DEBUG);


/**
 * * Constructor
 * @constructor
 */
var FileHelper = function() {
    
}

/**
 * * writeFile
 * * ファイルへ書き込み、Promiseを返却する
 * @param filepath
 * @param data
 * @returns {Promise}
 */
FileHelper.prototype.writeFile = function(
    filepath,    
    data
){
    return new Promise(function(resolve, reject){
        fs.writeFile(
            filepath,
            data,
            function(error) {
                if(error){
                    reject(error);
                }
                log.info(filepath + 'is created successfully');
                resolve();
            }
        );
    });
}

module.exports = FileHelper;