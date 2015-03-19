/**
 * * Parser
 * * XMLファイルのParserクラス
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/19/2015
 */
var Promise = require('bluebird');
var fs = require('fs');
var xml2js = require('xml2js');
var parser = new xml2js.Parser();

/**
 * * Constructor
 * @constructor
 */
var Parser = function() {
}

/**
 * * parse_file
 * * XMLファイルを読み込み、インスタンスを返却する
 * @param filename
 * @returns {Promise}
 */
Parser.prototype.parse_file = function(
    filename
){
    return new Promise(function(resolve, reject){
        fs.readFile(filename, function (err, data) {
            if(err)
                reject(err);
            parser.parseString(data, function (err, result) {
                if(err)
                    reject(err);
                resolve(result);
            });
        });
    });
}


module.exports = Parser;