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
    this.zip = {};
}

/**
 * * load_excel
 * * Excelデータをメンバ変数にセットする
 * @param excel_zipped_data
 */
Parser.prototype.load_excel = function(
    excel_zipped_data
    ){
    this.zip = excel_zipped_data;
}

/**
 * * parse_file_in_excel
 * * Excelファイルの中のファイルを読み込み/パースして返却する
 * @param file_path
 * @returns {Promise}
 */
Parser.prototype.parse_file_in_excel = function(
    file_path
    ){
    var parser_this = this;
    return new Promise(function(resolve, reject) {
        parser.parseString(
            parser_this.zip.file(file_path).asText(),
            function (err, result) {
                if(err)
                    reject(err);
                resolve(result);
            }
        );
    });
}

/**
 * * parse_dir_in_excel
 * * Excelファイルの中のディレクトリ内のファイルを読み込み/パースして返却する
 * @param dir
 * @returns {Promise}
 */
Parser.prototype.parse_dir_in_excel = function(
    dir
    ){
    var parser_this = this;
    var files = parser_this.zip.folder(dir).file(/.xml/);
    return new Promise(function(resolve, reject){
        var file_xmls = [];
        files.reduce(
            function(promise, file) {
                return promise.then(function(prior_file) {
                    if(prior_file)
                        file_xmls.push(prior_file);
                    return Promise.resolve()
                        .then(function(){
                            return parser_this.parse_file_in_excel(file.name);
                        }).then(function(file_xml){
                            file_xml.name = file.name.split('/')[file.name.split('/').length-1];
                            return Promise.resolve(file_xml);
                        });
                });
            },
            Promise.resolve()
        ).then(
            function(last_file) {
                if(last_file)
                    file_xmls.push(last_file);
                resolve(file_xmls);
            }
        ).catch(
            function(err){
                reject(err);
            }
        );
    });
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