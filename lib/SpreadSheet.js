/**
 * * SpreadSheet
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/15/2015
 */
var fs = require("fs");
var JSZip = require("jszip");
var xml2js = require('xml2js');
var _ = require("underscore");
var Promise = require('bluebird');

var parser = new xml2js.Parser();
var builder = new xml2js.Builder();

/**
 * * Constructor
 * @param spreadsheet_path
 * @constructor
 */
var SpreadSheet = function(spreadsheet_path) {
    this.spreadsheet_path = spreadsheet_path;
};

/**
 * * initialize
 * * 初期化処理
 * @returns {Promise}
 */
SpreadSheet.prototype.initialize = function(){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject) {
        // read a zip file
        fs.readFile(spreadsheet_this.spreadsheet_path, function (err, data) {
            if (err) {
                reject(err);
            }
            spreadsheet_this.zip = new JSZip(data);
            resolve();
        });
    });
}

/**
 * * copy_sheet
 * * シートのコピー
 * @param src_sheet_name
 * @param dest_sheet_name
 * @returns {Promise}
 */
SpreadSheet.prototype.copy_sheet = function(
    src_sheet_name,
    dest_sheet_name
){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject) {
        spreadsheet_this._available_sheetid(
            function(next_id){
                Promise
                .all([
                    spreadsheet_this._add_sheet_to_workbookxml_rels(next_id),
                    spreadsheet_this._add_sheet_to_workbookxml(next_id, dest_sheet_name),
                    spreadsheet_this._get_sheet_by_name(src_sheet_name)
                ])
                .then(function(results){
                    spreadsheet_this.zip
                    .file("xl/_rels/workbook.xml.rels",results[0])
                    .file("xl/workbook.xml",results[1])
                    .file('xl/worksheets/sheet'+next_id+'.xml',results[2]);
                    resolve(spreadsheet_this.zip);
                })
                .catch(function(err){
                    reject(err);
                });            
            },
            reject
        )
    });
};

/**
 * * bulk_copy_sheet
 * * シートの一括コピー
 * @param src_sheet_name
 * @param dest_sheet_names
 * @returns {Promise}
 */
SpreadSheet.prototype.bulk_copy_sheet = function(
    src_sheet_name,
    dest_sheet_names
){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject){
        dest_sheet_names.reduce(
            function(promise, sheet_name) {
                return promise.then(function(value) {
                    return spreadsheet_this.copy_sheet(src_sheet_name, sheet_name);
                });
            },
            Promise.resolve()
        ).then(
            function(result) {
                resolve(spreadsheet_this.zip);
            }
        ).catch(
            function(err){
                reject(err);
            }
        );
    });
};



/**
 * *  _available_sheetid
 * *  workbook.xml.relsより使用可能なシートIDを取得する
 * @param callback
 * @param error
 * @private
 */
SpreadSheet.prototype._available_sheetid = function(callback, error){
    parser.parseString(
        this.zip.file("xl/_rels/workbook.xml.rels").asText(),
        function (err, workbook_xml_rels) {
            if (err) {
                error(err);
            }
            var max_rel = _.max(workbook_xml_rels.Relationships.Relationship,
                function(e){
                    return Number(e['$'].Id.replace('rId',''));
                }
            );
            var next_id = 'rId' + ('00' + (parseInt((max_rel['$'].Id.replace('rId','')))+parseInt(1))).slice(-3);
            callback(next_id);
        }
    );
}

/**
 * * _add_sheet_to_workbookxml_rels
 * * workbook.xml.relsに新規シートIDを追加する
 * @param next_id
 * @returns {Promise}
 * @private
 */
SpreadSheet.prototype._add_sheet_to_workbookxml_rels = function(
    next_id
){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject){
        parser.parseString(
            spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels").asText(),
            function (err, workbook_xml_rels) {
                if (err) {
                    reject(err);
                }
                workbook_xml_rels.Relationships.Relationship.push(
                    { '$': { Id: next_id,
                            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                            Target: 'worksheets/sheet'+next_id+'.xml'
                           }
                    }
                );
                resolve(builder.buildObject(workbook_xml_rels));
            }
        );
    });
}
/**
 * * _add_sheet_to_workbookxml
 * * workbook.xmlに新規シートを追加する
 * @param next_id
 * @param dest_sheet_name
 * @returns {Promise}
 * @private
 */
SpreadSheet.prototype._add_sheet_to_workbookxml = function(
    next_id,
    dest_sheet_name
) {
    var spreadsheet_this = this; 
    return new Promise(function(resolve, reject){
        parser.parseString(
            spreadsheet_this.zip.file("xl/workbook.xml").asText(),
            function (err, workbook_xml) {
                if (err) {
                    reject(err);
                }
                workbook_xml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId',''), 'r:id': next_id } });
                resolve(builder.buildObject(workbook_xml));
            }
        );
    });
}
/**
 * *  _get_sheet_by_name
 * *  シート名称よりsheet.xmlの値を取得する
 * @param sheetname
 * @private
 */
SpreadSheet.prototype._get_sheet_by_name = function(sheetname){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject){
        parser.parseString(
            spreadsheet_this.zip.file("xl/workbook.xml").asText(),
            function (err, workbook_xml) {
                if (err) {
                    reject(err);
                }
                var sheetid = _.find(workbook_xml.workbook.sheets[0].sheet,
                    function(e){
                        return e['$'].name === sheetname;
                    })['$']['r:id'];
                parser.parseString(
                    spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels").asText(),
                    function (err, workbook_xml_rels) {
                        if (err) {
                            reject(err);
                        }
                        var target_file_path = _.max(workbook_xml_rels.Relationships.Relationship,
                            function(e){
                                return e['$'].Id === sheetid;
                            })['$'].Target;
                        resolve(spreadsheet_this.zip.file('xl/'+target_file_path).asText());
                    }
                );
            }
        );
    });
}

module.exports = SpreadSheet;

