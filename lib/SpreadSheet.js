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
 * * シートをコピーする
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
                var promise1 = new Promise(function(resolve1, reject1){
                    parser.parseString(
                        spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels").asText(),
                        function (err, workbook_xml_rels) {
                            if (err) {
                                reject1(err);
                            }

                            workbook_xml_rels.Relationships.Relationship.push(
                                { '$': { Id: next_id,
                                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                                    Target: 'worksheets/sheet'+next_id+'.xml'
                                }
                                }
                            );
                            resolve1(builder.buildObject(workbook_xml_rels));
                        }
                    );
                });

                var promise2 = new Promise(function(resolve2, reject2){
                    parser.parseString(
                        spreadsheet_this.zip.file("xl/workbook.xml").asText(),
                        function (err, workbook_xml) {
                            if (err) {
                                reject2(err);
                            }

                            workbook_xml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: '3', 'r:id': next_id } });
                            resolve2(builder.buildObject(workbook_xml));
                        }
                    );
                });

                var promise3 = new Promise(function(resolve3, reject3){
                    spreadsheet_this._get_sheet_by_name(
                        src_sheet_name,
                        resolve3,
                        reject3
                    );
                });
                
                Promise
                    .all([promise1,promise2,promise3])
                    .then(function(results){
                        var workbook_xml_rels = results[0];
                        var workbook_xml = results[1];
                        var src_sheet_string = results[2];

                        spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels",workbook_xml_rels);
                        spreadsheet_this.zip.file("xl/workbook.xml",workbook_xml);
                        spreadsheet_this.zip.file('xl/worksheets/sheet'+next_id+'.xml',src_sheet_string);

                        resolve(spreadsheet_this.zip);
                    })
                    .catch(function(err){
                        reject(err);
                    });            
            },
            reject
        )
        
    });
}

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
 * *  _get_sheet_by_name
 * *  シート名称よりsheet.xmlの値を取得する
 * @param sheetname
 * @param callback
 * @param error
 * @private
 */
SpreadSheet.prototype._get_sheet_by_name = function(sheetname, callback, error){

    var spreadsheet_this = this;
    parser.parseString(
        spreadsheet_this.zip.file("xl/workbook.xml").asText(),
        function (err, workbook_xml) {
            if (err) {
                error(err);
            }

            var sheetid = _.find(workbook_xml.workbook.sheets[0].sheet,
                function(e){
                    return e['$'].name === sheetname;
                })['$']['r:id'];

            parser.parseString(
                spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels").asText(),
                function (err, workbook_xml_rels) {
                    if (err) {
                        error(err);
                    }
                    
                    var target_file_path = _.max(workbook_xml_rels.Relationships.Relationship,
                        function(e){
                            return e['$'].Id === sheetid;
                        })['$'].Target;
                    
                    callback(spreadsheet_this.zip.file('xl/'+target_file_path).asText());
                }
            );

            
        }
    );

}

module.exports = SpreadSheet;

