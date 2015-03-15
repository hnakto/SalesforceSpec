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
                    .file('xl/worksheets/sheet'+next_id+'.xml',results[2].value);
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
                        var sheet = {
                            path: target_file_path, 
                            value:spreadsheet_this.zip.file('xl/'+target_file_path).asText()
                        };
                        resolve(sheet);
                    }
                );
            }
        );
    });
}

/**
 * * _get_cell_by_name
 * * シート名称,セル名称よりcellデータを取得する
 * @param sheetname
 * @param cell_name
 * @returns {Promise}
 * @private
 */
SpreadSheet.prototype._get_cell_by_name = function(
    sheetname,
    cell_name
){
    var spreadsheet_this = this;
    
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function(c){
        if(/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var row_string = cell_name.substr(index, cell_name.length-index);
    return new Promise(function(resolve, reject){
        spreadsheet_this._get_row_by_name(
            sheetname,
            row_string
        ).then(function(row){
            if(row === undefined)
                resolve(undefined);
            var cell = _.find(row.c, function(c){
                return c['$'].r === cell_name;
            });
            resolve(cell);
        }).catch(function(err){
            reject(err);
        });
    });
    
}

/**
 * * _add_int_value
 * * シートに数値を追加する
 * @param sheetname
 * @param cell_name
 * @param int_vlaue
 * @returns {Promise}
 * @private
 */
SpreadSheet.prototype._add_int_value = function(
    sheetname,
    cell_name,
    int_vlaue
) {
    var spreadsheet_this = this;

    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var row_string = cell_name.substr(index, cell_name.length - index);

    return new Promise(function (resolve, reject) {
        spreadsheet_this._get_sheet_by_name(sheetname)
            .then(function (sheet_xml_obj) {

                parser.parseString(
                    sheet_xml_obj.value,
                    function (err, sheet_xml) {
                        Promise.all([
                            spreadsheet_this._get_row_by_name(
                                sheetname,
                                row_string
                            ),
                            spreadsheet_this._get_cell_by_name(
                                sheetname,
                                cell_name
                            )
                        ]).then(function (results) {
                            var row = results[0];
                            var cell = results[1];

                            if (row === undefined) {
                                var new_row = {
                                    '$': { r: row_string, spans: '1:5' },
                                    c: [{ '$': { r: cell_name }, v: [ int_vlaue ] }]
                                };
                                spreadsheet_this._insert_row(sheet_xml, new_row);
                            }else if(cell === undefined){
                                var cell_value = { '$': { r: cell_name }, v: [ int_vlaue ] };
                                row.c.push(cell_value);
                                spreadsheet_this._update_row(sheet_xml, row);
                            }else{
                                var cell_value = { '$': { r: cell_name }, v: [ int_vlaue ] };
                                spreadsheet_this._update_cell(sheet_xml, row, cell_value);
                            }

                            spreadsheet_this.zip.file('xl/' + sheet_xml_obj.path, builder.buildObject(sheet_xml));
                            resolve(spreadsheet_this.zip);


                        }).catch(function (err) {
                            reject(err);
                        });
                    }
                );

        }).catch(function (err) {
            reject(err);
        });
    });
}

/**
 * * _get_row_by_name
 * * シート名称,行番号よりrowデータを取得する
 * @param sheetname
 * @param row_number
 * @returns {Promise}
 * @private
 */
SpreadSheet.prototype._get_row_by_name = function(
    sheetname,
    row_number
){
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject){
        spreadsheet_this._get_sheet_by_name(sheetname)
            .then(function(sheet_xml_obj) {
                parser.parseString(
                    sheet_xml_obj.value,
                    function (err, sheet_xml) {
                        var row = _.find(sheet_xml.worksheet.sheetData[0].row, function(e){
                            return e['$'].r === row_number;
                        });
                        resolve(row);
                    }
                )
            }).catch(function(err){
                reject(err);
            });
    });
}

/**
 * * _insert_row
 * * sheetインスタンスにrowを追加する
 * @param sheet
 * @param row
 * @private
 */
SpreadSheet.prototype._insert_row = function(
    sheet,
    row
    ) {
    sheet.worksheet.sheetData[0].row.push(row);
}

/**
 * * _update_row
 * * sheetインスタンスのrowプロパティを更新する
 * @param sheet
 * @param row
 * @private
 */
SpreadSheet.prototype._update_row = function(
    sheet,
    row
) {
    row.c = _.sortBy(row.c, function(e){return e['$'].r});
    for(var i = 0; i<sheet.worksheet.sheetData[0].row.length; i++){
        if(sheet.worksheet.sheetData[0].row[i]['$'].r === row['$'].r){
            sheet.worksheet.sheetData[0].row[i] = row;
        }
    }
}

/**
 * * _update_cell
 * * sheetインスタンスのcellプロパティを更新する
 * @param sheet
 * @param row
 * @param cell
 * @private
 */
SpreadSheet.prototype._update_cell = function(
    sheet,
    row,
    cell
) {
    row.c = _.sortBy(row.c, function(e){return e['$'].r});
    for(var i = 0; i<sheet.worksheet.sheetData[0].row.length; i++){
        if(sheet.worksheet.sheetData[0].row[i]['$'].r === row['$'].r){
            for(var j = 0; j < sheet.worksheet.sheetData[0].row[i].c.length; j++){
                if(sheet.worksheet.sheetData[0].row[i].c[j]['$'].r === cell['$'].r){
                    sheet.worksheet.sheetData[0].row[i].c[j] = cell;
                }
            }
        }
    }
}

module.exports = SpreadSheet;

