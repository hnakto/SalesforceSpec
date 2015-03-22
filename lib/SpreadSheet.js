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

var Utility = require('./Util');
var util = new Utility();

var SharedStrings = require('./SharedStrings');

var parser = new xml2js.Parser();
var builder = new xml2js.Builder();

var Log = require('log')
var log = new Log(Log.DEBUG);

var Parser2 = require('./Parser');
var parser2 = new Parser2();

/**
 * * Constructor
 * @param spreadsheet_path
 * @constructor
 */
var SpreadSheet = function(spreadsheet_path) {
    this.spreadsheet_path = spreadsheet_path;
    this.zip = {};
    this.shared_strings = {};
    this.workbookxml_rels = {};
    this.workbookxml = {};
    this.sheet_xmls = [];
};

/**
 * * initialize
 * * 初期化処理
 * @returns {Promise}
 */
SpreadSheet.prototype.initialize = function(){
    log.debug('SpreadSheet:initialize');
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject) {
        fs.readFile(spreadsheet_this.spreadsheet_path, function (err, data) {
            if (err) {
                reject(err);
            }
            spreadsheet_this.zip = new JSZip(data);
            parser2.load_excel(new JSZip(data));
            Promise.all([
                parser2.parse_file_in_excel('xl/sharedStrings.xml')
                ,parser2.parse_file_in_excel('xl/_rels/workbook.xml.rels')
                ,parser2.parse_file_in_excel('xl/workbook.xml')
                ,parser2.parse_dir_in_excel('xl/worksheets')
            ]).then(function(results){
                spreadsheet_this.shared_strings = new SharedStrings(results[0]);
                spreadsheet_this.workbookxml_rels = results[1];
                spreadsheet_this.workbookxml = results[2];
                spreadsheet_this.sheet_xmls = results[3];
                log.debug('SpreadSheet is initialized successfully');
                resolve();
            }).catch(function(err){
                log.error(err);
            })
        });
    });
}

/**
 * * generate
 * * 呼出し元にExcelバイナリデータを返却する
 * @returns {*}
 */
SpreadSheet.prototype.generate = function(){
    log.debug('SpreadSheet:generate');
    this.zip.file('xl/sharedStrings.xml', builder.buildObject(this.shared_strings.obj));
    var zipped_data = this.zip.generate({type:"nodebuffer"});
    return zipped_data;
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
    log.debug('SpreadSheet:copy_sheet from ' + src_sheet_name + ' to ' + dest_sheet_name);
    var spreadsheet_this = this;
    return new Promise(function(resolve, reject) {
        var next_id = spreadsheet_this._available_sheetid();
        spreadsheet_this.workbookxml_rels = spreadsheet_this._add_sheet_to_workbookxml_rels(next_id);
        spreadsheet_this.workbookxml = spreadsheet_this._add_sheet_to_workbookxml(next_id, dest_sheet_name);
        var dest_sheet = {};
        _.extend(dest_sheet, spreadsheet_this._get_sheet_by_name(src_sheet_name).value);
        dest_sheet.name = 'sheet'+next_id+'.xml';
        spreadsheet_this.sheet_xmls.push(dest_sheet);
        spreadsheet_this.zip.file("xl/_rels/workbook.xml.rels",builder.buildObject(spreadsheet_this.workbookxml_rels));
        spreadsheet_this.zip.file("xl/workbook.xml",builder.buildObject(spreadsheet_this.workbookxml));
        var dest_sheet_obj = {};
        dest_sheet_obj.worksheet = {};
        _.extend(dest_sheet_obj.worksheet, dest_sheet.worksheet);
        spreadsheet_this.zip.file('xl/worksheets/sheet'+next_id+'.xml',builder.buildObject(dest_sheet_obj));
        resolve(spreadsheet_this.zip);
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
    log.debug('SpreadSheet:bulk_copy_sheet');
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
SpreadSheet.prototype._available_sheetid = function(){
    var max_rel = _.max(this.workbookxml_rels.Relationships.Relationship,
        function(e){
            return Number(e['$'].Id.replace('rId',''));
        }
    );
    var next_id = 'rId' + ('00' + (parseInt((max_rel['$'].Id.replace('rId','')))+parseInt(1))).slice(-3);
    return next_id;
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
    this.workbookxml_rels.Relationships.Relationship.push(
        { '$': { Id: next_id,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: 'worksheets/sheet'+next_id+'.xml'
               }
        }
    );
    return this.workbookxml_rels;

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
    this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId',''), 'r:id': next_id } });
    return this.workbookxml;
}
/**
 * *  _get_sheet_by_name
 * *  シート名称よりsheet.xmlの値を取得する
 * @param sheetname
 * @private
 */
SpreadSheet.prototype._get_sheet_by_name = function(sheetname){
    var sheetid = _.find(this.workbookxml.workbook.sheets[0].sheet,
        function(e){
            return e['$'].name === sheetname;
        })['$']['r:id'];
    var target_file_path = _.max(this.workbookxml_rels.Relationships.Relationship,
        function(e){
            return e['$'].Id === sheetid;
        })['$'].Target;
    var target_file_name = target_file_path.split('/')[target_file_path.split('/').length-1];
    var sheet_xml = _.find(this.sheet_xmls,
        function(e){
            return e.name === target_file_name;
        });
    var sheet = {
        path: target_file_path,
        value: sheet_xml
    };
    return sheet;
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
    var row = spreadsheet_this._get_row_by_name(
        sheetname,
        row_string
    );
    if(row === undefined)
        return undefined;
    
    var cell = _.find(row.c, function(c){
        return c['$'].r === cell_name;
    });
    return cell;
}

/**
 * * add_row
 * * シートに行を追加する
 * @param sheetname
 * @param row_number
 * @param cell_values
 * @returns {Promise}
 */
SpreadSheet.prototype.add_row = function(
    sheetname,
    row_number,
    cell_values
){
    log.debug('SpreadSheet:' + sheetname + ' add row ' + row_number);
    var key_values = [];
    for(var i = 0; i<cell_values.length; i++){
        var col_string = util.convert_alphabet(i);
        key_values.push({cell_name: (col_string+row_number), value:cell_values[i]});
    }
    return this.bulk_add_value(sheetname, key_values);
}
    
/**
 * * bulk_add_value
 * * 一括でシートに値を追加する
 * @param sheetname
 * @param key_values
 * @returns {Promise}
 */
SpreadSheet.prototype.bulk_add_value = function(
    sheetname,
    key_values
){
    var spreadsheet_this = this;
    var sheet_xml = this._get_sheet_by_name(sheetname);
    _.each(key_values, function(cell){
        spreadsheet_this.add_value(
            sheet_xml,
            sheetname,
            cell.cell_name,
            cell.value
        );
    });
}

/**
 * * add_value
 * * シートに値を追加する
 * @param sheetname
 * @param cell_name
 * @param value
 * @returns {Promise}
 */
SpreadSheet.prototype.add_value = function(
    sheet_xml,
    sheetname,
    cell_name,
    value
) {
    var spreadsheet_this = this;

    var cell_value;
    if(util.is_number(value)){
        cell_value = { '$': { r: cell_name }, v: [ value ] };
    }else{
        var next_index = spreadsheet_this.shared_strings.add_string(value);
        cell_value = { '$': { r: cell_name, t: 's' }, v: [ next_index ] };
    }
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var row_string = cell_name.substr(index, cell_name.length - index);
    var row = spreadsheet_this._get_row_by_name(
        sheetname,
        row_string
    );
    var cell = spreadsheet_this._get_cell_by_name(
        sheetname,
        cell_name
    );
    if (row === undefined) {
        var new_row = {
            '$': { r: row_string, spans: '1:5' },
            c: [cell_value]
        };
        spreadsheet_this._insert_row(sheet_xml, new_row);
    }else if(cell === undefined){
        row.c.push(cell_value);
        spreadsheet_this._update_row(sheet_xml, row);
    }else{
        cell_value['$'].s = cell['$'].s;
        spreadsheet_this._update_cell(sheet_xml, row, cell_value);
    }
    var sheet_obj = {};
    sheet_obj.worksheet = {};
    _.extend(sheet_obj.worksheet, sheet_xml.value.worksheet);
    spreadsheet_this.zip.file('xl/' + sheet_xml.path, builder.buildObject(sheet_obj));
    return;
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
    var sheet_xml = this._get_sheet_by_name(sheetname);
    var row = _.find(sheet_xml.value.worksheet.sheetData[0].row, function(e){
        return e['$'].r === row_number;
    });
    return row;
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
    sheet.value.worksheet.sheetData[0].row.push(row);
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
    for(var i = 0; i<sheet.value.worksheet.sheetData[0].row.length; i++){
        if(sheet.value.worksheet.sheetData[0].row[i]['$'].r === row['$'].r){
            sheet.value.worksheet.sheetData[0].row[i] = row;
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
    for(var i = 0; i<sheet.value.worksheet.sheetData[0].row.length; i++){
        if(sheet.value.worksheet.sheetData[0].row[i]['$'].r === row['$'].r){
            for(var j = 0; j < sheet.value.worksheet.sheetData[0].row[i].c.length; j++){
                if(sheet.value.worksheet.sheetData[0].row[i].c[j]['$'].r === cell['$'].r){
                    sheet.value.worksheet.sheetData[0].row[i].c[j] = cell;
                }
            }
        }
    }
}

module.exports = SpreadSheet;

