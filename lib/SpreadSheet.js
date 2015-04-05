/**
 * * SpreadSheet
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/15/2015
 */
var Promise = require('bluebird');
var xml2js = require('xml2js');
var fs = Promise.promisifyAll(require("fs"));
var parseString = Promise.promisify(xml2js.parseString);
var JSZip = require("jszip");
var _ = require("underscore");
var moment = require('moment');
var builder = new xml2js.Builder();
var Log = require('log');
var log = new Log(Log.DEBUG);

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
    log.debug('SpreadSheet(' + this.spreadsheet_path + '):initialize');
    var spreadsheet_this = this;
    return fs.readFileAsync(spreadsheet_this.spreadsheet_path)
        .then(function(data){
            spreadsheet_this.zip = new JSZip(data);
            return Promise.props({
                shared_strings: parseString(spreadsheet_this.zip.file('xl/sharedStrings.xml').asText()),
                workbookxml_rels: parseString(spreadsheet_this.zip.file('xl/_rels/workbook.xml.rels').asText()),
                workbookxml: parseString(spreadsheet_this.zip.file('xl/workbook.xml').asText()),
                sheet_xmls :spreadsheet_this._parse_dir_in_excel('xl/worksheets')
            }).then(function(template_obj){
                spreadsheet_this.shared_strings.initialize(template_obj.shared_strings);
                spreadsheet_this.workbookxml_rels = template_obj.workbookxml_rels;
                spreadsheet_this.workbookxml = template_obj.workbookxml;
                spreadsheet_this.sheet_xmls = template_obj.sheet_xmls;
                log.debug('SpreadSheet(' + spreadsheet_this.spreadsheet_path + ') is initialized successfully');
                spreadsheet_this.set_template();
            });
        });
};

SpreadSheet.prototype.generate = function(){
    log.info('SpreadSheet(' + this.spreadsheet_path + '):generate');
    this.zip
    .file("xl/_rels/workbook.xml.rels",builder.buildObject(this.workbookxml_rels))
    .file("xl/workbook.xml",builder.buildObject(this.workbookxml))
    .file('xl/sharedStrings.xml', builder.buildObject(this.shared_strings.get_obj()));
    var spreadsheet_this = this;
    _.each(this.sheet_xmls, function(sheet){
       if(sheet.name){
           var sheet_obj = {};
           sheet_obj.worksheet = {};
           _.extend(sheet_obj.worksheet, sheet.worksheet);
           spreadsheet_this.zip.file('xl/worksheets/'+sheet.name, builder.buildObject(sheet_obj));
       }
    });
    return this.zip.generate({type:"nodebuffer"});
};

/** テンプレート設定処理 */
SpreadSheet.prototype.set_template = function(sheetname){
    this.set_row(
        'base',
        3,
        [process.env.OUTPUT_SYSTEM_NAME,'','','','','','','','','','','','','','','','','','','','','','','','','','',moment().format("YYYY/MM/DD")]
    );
    this.set_row(
        'base',
        4,
        ['','','','','','','','','','','','','','','','','','','','','','','','','','','',process.env.OUTPUT_CREATED_BY]
    );

};

SpreadSheet.prototype.copy_sheet = function(
    src_sheet_name,
    dest_sheet_name
){
    log.debug('SpreadSheet(' + this.spreadsheet_path + '):copy_sheet from ' + src_sheet_name + ' to ' + dest_sheet_name);
    var next_id = this.available_sheetid();
    this.workbookxml_rels.Relationships.Relationship.push(
        { '$': { Id: next_id,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            Target: 'worksheets/sheet'+next_id+'.xml'
        }
        }
    );
    this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId',''), 'r:id': next_id } });

    var src_sheet = this.sheet_by_name(src_sheet_name).value;
    var copied_sheet = JSON.parse(JSON.stringify(src_sheet));
    copied_sheet.name = 'sheet'+next_id+'.xml';
    this.sheet_xmls.push(copied_sheet);
};

SpreadSheet.prototype.bulk_copy_sheet = function(
    src_sheet_name,
    dest_sheet_names
){
    var spreadsheet_this = this;
    _.each(dest_sheet_names, function(dest_sheet_name){
        spreadsheet_this.copy_sheet(src_sheet_name, dest_sheet_name);
    });
};

SpreadSheet.prototype.available_sheetid = function(){
    var max_rel = _.max(this.workbookxml_rels.Relationships.Relationship,
        function(e){
            return Number(e['$'].Id.replace('rId',''));
        }
    );
    var next_id = 'rId' + ('00' + (parseInt((max_rel['$'].Id.replace('rId','')))+parseInt(1))).slice(-3);
    return next_id;
};

SpreadSheet.prototype.sheet_by_name = function(sheetname){
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

SpreadSheet.prototype.cell_by_name = function(
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
    var row = spreadsheet_this.row_by_name(sheetname,row_string);
    if(row === undefined)
        return undefined;
    var cell = _.find(row.c, function(c){
        return c['$'].r === cell_name;
    });
    return cell;
};

SpreadSheet.prototype.set_row = function(
    sheetname,
    row_number,
    cell_values,
    existing_setting
){
    var key_values = [];
    for(var i = 0; i<cell_values.length; i++){
        var col_string = _convert_alphabet(i);
        key_values.push({cell_name: (col_string+row_number), value:cell_values[i]});
    }
    return this.bulk_set_value(sheetname, key_values,existing_setting);
}
    
SpreadSheet.prototype.bulk_set_value = function(
    sheetname,
    key_values,
    existing_setting
){
    var spreadsheet_this = this;
    var sheet_xml = this.sheet_by_name(sheetname);
    _.each(key_values, function(cell){
        spreadsheet_this.set_value(
            sheet_xml,
            sheetname,
            cell.cell_name,
            cell.value,
            existing_setting
        );
    });
}

SpreadSheet.prototype.set_value = function(
    sheet_xml,
    sheetname,
    cell_name,
    value,
    existing_setting
) {
    if(value === '')
        return;

    var spreadsheet_this = this;

    var cell_value;
    if(_is_number(value)){
        cell_value = { '$': { r: cell_name }, v: [ value ] };
    }else{
        var next_index;
        if(existing_setting && existing_setting[value]){
            next_index = existing_setting[value];
        }else{
            next_index = spreadsheet_this.shared_strings.add_string(value);
        }
        cell_value = { '$': { r: cell_name, t: 's' }, v: [ next_index ] };
    }
    var row_string = _get_row_string(cell_name);
    var row = spreadsheet_this.row_by_name(sheetname,row_string);
    var cell = spreadsheet_this.cell_by_name(sheetname, cell_name);
    if (row === undefined) {
        var new_row = {
            '$': { r: row_string, spans: '1:5' },
            c: [cell_value]
        };
        sheet_xml.value.worksheet.sheetData[0].row.push(new_row);
    }else if(cell === undefined){
        row.c.push(cell_value);
        spreadsheet_this._update_row(sheet_xml, row);
    }else{
        cell_value['$'].s = cell['$'].s;
        if(cell_value['$'].t){
            cell['$'].t = 's';
            cell.v = [ next_index ];
        }else{
            cell.v = [ value ];
            if(cell['$'].t) delete cell['$'].t;
        }
        spreadsheet_this._update_cell(sheet_xml, row, cell);
    }
    return;
};

SpreadSheet.prototype.row_by_name = function(
    sheetname,
    row_number
){
    var sheet_xml = this.sheet_by_name(sheetname);
    var row = _.find(sheet_xml.value.worksheet.sheetData[0].row, function(e){
        return e['$'].r === row_number;
    });
    return row;
};

SpreadSheet.prototype._update_row = function(
    sheet,
    row
) {
    row.c = _.sortBy(row.c, function(e){return _revert_number(_col_string(e['$'].r))});
    _.each(sheet.value.worksheet.sheetData[0].row, function(existing_row){
        if(existing_row['$'].r == row['$'].r)
            existing_row = row;
    });
};

SpreadSheet.prototype._update_cell = function(
    sheet,
    row,
    cell
) {
    row.c = _.sortBy(row.c, function(e){return _revert_number(_col_string(e['$'].r))});
    _.each(sheet.value.worksheet.sheetData[0].row, function(existing_row){
        if(existing_row['$'].r == row['$'].r){
            _.each(existing_row.c, function(existing_cell){
                if(existing_cell['$'].r === cell['$'].r)
                    existing_cell = cell;
            });     
        }
    });
};

SpreadSheet.prototype._parse_dir_in_excel = function(
    dir
){
    var spreadsheet_this = this;
    var files = spreadsheet_this.zip.folder(dir).file(/.xml/);
    var file_xmls = [];
    return files.reduce(
            function(promise, file) {
                return promise.then(function(prior_file) {
                    return Promise.resolve()
                        .then(function(){
                            return parseString(spreadsheet_this.zip.file(file.name).asText());
                        }).then(function(file_xml){
                            file_xml.name = file.name.split('/')[file.name.split('/').length-1];
                            file_xmls.push(file_xml);
                            return file_xmls;
                        });
                });
            },
            Promise.resolve()
        );
};

SpreadSheet.prototype.shared_strings = (function() {
    var obj = {};
    var count = 0;
    return {
        initialize: function(obj) {
            this.obj = obj;
            this.count = parseInt(obj.sst.si.length)-parseInt(1);
        },
        get_obj: function(){
            return this.obj;
        },
        add_string: function(value) {
            var new_string = { t: [ value ], phoneticPr: [
                { '$': { fontId: '1' } }
            ] };
            this.obj.sst.si.push(new_string);
            this.count = parseInt(this.obj.sst.si.length) - parseInt(1);
            return this.count;
        }
    }
})();


function _is_number(value){
    if( typeof(value) != 'number' && typeof(value) != 'string' )
        return false;
    else
        return (value == parseFloat(value) && isFinite(value));
}

function _convert(value){
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')[value];
}

function _convert_alphabet(value){
    var number1 = Math.floor(value/(26*26));
    var number2 = Math.floor((value-number1*26*26)/26);
    var number3 = value-(number1*26*26+number2*26);

    var alphabet1 = _convert(number1) === 'A' ? '' : _convert(number1 - 1);
    var alphabet2 = (alphabet1 === '' && _convert(number2) === 'A') ? '' : _convert(number2 - 1);
    var alphabet3 = _convert(number3);

    var alphabet = alphabet1 + alphabet2 + alphabet3;
    return alphabet;
}

function _revert(alphabet){
    return alphabet.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

function _revert_number(alphabet){
    var alphabet_with_zero = ('00'+alphabet).slice(-3).split('');
    var value = 0;
    if(alphabet_with_zero[0] !== '0')
        value = value + _revert(alphabet_with_zero[0])*26*26;
    if(alphabet_with_zero[1] !== '0')
        value = value + _revert(alphabet_with_zero[1])*26;
    if(alphabet_with_zero[2] !== '0')
        value = value + _revert(alphabet_with_zero[2]);
    return value;
}

function _col_string(cell_name){
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var col_string = cell_name.substr(0, index);
    return col_string;
}

function _get_row_string(cell_name){
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var row_string = cell_name.substr(index, cell_name.length - index);
    return row_string;
}

function _get_col_string(cell_name){
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var col_string = cell_name.substr(0, index);
    return col_string;
};

module.exports = SpreadSheet;

