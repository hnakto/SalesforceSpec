/**
 * * utility class
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/17/2015
 */

var _ = require('underscore');

/**
 *  * Utility
 * @constructor
 */
var Utility = function() {
}
/**
 * * is_number
 * * 引数が数値がどうかを判定する。
 * * '12345'等の文字列として定義されている数値も数値とする。
 * @param value
 * @returns {*}
 */
Utility.prototype.is_number = function(value){
    if( typeof(value) != 'number' && typeof(value) != 'string' )
        return false;
    else
        return (value == parseFloat(value) && isFinite(value));
}

/**
 * * convert_alphabet
 * * 列番号をExcelの列アルファベットに変換する
 * @param value
 * @returns {string}
 */
Utility.prototype.convert_alphabet = function(value){
    var number1 = Math.floor(value/(26*26));
    var number2 = Math.floor((value-number1*26*26)/26);
    var number3 = value-(number1*26*26+number2*26);
    
    var alphabet1 = this._convert(number1) === 'A' ? '' : this._convert(number1 - 1);
    var alphabet2 = (alphabet1 === '' && this._convert(number2) === 'A') ? '' : this._convert(number2 - 1);
    var alphabet3 = this._convert(number3);
    
    var alphabet = alphabet1 + alphabet2 + alphabet3;
    return alphabet;
};

/**
 * * _convert
 * * アルファベットの配列を取得する
 * @param value
 * @returns {*}
 * @private
 */
Utility.prototype._convert = function(value){
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')[value];
};

/**
 * * revert_number
 * * Excel列形式の文字列を数値に変換する
 * @param alphabet
 * @returns {number}
 */
Utility.prototype.revert_number = function(alphabet){
    var alphabet_with_zero = ('00'+alphabet).slice(-3).split('');
    var value = 0;
    if(alphabet_with_zero[0] !== '0')
        value = value + this._revert(alphabet_with_zero[0])*26*26;
    if(alphabet_with_zero[1] !== '0')
        value = value + this._revert(alphabet_with_zero[1])*26;
    if(alphabet_with_zero[2] !== '0')
        value = value + this._revert(alphabet_with_zero[2]);

    return value;
};

/**
 * * _revert
 * * A-Zのアルファベットを数値に変換する
 * @param alphabet
 * @returns {number}
 * @private
 */
Utility.prototype._revert = function(alphabet){
    return alphabet.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
};

/**
 * * get_row_string
 * * Excelセル形式の文字列から行番号を取得する
 * @param cell_name
 * @returns {string}
 */
Utility.prototype.get_row_string = function(cell_name){
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var row_string = cell_name.substr(index, cell_name.length - index);
    return row_string;
};

/**
 * * get_col_string
 * * Excelセル形式の文字列から列文字列を取得する
 * @param cell_name
 * @returns {string}
 */
Utility.prototype.get_col_string = function(cell_name){
    var cell_name_array = cell_name.split('');
    var index = 0;
    _.each(cell_name_array, function (c) {
        if (/^[a-zA-Z()]+$/.test(c))
            index++;
    });
    var col_string = cell_name.substr(0, index);
    return col_string;
};

/**
 * * convert_salesforce_type
 * * Salesforce Metadata APIのtype名称から型の日本語名を取得する
 * @param type
 * @returns {string}
 */
Utility.prototype.convert_salesforce_type = function(type){
    switch (type) {
        case 'Email':
            return 'メール';
        case 'Html':
            return 'テキストエリア (リッチ)';
        case 'LongTextArea':
            return 'ロングテキストエリア';
        case 'Number':
            return '数値';
        case 'Picklist':
            return '選択リスト';
        case 'MultiselectPicklist':
            return '選択リスト (複数選択)';
        case 'Location':
            return '地理位置情報';
        case 'Currency':
            return '通貨';
        case 'Phone':
            return '電話';
        case 'Date':
            return '日付';
        case 'AutoNumber':
            return '自動採番号';
        case 'DateTime':
            return '日付/時間';
        case 'Lookup':
            return '参照関係';
        case 'MasterDetail':
            return '主従関係';
        case 'Checkbox':
            return 'チェックボックス';
        case 'Text':
            return 'テキスト';
        case 'EncryptedText':
            return 'テキスト(暗号化)';
        case 'TextArea':
            return 'テキストエリア';
        case 'Percent':
            return 'パーセント'
        case 'Url':
            return 'URL';
    }
}

/**
 * * desc
 * * RetrieveしたBooleanフィールドの整形処理
 * @param flag
 * @returns {string}
 */
Utility.prototype.desc = function(flag){
    if(flag === undefined
        || flag === null
        || flag[0] === undefined
        || flag[0] === null
        || flag[0] === false
        || flag[0] === 'false'){
        return '';
    }else{
        return '●';
    }
}

module.exports = Utility;

