/**
 * * utility class
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/17/2015
 */

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
    
    var alphabet1 = this._convert(number1) === 'A' ? '' : this._convert(number1);
    var alphabet2 = (alphabet1 === '' && this._convert(number2) === 'A') ? '' : this._convert(number2);
    var alphabet3 = this._convert(number3);
    
    var alphabet = alphabet1 + alphabet2 + alphabet3;
    return alphabet;
}

/**
 * * _convert
 * *
 * @param value
 * @returns {*}
 * @private
 */
Utility.prototype._convert = function(value){
    return "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split('')[value];
}

module.exports = Utility;

