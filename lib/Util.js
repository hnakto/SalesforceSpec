/**
 * * utility functions
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/15/2015
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

module.exports = Utility;

