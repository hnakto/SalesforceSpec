/**
 * * SharedStrings
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/15/2015
 */

/**
 * * Constructor
 * @param obj
 * @constructor
 */
var SharedStrings = function(obj) {
    this.obj = obj;
    this.count = parseInt(obj.sst.si.length)-parseInt(1);
};

/**
 * * add_string
 * * SharedStrings.xmlに変数を追加する
 * * 呼出し元に追加した変数のインデックスを返却する
 * @param value
 * @returns {*}
 */
SharedStrings.prototype.add_string = function(value){
    var new_string = { t: [ value ], phoneticPr: [ { '$': { fontId: '1' } } ] };
    this.obj.sst.si.push(new_string);
    return this._update_count();
};

/**
 * * get_string_count
 * * 呼出し元にカウントを返却する 
 * @returns {number|*}
 */
SharedStrings.prototype.get_string_count = function(){
    return this.count;
}

/**
 * * _update_count
 * * カウントを更新する 
 * @returns {number}
 * @private
 */
SharedStrings.prototype._update_count = function(){
    this.count = parseInt(this.obj.sst.si.length)-parseInt(1);
    return this.count;
};

module.exports = SharedStrings;
