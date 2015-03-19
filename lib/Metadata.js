/**
 * * Metadata
 * * Salesforce組織のメタデータを格納するDtoクラス
 * *  
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/08/2015
 */

/**
 * * Constructor
 * @constructor
 */
var Metadata = function() {
    //入力規則
    this.validation_rules = {};
    //レコードタイプ
    this.recordtypes = {};
    //オブジェクト × レコードタイプ一覧 × 選択リスト値一覧
    this.recordtype_picklist = {};
    //プロファイル × レコードタイプ × ページレイアウト
    this.layout_assignment = {};
    //フィールド
    this.fields = {};
    //オブジェクト一覧
    this.objs = [];
    //カスタムオブジェクト一覧
    this.custom_objs = [];
}

module.exports = Metadata;

