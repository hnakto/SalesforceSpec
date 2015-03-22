/**
 * * SalesforceSpec
 * * Metadata APIを使用して、Salesforce組織から定義情報を抽出してExcelファイルに出力する。
 * * <<出力情報>>
 * *  - 項目一覧
 * *  - ページレイアウト一覧,プロファイル×レコードタイプ×ページレイアウトのマッピング
 * *  - 項目のページレイアウト配置状況一覧
 * *  - レコードタイプ一覧,及びレコードタイプ毎の選択リストの選択値
 * *  - 入力規則一覧
 * *  - ワークフロー,アクション(項目自動更新,メールアラート)一覧
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/17/2015
 */

//require
var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var SalesforceSpec = require('./lib/SalesforceSpec');
var spec = new SalesforceSpec();
var SpreadSheet = require('./lib/SpreadSheet');
var spread = new SpreadSheet('./template/SpecTemplate.xlsx');
var Utility = require('./lib/Util');
var util = new Utility();

//Salesforce
var sf_objs = [];   //カスタムオブジェクト
var sf_rules = [];     //入力規則

Promise.all([
    spec.initialize(),
    spread.initialize()
]).then(function() {
    spread.bulk_copy_sheet('field',spec.metadata.custom_objs);
    set_fields('Opportunity__c', spec.metadata.fields['Opportunity__c']);
    var zip = spread.generate();
    return new Promise(function(resolve, reject){
        fs.writeFile(
            "./work/Specification.xlsx",
            zip,
            function(error) {
                if(error){
                    reject(error);
                }
                resolve();
            }
        );
    });
}).then(function(){
    console.log('Successfully ended');
}).catch(function(err){
    console.log(err);
});


/**
 * * set_fields
 * * 1シートに値をセットする
 * @param sheetname
 * @param fields
 */
function set_fields(
    sheetname, 
    fields
){
    var row_number = 7;
    _.each(fields, function(field){
        spread.add_row(
            sheetname,
            row_number++,
            ['',(row_number-7),field.label,'',field.apiname,'',field.type,'',field.formula ? field.formula : field.picklistValues,
                '','',field.description,'','','',field.required,field.unique,field.externalId]
        );
    })
}

/**
 * * bulk_set_fields
 * * 複数のシートに値をセットする
 * @param sheetnames
 * @param object_fields
 */
function bulk_set_fields(
    sheetnames,
    object_fields
) {
    _.each(sheetnames, function(object_name){
        set_fields(object_name, object_fields[object_name]);
    })
}