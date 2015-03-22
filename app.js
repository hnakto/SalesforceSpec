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
var Utility = require('./lib/Util');
var util = new Utility();

var spread_custom_field = new SpreadSheet('./template/CustomField.xlsx');
var spread_validation_rule = new SpreadSheet('./template/Validation.xlsx');

Promise.all([
    spec.initialize(),
    spread_custom_field.initialize(),
    spread_validation_rule.initialize()
]).then(function() {
    //カスタム項目一覧
    spread_custom_field.bulk_copy_sheet('field',spec.metadata.custom_objs);
    set_fields('Opportunity__c', spec.metadata.fields['Opportunity__c']);
    //入力規則
    set_validation_rules('Validation', spec.metadata.validation_rules);

    //項目定義書を出力
    var zip = spread_custom_field.generate();
    return new Promise(function(resolve, reject){
        fs.writeFile(
            "./work/項目定義一覧.xlsx",
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
    console.log('項目定義一覧 is created successfully');
    var zip = spread_validation_rule.generate();
    return new Promise(function(resolve, reject){
        fs.writeFile(
            "./work/入力規則一覧.xlsx",
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
    console.log('入力規則一覧 is created successfully');
}).catch(function(err){
    console.log(err);
});


/***
 * * set_validation_rules
 * * (入力規則)
 * @param sheetname
 * @param rules
 */
function set_validation_rules(
    sheetname,
    rules
){
    var row_number = 7;
    _.each(Object.keys(rules), function(objectname){
        var rules_in_object = rules[objectname];
        _.each(rules_in_object, function(rule){
            spread_validation_rule.add_row(
                sheetname,
                row_number++,
                ['',(row_number-7),rule.active,objectname,'',rule.fullName,'',rule.errorMessage,
                    '','','',rule.errorConditionFormula]
            );
        });
    })

    
}    
/**
 * * set_fields
 * * (項目定義書)1シートに値をセットする
 * @param sheetname
 * @param fields
 */
function set_fields(
    sheetname, 
    fields
){
    var row_number = 7;
    _.each(fields, function(field){
        spread_custom_field.add_row(
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