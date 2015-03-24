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
var FileHelper = require('./lib/FileHelper');
var fileHelper = new FileHelper();

var Log = require('log')
var log = new Log(Log.DEBUG);


var spread_custom_field = new SpreadSheet('./template/CustomField.xlsx');
var spread_validation_rule = new SpreadSheet('./template/Validation.xlsx');
var spread_crud = new SpreadSheet('./template/CRUD.xlsx');

Promise.all([
    spec.initialize(),
    spread_custom_field.initialize(),
    spread_validation_rule.initialize(),
    spread_crud.initialize()
]).then(function() {
    
    set_fields('Opportunity__c', spec.metadata.fields['Opportunity__c']);
    set_validation_rules();
    set_profile_crud();

    return Promise.resolve();
}).then(function(){
    return fileHelper.writeFile(
        "./work/項目定義書.xlsx",
        spread_custom_field.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/入力規則一覧.xlsx",
        spread_validation_rule.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/プロファイル権限一覧.xlsx",
        spread_crud.generate()
    );
}).then(function(){
    log.info('successfully finished.');
}).catch(function(err){
    log.error(err);
});

/**
 * * set_profile_crud
 * * CRUD表を作成
 * *  
 */
function set_profile_crud(){
    spread_crud.add_row(
        'CRUD',
        6,
        ['','オブジェクト','','','CRUD'].concat(spec.metadata.valid_profiles)
    );
    var profile_crud = spec.metadata.profile_crud;
    for(var i = 0; i<spec.metadata.custom_objs.length; i++){
        var obj_apiname = spec.metadata.custom_objs[i];
        var objname = spec.get_labelname(obj_apiname);

        var permission_c = ['',objname,'','','作成'];
        var permission_r = ['',objname,'','','読み取り'];
        var permission_u = ['',objname,'','','更新'];
        var permission_d = ['',objname,'','','削除'];
        var permission_all_r = ['',objname,'','','すべて参照'];
        var permission_all_u = ['',objname,'','','すべて更新'];

        for(var j = 0; j<spec.metadata.valid_profiles.length; j++){
            var profile_name = spec.metadata.valid_profiles[j];
            var permission = (profile_crud[profile_name])[obj_apiname];
            permission_c.push(permission? permission.allowCreate : '');
            permission_r.push(permission? permission.allowRead : '');
            permission_u.push(permission? permission.allowEdit : '');
            permission_d.push(permission? permission.allowDelete : '');
            permission_all_r.push(permission? permission.viewAllRecords : '');
            permission_all_u.push(permission? permission.modifyAllRecords : '');
        }
        spread_crud.add_row('CRUD',i*6+7,permission_c);
        spread_crud.add_row('CRUD',i*6+8,permission_r);
        spread_crud.add_row('CRUD',i*6+9,permission_u);
        spread_crud.add_row('CRUD',i*6+10,permission_d);
        spread_crud.add_row('CRUD',i*6+11,permission_all_r);
        spread_crud.add_row('CRUD',i*6+12,permission_all_u);
    }
}



/***
 * * set_validation_rules
 * * (入力規則)
 */
function set_validation_rules(){
    var row_number = 7;
    _.each(Object.keys(spec.metadata.validation_rules), function(objectname){
        var rules_in_object = spec.metadata.validation_rules[objectname];
        _.each(rules_in_object, function(rule){
            spread_validation_rule.add_row(
                'Validation',
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
    spread_custom_field.bulk_copy_sheet('field',spec.metadata.custom_objs);
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