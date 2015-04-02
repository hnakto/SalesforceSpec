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
var spread_crud = new SpreadSheet('./template/ObjectPermission.xlsx');
var spread_field_permission = new SpreadSheet('./template/FieldPermission.xlsx');
var spread_workflow = new SpreadSheet('./template/WorkFlow.xlsx');
var spread_page_layout = new SpreadSheet('./template/PageLayout.xlsx');

var target_obj = [
    //ここにターゲットを記述
];

Promise.all([
    spec.initialize(),
    spread_custom_field.initialize(),
    spread_validation_rule.initialize(),
    spread_crud.initialize(),
    spread_field_permission.initialize(),
    spread_workflow.initialize(),
    spread_page_layout.initialize()
]).then(function() {
    return Promise.props({
        page_layouts: spec.page_layouts(),
        field_label: spec.field_label(),
        field_type: spec.field_type(),
        object_label: spec.object_label()
    });
}).then(function(result) {
    set_layout_assignment(result);

    spread_custom_field.bulk_copy_sheet('base',target_obj);
    bulk_set_fields(target_obj, spec.metadata.fields);
    set_validation_rules();
    set_profile_crud();
    
    spread_field_permission.bulk_copy_sheet('base', spec.metadata.custom_objs);
    set_all_field_permissions();

    set_workflow();

    return Promise.resolve();
}).then(function(){
    return fileHelper.writeFile(
        "./work/レイアウト一覧.xlsx",
        spread_page_layout.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/入力規則一覧.xlsx",
        spread_validation_rule.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/オブジェクト権限一覧.xlsx",
        spread_crud.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/項目レベル権限一覧.xlsx",
        spread_field_permission.generate()
    );
}).then(function(){
    return fileHelper.writeFile(
        "./work/ワークフロー一覧.xlsx",
        spread_workflow.generate()
    );
}).then(function(){
    log.info('successfully finished.');
}).catch(function(err){
    log.error(err);
});

function set_layout_assignment(result){
    spread_page_layout.bulk_copy_sheet('base', spec.metadata.custom_objs);
    _.each(spec.metadata.custom_objs, function(obj_name){
        var obj = result.page_layouts[obj_name];
        var row_number = 7;
        _.each(Object.keys(obj), function(field_name){
            var field = obj[field_name];
            var insert_row0 = ['', '項目名', '','','','','','型',''];
            var insert_row1 = ['', result.field_label[obj_name+'.'+field_name],'','',field_name, '', '', result.field_type[obj_name+'.'+field_name], '参照可能'];
            var insert_row2 = ['', result.field_label[obj_name+'.'+field_name],'','',field_name, '', '', result.field_type[obj_name+'.'+field_name], '参照のみ'];
            _.each(Object.keys(field), function(layout_name){
                var layout_assignment = field[layout_name];
                insert_row0.push(layout_name);
                if(layout_assignment === 'Readonly'){
                    insert_row1.push('●');
                    insert_row2.push('●');
                }else if(layout_assignment === 'Edit'){
                    insert_row1.push('●');
                    insert_row2.push('');
                }else if(layout_assignment === 'Required'){
                    insert_row1.push('必須');
                    insert_row2.push('');
                }
            });
            /**
            spread_page_layout.add_row(
                obj_name,
                3,
                ['','','','','','','','','','','','','','','','','','',result.object_label[obj_name] + '(' + obj_name + ')']
            );
             */
            spread_page_layout.add_row(
                obj_name,
                6,
                insert_row0
            );
            spread_page_layout.add_row(
                obj_name,
                row_number++,
                insert_row1
            );
            spread_page_layout.add_row(
                obj_name,
                row_number++,
                insert_row2
            );
        });
    });
}


/**
 * * set_profile_crud
 * * CRUD表を作成
 * *  
 */
function set_profile_crud(){
    spread_crud.add_row(
        'base',
        6,
        ['','オブジェクト','','','','','','CRUD'].concat(spec.metadata.valid_profiles)
    );
    var index_on_mark = spread_crud.shared_strings.add_string('●');
    var mark = {'●':index_on_mark};
    
    var profile_crud = spec.metadata.profile_crud;
    for(var i = 0; i<spec.metadata.custom_objs.length; i++){
        var obj_apiname = spec.metadata.custom_objs[i];
        var objname = spec.get_labelname(obj_apiname);

        var permission_c = ['',objname,'','',obj_apiname,'','','作成'];
        var permission_r = ['',objname,'','',obj_apiname,'','','読み取り'];
        var permission_u = ['',objname,'','',obj_apiname,'','','更新'];
        var permission_d = ['',objname,'','',obj_apiname,'','','削除'];
        var permission_all_r = ['',objname,'','',obj_apiname,'','','すべて参照'];
        var permission_all_u = ['',objname,'','',obj_apiname,'','','すべて更新'];

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
        spread_crud.add_row('base',i*6+7,permission_c,mark);
        spread_crud.add_row('base',i*6+8,permission_r,mark);
        spread_crud.add_row('base',i*6+9,permission_u,mark);
        spread_crud.add_row('base',i*6+10,permission_d,mark);
        spread_crud.add_row('base',i*6+11,permission_all_r,mark);
        spread_crud.add_row('base',i*6+12,permission_all_u,mark);
    }
}

/**
 * * set_workflow
 * * ワークフロー一覧
 */
function set_workflow(){
    var row_number = 10;
    _.each(Object.keys(spec.metadata.work_flows), function(objectname){
        var workflow_in_object = spec.metadata.work_flows[objectname];
        _.each(workflow_in_object, function(action){
            spread_workflow.add_row(
                'base',
                row_number++,
                ['',(row_number-10),action.workflow_active,objectname,'','',action.workflow_fullName,'','','',action.workflow_trigger_code,
                    action.workflow_criteria,'','',action.workflow_description,'','','','','','',action.type,'',action.fullName,'','',action.name_or_description,
                    '','','','',action.field_or_recipients,'',action.update_value_or_template]
            );
        });
    })
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
                'base',
                row_number++,
                ['',(row_number-7),rule.active,rule.objectname,'','',rule.fullName,'','','',rule.errorDisplayField,'','',
                    rule.errorConditionFormula,'','','','','','',rule.errorMessage,'','','','','','',rule.description]
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
    /**
    spread_page_layout.add_row(
        sheetname,
        3,
        ['','','','','','','','','','','','','','','','','','',object_label[sheetname] + '(' + sheetname + ')']
    );
     */

    _.each(fields, function(field){
        spread_custom_field.add_row(
            sheetname,
            row_number++,
            ['',(row_number-7),field.label,'','','',field.apiname,'','','',field.type,'',field.formula ? field.formula : field.picklistValues,
                '','','',field.defaultValue,'','',field.required,field.unique,field.externalId,field.trackHistory,field.trackTrending,field.description]
        );
    })
}

/**
 * * set_all_field_permissions
 * * すべてのシートに、Field Permissionを保存する
 */
function set_all_field_permissions(){
    var index_on_mark = spread_field_permission.shared_strings.add_string('●');
    _.each(spec.metadata.custom_objs, function(obj){
        set_field_permissions(
            obj,
            spec.metadata.valid_profiles,
            spec.metadata.field_permission,
            spec.metadata.fields[obj],
            index_on_mark
        );
    });
}

/**
 * * set_field_permissions
 * * 引数のシートに、Field Permissionを保存する
 * @param sheetname
 * @param profiles
 * @param field_permissions
 * @param fields
 */
function set_field_permissions(
    sheetname,
    profiles,
    field_permissions,      //Profile名 × fieldのAPI名(--__c.--__c) → {readable: ''or'●', readonly: ''or'●'}
    fields,
    index_on_mark
){
    var row_number = 7;
    var header_row = ['','オブジェクト','','','','','','CRUD'];
    _.each(profiles, function(profile) {
        header_row.push(profile);
    });
    spread_field_permission.add_row(
        sheetname,
        6,
        header_row
    );
    _.each(fields, function(field){
        var field_row = ['',field.label,'','','',field.apiname,'','',''];

        var row_entry_readble = ['',field.label,'','',field.apiname,'','','参照可能'];;
        var row_entry_readonly = ['',field.label,'','',field.apiname,'','','参照のみ'];;
        _.each(profiles, function(profile) {
            var permissions = field_permissions[profile];
            var field_full_name = sheetname + '.' + field.apiname;
            
            if(permissions[field_full_name]){
                if(permissions[field_full_name].readable !== undefined){
                    row_entry_readble.push(permissions[field_full_name].readable);
                }else{
                    row_entry_readble.push('');
                }
                if(permissions[field_full_name].readonly !== undefined){
                    row_entry_readonly.push(permissions[field_full_name].readonly);
                }else{
                    row_entry_readonly.push('');
                }                
            }else{
                row_entry_readble.push('');
                row_entry_readonly.push('');
            }
        });
        spread_field_permission.add_row(
            sheetname,
            row_number++,
            row_entry_readble,
            {'●': index_on_mark}
        );
        spread_field_permission.add_row(
            sheetname,
            row_number++,
            row_entry_readonly,
            {'●': index_on_mark}
        );
    });
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
        set_fields(object_name,object_fields[object_name]);
    })
}