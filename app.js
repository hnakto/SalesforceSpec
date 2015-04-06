/**
 * * SalesforceSpec
 * * Metadata APIを使用して、Salesforce組織から定義情報を抽出してExcelファイルに出力する。
 * * <<出力情報>>
 * *  - オブジェクト一覧, 項目一覧, オブジェクト権限一覧, レイアウト一覧
 * *  - ワークフロー一覧, 入力規則一覧,項目レベル権限一覧
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/17/2015
 */

//require
var _ = require('underscore');
var Promise = require('bluebird');
var fs = Promise.promisifyAll(require("fs"));
var Log = require('log');
var log = new Log(Log.DEBUG);
var yaml = require('js-yaml');
var moment = require('moment');

var SpreadSheet = require('./lib/SpreadSheet');
var SalesforceSpec = require('./lib/SalesforceSpec');
var spec = new SalesforceSpec();

var config = load_config();

var spread_custom_field = new SpreadSheet(config.template_directory + config.template_file.custom_field);
var spread_validation_rule = new SpreadSheet(config.template_directory + config.template_file.validation);
var spread_crud = new SpreadSheet(config.template_directory + config.template_file.object_permission);
var spread_field_permission = new SpreadSheet(config.template_directory + config.template_file.field_permission);
var spread_workflow = new SpreadSheet(config.template_directory + config.template_file.workflow);
var spread_page_layout = new SpreadSheet(config.template_directory + config.template_file.page_layout);

Promise.all([
    spec.initialize(),
    spread_custom_field.initialize(),
    spread_validation_rule.initialize(),
    spread_crud.initialize(),
    spread_field_permission.initialize(),
    spread_workflow.initialize(),
    spread_page_layout.initialize()
]).then(function() {
    return Promise.all([
        build_page_layout(),
        build_validation(),
        build_workflow(),
        build_field_permission(),
        build_object_permission(),
        build_custom_field()
    ]);
}).then(function(){
    log.info('successfully finished.');
}).catch(function(err){
    log.error(err);
});

function build_page_layout(){
    var documentname = config.output_file.page_layout.replace('.xlsx','');
    set_summary_header(spread_page_layout, documentname, '-');
    spread_page_layout.set_row('summary',6,['','No','オブジェクト名','','','','','レイアウト名','','','','','備考']);
    var row_number = 7;
    _.each(Object.keys(spec.page_layout_list()), function(object_name){
        var layouts = spec.page_layout_list()[object_name];
        _.each(layouts, function(layout){
            spread_page_layout.set_row('summary',row_number,['',(row_number++ - 6),object_name,'','','','',layout]);
        });
    });
    set_header(spread_page_layout, documentname, '-');
    spread_page_layout.bulk_copy_sheet('base', spec.object_names);
    _.each(spec.object_names, function(obj_name){
        var obj = spec.page_layouts()[obj_name];
        row_number = 7;
        _.each(Object.keys(obj), function(field_name){
            var field = obj[field_name];
            var insert_row0 = ['', '項目名', '','','','','','型',''];
            var layouts = Object.keys(field);
            layouts = _.map(layouts, function(e){return e.replace(obj_name+'-','')});
            spread_page_layout.set_row(obj_name,3,insert_row0.concat(layouts));
            spread_page_layout.set_row(obj_name,3,
                [config.system_name,'','','','','','','','',documentname,'','','','','','','','',
                    obj_name,'','','','','','','','',moment().format("YYYY/MM/DD"),'','','','']);

            spread_page_layout.set_row(obj_name,6,insert_row0.concat(layouts));
            
            var insert_row1 = ['', spec.field_label()[obj_name+'.'+field_name],'','',field_name, '', '', spec.field_type()[obj_name+'.'+field_name], '参照可能'];
            var insert_row2 = ['', spec.field_label()[obj_name+'.'+field_name],'','',field_name, '', '', spec.field_type()[obj_name+'.'+field_name], '参照のみ'];
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
            
            spread_page_layout.set_row(obj_name,row_number++,insert_row1);
            spread_page_layout.set_row(obj_name,row_number++,insert_row2);
        });
    });
    log.info('page_layout is created successfully');
    return fs.writeFileAsync(config.output_directory + config.output_file.page_layout,spread_page_layout.generate());
}

function build_object_permission(){
    var profiles = spec.valid_profile();
    set_header(spread_crud, config.output_file.object_permission.replace('.xlsx',''), '-');
    spread_crud.set_row('base',6,['','オブジェクト','','','','','','CRUD'].concat(profiles));

    var index_on_mark = spread_crud.shared_strings.add_string('●');
    var mark = {'●':index_on_mark};
    var object_permission = spec.object_permission();
    for(var i = 0; i<spec.object_names.length; i++){
        var obj_apiname = spec.object_names[i];
        var objname = spec.label_name(obj_apiname);
        var permission_c = ['',objname,'','',obj_apiname,'','','作成'];
        var permission_r = ['',objname,'','',obj_apiname,'','','読み取り'];
        var permission_u = ['',objname,'','',obj_apiname,'','','更新'];
        var permission_d = ['',objname,'','',obj_apiname,'','','削除'];
        var permission_all_r = ['',objname,'','',obj_apiname,'','','すべて参照'];
        var permission_all_u = ['',objname,'','',obj_apiname,'','','すべて更新'];
        for(var j = 0; j<profiles.length; j++){
            var profile_name = profiles[j];
            var permission = (object_permission[profile_name])[obj_apiname];
            permission_c.push(permission? permission.allowCreate : '');
            permission_r.push(permission? permission.allowRead : '');
            permission_u.push(permission? permission.allowEdit : '');
            permission_d.push(permission? permission.allowDelete : '');
            permission_all_r.push(permission? permission.viewAllRecords : '');
            permission_all_u.push(permission? permission.modifyAllRecords : '');
        }
        spread_crud.set_row('base',i*6+7,permission_c,mark);
        spread_crud.set_row('base',i*6+8,permission_r,mark);
        spread_crud.set_row('base',i*6+9,permission_u,mark);
        spread_crud.set_row('base',i*6+10,permission_d,mark);
        spread_crud.set_row('base',i*6+11,permission_all_r,mark);
        spread_crud.set_row('base',i*6+12,permission_all_u,mark);
    }
    log.info('object_permission is created successfully');
    return fs.writeFileAsync(config.output_directory + config.output_file.object_permission, spread_crud.generate());
}

function build_workflow(){
    var row_number = 10;
    set_header(spread_workflow, config.output_file.workflow.replace('.xlsx',''), '-');
    spread_workflow.set_row('base',6,['','','評価条件']);
    spread_workflow.set_row('base',7,['','','','1 : 作成されたとき   2 : 作成されたとき、および編集されるたび   3 : 作成されたとき、およびその後基準を満たすように編集されたとき']);
    spread_workflow.set_row('base',9,['','No','有効','オブジェクト','','','ワークフロー名','','','','評価条件','条件','','','説明','','','','','','','アクション種別','','アクション名','','','表示名 / 説明','','','','','対象項目 / 受信対象','','更新値 / 送信テンプレート']);
    
    _.each(Object.keys(spec.workflow()), function(objectname){
        var workflow_in_object = spec.workflow()[objectname];
        _.each(workflow_in_object, function(action){
            spread_workflow.set_row(
                'base',
                row_number++,
                ['',(row_number-10),action.workflow_active,objectname,'','',action.workflow_fullName,'','','',action.workflow_trigger_code,
                    action.workflow_criteria,'','',action.workflow_description,'','','','','','',action.type,'',action.fullName,'','',action.name_or_description,
                    '','','','',action.field_or_recipients,'',action.update_value_or_template]
            );
        });
    });
    log.info('workflow is created successfully');
    return fs.writeFileAsync(config.output_directory + config.output_file.workflow,spread_workflow.generate());
}

function build_validation(){
    var row_number = 7;
    set_header(spread_validation_rule, config.output_file.validation.replace('.xlsx',''), '-');
    spread_validation_rule.set_row(
        'base',6,
        ['','No.','有効','オブジェクト','','','入力規則名','','','','エラー表示場所','','',
            '評価条件 / 数式','','','','','','','エラーメッセージ','','','','','','','説明']
    );
    var validation_rules = spec.validation_rule();
    _.each(Object.keys(validation_rules), function(objectname){
        var rules_in_object = validation_rules[objectname];
        _.each(rules_in_object, function(rule){
            spread_validation_rule.set_row(
                'base',
                row_number++,
                ['',(row_number-7),rule.active,rule.objectname,'','',rule.fullName,'','','',rule.errorDisplayField,'','',
                    rule.errorConditionFormula,'','','','','','',rule.errorMessage,'','','','','','',rule.description]
            );
        });
    });
    return fs.writeFileAsync(config.output_directory + config.output_file.validation,spread_validation_rule.generate());
}

function build_custom_field(){
    var documentname = config.output_file.custom_field.replace('.xlsx','');
    set_header(spread_custom_field, documentname, '-');
    spread_custom_field.set_row('base',6,
        ['','No','表示ラベル','','','','API参照名','','','','型','','数式/選択リスト値','','','',
         '初期値','','','必須','ユニーク','外部','履歴','トレンド','説明']);
    spread_custom_field.bulk_copy_sheet('base', spec.object_names);
    var all_custom_fields = spec.custom_field();
    _.each(spec.object_names, function(obj_name){
        var row_number = 7;
        spread_custom_field.set_row(obj_name,3,
            [config.system_name,'','','','','','','','',documentname,'','','','','','','','',
               obj_name,'','','','','','','','',moment().format("YYYY/MM/DD"),'','','','']);

        var custom_fields = all_custom_fields[obj_name];
        _.each(custom_fields, function(field){
            spread_custom_field.set_row(
                obj_name,
                row_number++,
                ['',(row_number-7),field.label,'','','',field.apiname,'','','',field.type,'',field.formula ? field.formula : field.picklistValues,
                    '','','',field.defaultValue,'','',field.required,field.unique,field.externalId,field.trackHistory,field.trackTrending,field.description]
            );
        });
    });
    return fs.writeFileAsync(config.output_directory + config.output_file.custom_field, spread_custom_field.generate());
}

function build_field_permission(){
    var documentname = config.output_file.field_permission.replace('.xlsx','');
    set_header(spread_field_permission, documentname, '-');

    spread_field_permission.bulk_copy_sheet('base', spec.object_names);
    var index_on_mark = spread_field_permission.shared_strings.add_string('●');
    _.each(spec.object_names, function(sheetname){
        spread_field_permission.set_row(sheetname,3,
            [config.system_name,'','','','','','','','',documentname,'','','','','','','','',
                sheetname,'','','','','','','','',moment().format("YYYY/MM/DD"),'','','','']);

        var profiles = spec.valid_profile();
        var field_permissions = spec.field_permission();
        var fields = spec.custom_field()[sheetname];
        var row_number = 7;
        var header_row = ['','オブジェクト','','','','','','CRUD'];
        _.each(profiles, function(profile) {
            header_row.push(profile);
        });
        spread_field_permission.set_row(sheetname,6,header_row);
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
            spread_field_permission.set_row(sheetname, row_number++, row_entry_readble, {'●': index_on_mark});
            spread_field_permission.set_row(sheetname, row_number++, row_entry_readonly,{'●': index_on_mark});
        });
    });
    return fs.writeFileAsync(config.output_directory + config.output_file.field_permission,spread_field_permission.generate());
}

function set_header(spreadsheet, documentname, document_target){
    spreadsheet.set_row('base',1,
        ['システム名','','','','','','','','','ドキュメント名','','','','','','','','',    
         'ドキュメント対象','','','','','','','','','作成日','','','','最終更新日']);    
    spreadsheet.set_row('base',2,    
        ['システム名','','','','','','','','','ドキュメント名','','','','','','','','',    
         'ドキュメント対象','','','','','','','','','作成者','','','','最終更新者']);
    spreadsheet.set_row('base',3,
        [config.system_name,'','','','','','','','',documentname,'','','','','','','','',
         document_target,'','','','','','','','',moment().format("YYYY/MM/DD"),'','','','']);
    spreadsheet.set_row('base',4,
        [config.system_name,'','','','','','','','',documentname,'','','','','','','','',
         document_target,'','','','','','','','',config.created_by,'','','','']);
}

function set_summary_header(spreadsheet, documentname, document_target){
    spreadsheet.set_row('summary',1,
        ['システム名','','','','','','ドキュメント名','','','','','',
            'ドキュメント対象','','','','','','作成日','','','最終更新日']);
    spreadsheet.set_row('summary',2,
        ['システム名','','','','','','ドキュメント名','','','','','',
            'ドキュメント対象','','','','','','作成者','','','最終更新者']);
    spreadsheet.set_row('summary',3,
        [config.system_name,'','','','','',documentname,'','','','','',
            document_target,'','','','','',moment().format("YYYY/MM/DD"),'','','']);        
    spreadsheet.set_row('summary',4,
        [config.system_name,'','','','','',documentname,'','','','','',
            document_target,'','','','','',config.created_by,'','','']);
}
function load_config(){
    var config = yaml.safeLoad(fs.readFileSync('./yaml/config.yml', 'utf8'));
    return config;
}