/**
 * * SalesforceSpec
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/08/2015
 */

var Promise = require('bluebird');
var tmp = require('tmp');
var jsforce = require('jsforce');
var fs = Promise.promisifyAll(require("fs"));
var _ = require('underscore');
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var parser = new xml2js.Parser();
var walk = require('walk');
var AdmZip = require('adm-zip');
var Metadata = require('./Metadata');
var Parser2 = require('./Parser');
var parser2 = new Parser2();
var Util = require('./Util');
var util = new Util();

var Log = require('log')
var log = new Log(Log.DEBUG);

require('dotenv').load();

/**
 * * Constructor
 * @constructor
 */
var SalesforceSpec = function() {
    this.salesforce_username = process.env.SALESFORCE_USERNAME;
    this.salesforce_password = process.env.SALESFORCE_PASSWORD;
    this.salesforce_host = process.env.SALESFORCE_HOST;
    this.salesforce_host = 'https://' + this.salesforce_host;
    this.temporary_directory = process.env.TEMPORARY_DIRECTORY;
    this.config_file = process.env.CONFIG_FILE;
    this.conn = new jsforce.Connection({loginUrl: this.salesforce_host, pollTimeout: 100000});
    this.ext_dir = '';  //Retrieveしたパッケージの展開ディレクトリパス
    this.objs = [];     //カスタムオブジェクト一覧{apiname:API参照名, label:表示名}
    this.metadata = {}; //Salesforce組織のメタデータ情報
};

/**
 * * initialize
 * * 初期化処理。Salesforce組織へ認証を行う。
 * @returns {Promise}
 */
SalesforceSpec.prototype.initialize = function(){
    log.debug('SalesforceSpec:initialize');
    var spec_this = this;

    return Promise.resolve()
        .then(function(){
            return spec_this.conn.login(
                spec_this.salesforce_username,
                spec_this.salesforce_password
            );
        }).then(function(){
            return spec_this.retrieve_package();
        }).then(function(custom_objs){
            spec_this.metadata = new Metadata();
            spec_this.metadata.custom_objs = custom_objs;
            return spec_this.get_objs();
        }).then(function(objs){
            spec_this.metadata.objs = objs;
            return spec_this.get_validation_rules();
        }).then(function(validation_rules){
            spec_this.metadata.validation_rules = validation_rules;
            return spec_this.get_recordtypes();
        }).then(function(recordtypes){
            spec_this.metadata.recordtypes = recordtypes;
            return spec_this.get_recordtype_picklist();
        }).then(function(recordtype_picklist){
            spec_this.metadata.recordtype_picklist = recordtype_picklist;
            return spec_this.get_layout_assignment();
        }).then(function(layout_assignment){
            spec_this.metadata.layout_assignment = layout_assignment;
            return spec_this.get_fields();
        }).then(function(fields){
            spec_this.metadata.fields = fields;
            return spec_this.get_profile_crud();
        }).then(function(profile_crud){
            spec_this.metadata.profile_crud = profile_crud;
            spec_this.metadata.valid_profiles = spec_this.get_valid_profiles();
            return spec_this.get_field_permission();
        }).then(function(field_permission){
            spec_this.metadata.field_permission = field_permission;
            return spec_this.get_workflow();
            //work_flows
        }).then(function(work_flows){
            spec_this.metadata.work_flows = work_flows;
        }).catch(function(err){
            console.log(err);
            return;
        });
}

/**
 * * retrieve_package
 * *  1. Create temporary directory 
 * *  2. Retrieve metadata
 * *  3. Unzip metadata file
 * @returns {Promise} path of unzipped directory
 */
SalesforceSpec.prototype.retrieve_package = function(){
    log.debug('SalesforceSpec:retrieve_package');
    var spec_this = this;
    return Promise
        .all([
             spec_this.desc_objects()
            ,spec_this.desc_layouts()
            ,spec_this.desc_profiles()
        ]).then(function(results){
            var objs = results[0];
            var layouts = results[1];
            var profiles = results[2];

            var package = {};
            package.types = [];
            _.each(objs, function(e){
                package.types.push({'name': 'CustomObject', 'members': e});
                package.types.push({'name': 'Workflow', 'members': e});
            })
            _.each(layouts, function(e){
                package.types.push({'name': 'Layout', 'members': e});
            })
            _.each(profiles, function(e){
                package.types.push({'name': 'Profile', 'members': e});
            })
            package.version = '33.0';
            return new Promise(function(resolve, reject){
                tmp.dir({dir:spec_this.temporary_directory},function(err, dir) {
                    if (err)
                        reject(err);
                    spec_this.conn.metadata.pollTimeout = 100000;
                    log.debug('SalesforceSpec:temp directory is successfully created');
                    var pipe_result = spec_this.conn.metadata.retrieve({unpackaged: package})
                        .stream()
                        .pipe(
                        fs.createWriteStream(
                                dir +'/'+ "unpackaged.zip"
                        )
                    ).on('finish',function(e){
                            var zip = new AdmZip('./' + dir + '/unpackaged.zip');
                            zip.extractAllTo('./' + dir + '/', true);
                            spec_this.ext_dir = './' + dir +'/'+ 'unpackaged/';
                            log.debug('SalesforceSpec:metadata is successfully retrieve.');
                            resolve(objs);
                        }).on('error',function(e){
                            reject(e);
                        });
                });
            });
    });
}

/**
 * * parse_directory
 * * 引数で指定したメタデータタイプのファイルをパース、返却する
 * @param metadata_type
 */
SalesforceSpec.prototype.parse_directory = function(metadata_type){
    var spec_this = this;
    log.info('parse_directory. metadata_type:' + metadata_type);
    return fs.readdirAsync(spec_this.ext_dir + '/' + metadata_type + '/')
    .map(function(file_name_with_extension){
        var file_name = file_name_with_extension.match(/(.+)(\.[^.]+$)/)[1];
        return fs.readFileAsync(spec_this.ext_dir + '/' + metadata_type + '/' + file_name_with_extension, 'utf8')
            .then(function(content) {
                return {file_name: file_name, file_content: content};
            }
        );
    }).each(function(file){
        return parseString(file.file_content)
            .then (function(file_content){
                file.file_content = file_content;
                return file;
            }
        );
    })
};

SalesforceSpec.prototype.field_type = function(){
    var type_label = {
        Email:'メール',Html:'リッチテキストエリア',LongTextArea:'ロングテキストエリア',Number:'数値',
        Picklist:'選択リスト',MultiselectPicklist:'複数選択リスト',Location:'位置',Currency:'通貨',
        Phone:'電話番号',Date:'日付',DateTime:'日付/時間',Lookup:'参照関係',MasterDetail:'主従関係',Checkbox:'チェックボックス',
        Text:'テキスト',EncryptedText:'テキスト(暗号化)',TextArea:'テキストエリア',Percent:'パーセント',Url:'URL',
        AutoNumber:'自動採番項目',Summary:'積上集計'
    }
    return Promise.props({
        objects:this.parse_directory('objects')
    }).then(function(meta){
        var all_fields = {};
        _.each(meta.objects, function(obj){
            var object_name = obj.file_name;
            _.each(obj.file_content.CustomObject.fields, function(field){
                all_fields[object_name + '.' + field.fullName[0]] = type_label[field.type[0]];
            });
        });
        return all_fields;
    });
};



SalesforceSpec.prototype.field_label = function(){
    return Promise.props({
        objects:this.parse_directory('objects')
    }).then(function(meta){
        var all_fields = {};
        _.each(meta.objects, function(obj){
            var object_name = obj.file_name;
            _.each(obj.file_content.CustomObject.fields, function(field){
                all_fields[object_name + '.' + field.fullName[0]] = field.label[0];
            });
        });
        return all_fields;
    });
};

SalesforceSpec.prototype.page_layouts = function(){
    return Promise.props({
        objects:this.parse_directory('objects'),   
        layouts:this.parse_directory('layouts')
    }).then(function(meta){
        var layout_summary = {}; //レイアウト名 * "FieldName = {}"
        _.each(meta.layouts, function(layout){
            var field_names = {};
            _.each(layout.file_content.Layout.layoutSections, function(section){
                _.each(section.layoutColumns,function(column){
                    _.each(column.layoutItems, function(item){
                        if(item.field)
                            field_names[item.field[0]]=item.behavior[0];
                    });
                });
            });
            layout_summary[layout.file_name] = field_names;
        });
        var all_layouts = {};
        _.each(meta.objects, function(obj){
            var layouts_on_thisobject = {};
            _.each(Object.keys(layout_summary), function(layout_name){
                var is_layout_thisobject = new RegExp('^'+obj.file_name);
                if(is_layout_thisobject.test(layout_name)){
                    layouts_on_thisobject[layout_name] = layout_summary[layout_name];
                }
            });
            all_layouts[obj.file_name] = layouts_on_thisobject;
        });
        var all_fields = {};
        _.each(meta.objects, function(obj){
            var object_name = obj.file_name;
            var fields = [];
            _.each(obj.file_content, function(field){
                _.each(field.fields, function(field_entry){
                    fields.push(field_entry.fullName[0]);
                });
            });
            all_fields[object_name] = fields;
        });
        var mapping = {};
        _.each(Object.keys(all_fields), function(obj_name){
            var fields = all_fields[obj_name];
            var this_layout = all_layouts[obj_name];
            var mapping_by_object = {};
            _.each(fields, function(field_name){
                var mapping_by_object_field = {};
                _.each(Object.keys(this_layout), function(layout_name){
                    var layout_assignment = this_layout[layout_name];
                    mapping_by_object_field[layout_name] = layout_assignment[field_name] === undefined ? '' : layout_assignment[field_name];
                });
                mapping_by_object[field_name] = mapping_by_object_field;
            });
            mapping[obj_name] = mapping_by_object;
        });
        return mapping;
    })
};

/**
 * * get_validation_rules
 * * 入力規則一覧を取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_validation_rules = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        var walker = walk.walk(spec_this.ext_dir + '/objects/');
        var validations = {};
        walker.on("file",
            function(root, fileStat, next){
                fs.readFile(root+'/'+fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var rules = [];
                        if(result.CustomObject.validationRules){
                            _.each(result.CustomObject.validationRules, function(e){
                                var rule_entry = {};
                                rule_entry.fullName = e.fullName ? e.fullName[0] : null;
                                rule_entry.active = util.desc(e.active);
                                rule_entry.description = e.description ? e.description[0] : null;
                                rule_entry.errorConditionFormula = e.errorConditionFormula ? e.errorConditionFormula[0] : null;
                                rule_entry.errorMessage = e.errorMessage ? e.errorMessage[0] : null;
                                rule_entry.objectname = fileStat.name.replace('.object', '');
                                rule_entry.errorDisplayField = e.errorDisplayField ? e.errorDisplayField[0] : 'ページの最上部';
                                rules.push(rule_entry);
                            });
                        }
                        validations[fileStat.name.replace('.object','')] = rules;
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function(){
                resolve(validations);
            }
        );
        walker.on("error",
            function(e){
                reject(e);
            }
        );
    });
}

/**
 * * get_recordtypes
 * * レコードタイプ一覧を取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_recordtypes = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        var walker = walk.walk(spec_this.ext_dir + '/objects/');
        var recordtypes = {};
        walker.on("file",
            function(root, fileStat, next){
                fs.readFile(root+'/'+fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var types = [];
                        if(result.CustomObject.recordTypes){
                            _.each(result.CustomObject.recordTypes, function(e){
                                type = {};
                                type.fullName = e.fullName? e.fullName[0] : null;
                                type.active = e.active? e.active[0] : null;
                                type.description = e.description? e.description[0] : null;
                                type.label = e.label? e.label[0] : null;
                                types.push(type);
                            });
                        }
                        recordtypes[fileStat.name.replace('.object','')] = types;
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function(){
                resolve(recordtypes);
            }
        );
        walker.on("error",
            function(e){
                reject(e);
            }
        );
    });
}

/**
 * * get_recordtype_picklist
 * * レコードタイプ毎の選択リスト値を取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_recordtype_picklist = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        var walker = walk.walk(spec_this.ext_dir + '/objects/');
        var recordtype_picklist = {};
        walker.on("file",
            function(root, fileStat, next){
                fs.readFile(root+'/'+fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var map = {};
                        _.each(result.CustomObject.recordTypes, function(record_type){
                            var picklistValues = {};
                            _.each(record_type.picklistValues, function(e){
                                var pick_list_name = e.picklist;
                                var values = '';
                                _.each(e.values, function(v){
                                    values = values + v.fullName + '\r\n';
                                });
                                picklistValues[pick_list_name] = values;
                            });
                            map[record_type.fullName] = picklistValues;
                        });
                        recordtype_picklist[fileStat.name.replace('.object','')] = map;
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function(){
                resolve(recordtype_picklist);
            }
        );
        walker.on("error",
            function(e){
                reject(e);
            }
        );
    });
}

/**
 * * get_layout_assignment
 * * プロファイル × レコードタイプ × ページレイアウト のマッピングを取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_layout_assignment = function(){
    var spec_this = this;
    return Promise.resolve()
        .then(function(){
            return spec_this.get_recordtypes();
        })
        .then(function(recordtypes){
            var recordtype_name = {};
            _.each(_.keys(recordtypes), function(objectname){
                _.each(recordtypes[objectname],function(record_type){
                    recordtype_name[objectname+'.'+record_type.fullName] = {name:objectname, label:record_type.label};
                });
            });

            var walker = walk.walk(spec_this.ext_dir +'/profiles/');
            var record_type_profile = [];
            walker.on("file",
                function(root, fileStat, next){

                    fs.readFile(root+'/'+fileStat.name, function (err, data) {
                        var record_types = {};
                        parser.parseString(data, function (err, result) {
                            _.each(result.Profile.layoutAssignments, function(e){
                                if(e.recordType){
                                    record_types[e.recordType] = e.layout;
                                }
                            });
                            record_type_profile[fileStat.name.replace('.profile','')] = record_types;
                        });
                        next();
                    });
                }
            );
            walker.on("end",
                function() {
                    var obj1 = {};
                    _.each(_.keys(recordtypes), function (objname) {
                        var obj2 = {};
                        _.each(_.keys(record_type_profile), function (profilename) {
                            var obj3 = {};
                            _.each(_.keys(recordtype_name), function (object_recordtype_name) {
                                if (object_recordtype_name.indexOf(objname + '.') != -1) {
                                    var recodtype_layout = record_type_profile[profilename];
                                    var layoutname = recodtype_layout[object_recordtype_name][0].replace(objname+'-','');
                                    obj3[recordtype_name[object_recordtype_name].label] = layoutname;
                                }
                            });
                            obj2[profilename] = obj3;
                        });
                        obj1[objname] = obj2;
                    });
                    Promise.resolve(obj1);
                }
            );
            
        }).catch(function(err){
            Promise.reject(err);
        });
}

/**
 * * get_fields
 * * カスタム項目の一覧を取得する
 * @param dir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_fields = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        var walker = walk.walk(spec_this.ext_dir +'/objects/');
        var object_fields = {};
        walker.on("file",
            function(root, fileStat, next){
                fs.readFile(root+'/'+fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var fields = [];
                        if(result.CustomObject.fields){
                            _.each(result.CustomObject.fields, function(field){
                                var field_entry = {};
                                field_entry.label = field.label[0];
                                field_entry.apiname = field.fullName[0];
                                field_entry.type = util.convert_salesforce_type(field.type[0]);
                                if(field_entry.type === util.convert_salesforce_type('Lookup')
                                    || field_entry.type === util.convert_salesforce_type('MasterDetail')){
                                    field_entry.type = field_entry.type + '(' + spec_this.get_labelname(field.referenceTo[0]) + ')';
                                }else if(field.formula){
                                    field_entry.type = '数式(' + field_entry.type + ')';
                                }
                                field_entry.size = field['length']? field['length'][0] : (field.precision? '(' + field.precision[0] + ',' + field.scale[0] + ')' : '');
                                field_entry.externalId = util.desc(field.externalId);
                                field_entry.formula = field.formula? field.formula[0] : null;
                                field_entry.required = util.desc(field.required);
                                field_entry.unique = util.desc(field.unique);
                                field_entry.trackHistory = util.desc(field.trackHistory);
                                field_entry.trackTrending = util.desc(field.trackTrending);
                                field_entry.defaultValue = field.defaultValue? field.defaultValue[0] : '';
                                field_entry.picklistValues = '';
                                if(field.picklist){
                                    _.each(field.picklist[0].picklistValues, function(value){
                                        field_entry.picklistValues = field_entry.picklistValues + value.fullName[0] + '\r\n';
                                    });
                                }
                                field_entry.description = field.description? field.description[0] : '';
                                fields.push(field_entry);
                            });
                            object_fields[fileStat.name.replace('.object','')] = fields;
                        }
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function(){
                resolve(object_fields);
            }
        );
        walker.on("error",
            function(e){
                reject(e);
            }
        );
    });
}

/**
 * * get_profile_crud
 * * プロファイルのObject Permissionを取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_profile_crud = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        var walker = walk.walk(spec_this.ext_dir +'/profiles/');
        var profile_crud = {};
        walker.on("file",
            function(root, fileStat, next) {
                fs.readFile(root + '/' + fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var permissions = {};
                        if (result.Profile.objectPermissions) {
                            _.each(result.Profile.objectPermissions, function(permission){
                                var permission_entry = {};
                                permission_entry.allowCreate = util.desc(permission.allowCreate);
                                permission_entry.allowDelete = util.desc(permission.allowDelete);
                                permission_entry.allowEdit = util.desc(permission.allowEdit);
                                permission_entry.allowRead = util.desc(permission.allowRead);
                                permission_entry.modifyAllRecords = util.desc(permission.modifyAllRecords);
                                permission_entry.viewAllRecords = util.desc(permission.viewAllRecords);
                                permissions[permission.object] = permission_entry;
                            });
                        }
                        profile_crud[decodeURI(fileStat.name.replace('.profile',''))] = permissions;
                    });
                    next();
                })
            }
        );
        walker.on("end",
            function(){
                resolve(profile_crud);
            }
        );
        walker.on("error",
            function(e){
                reject(e);
            }
        );
    });
};

/**
 * * get_field_permission
 * * RetrieveしたファイルからField Permissionを取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_field_permission = function() {
    var spec_this = this;
    return new Promise(function (resolve, reject) {
        var field_permissions = {};
        var walker = walk.walk(spec_this.ext_dir + '/profiles/');
        walker.on("file",
            function (root, fileStat, next) {
                fs.readFile(root + '/' + fileStat.name, function (err, data) {
                    if(err){
                        reject(err);
                    }
                    parser.parseString(data, function (err, result) {
                        if(err){
                            reject(err);
                        }
                        var permissions = {};
                        _.each(result.Profile.fieldPermissions, function(permission){
                            permissions[permission.field[0].replace('\'','')] = {};
                            permissions[permission.field[0].replace('\'','')].readable = (permission.readable[0] === 'true') ? '●' : '';
                            permissions[permission.field[0].replace('\'','')].readonly = (permission.editable[0] === 'true') ? '' : '●';
                        });
                        field_permissions[decodeURI(fileStat.name.replace('.profile',''))] = permissions;
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function () {
                resolve(field_permissions);
            }
        );
        walker.on("error",
            function (e) {
                reject(e);
            }
        );
    });
}

/**
 * * get_workflow
 * * RetrieveしたファイルからWorkFlowを取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_workflow = function() {
    var spec_this = this;
    var trigger_code = {onCreateOnly: '1', onAllChanges: '2', onCreateOrTriggeringUpdate:'3'};
    var action_type = {FieldUpdate:'項目自動更新', Alert:'メールアラート'};
    return new Promise(function (resolve, reject) {
        var work_flows = {};
        var walker = walk.walk(spec_this.ext_dir + '/workflows/');
        walker.on("file",
            function (root, fileStat, next) {
                fs.readFile(root + '/' + fileStat.name, function (err, data) {
                    if (err) {
                        reject(err);
                    }
                    parser.parseString(data, function (err, result) {
                        var alerts = result.Workflow.alerts;
                        var fieldUpdates = result.Workflow.fieldUpdates;
                        var rules = result.Workflow.rules;
                        var object_rules = [];
                        _.each(rules, function(rule){
                            _.each(rule.actions, function(action){
                                var action_entory = {};
                                if(action.type[0] === 'FieldUpdate'){
                                    target_update = _.find(fieldUpdates, function(e){return e.fullName[0] === action.name[0]});
                                    action_entory.workflow_fullName = decodeURI(rule.fullName[0]);
                                    action_entory.workflow_active = util.desc(rule.active[0]);
                                    action_entory.workflow_trigger_code = trigger_code[rule.triggerType[0]];
                                    action_entory.workflow_description = rule.description ? rule.description[0] : '';

                                    action_entory.fullName = target_update.fullName[0];
                                    action_entory.field_or_recipients = target_update.field[0];
                                    action_entory.name_or_description = target_update.name[0];
                                    action_entory.type = action_type[action.type[0]];
                                    action_entory.update_value_or_template = '';
                                    if(target_update.operation[0] === 'Formula'){
                                        action_entory.update_value_or_template = target_update.formula[0];
                                    }else if(target_update.operation[0] === 'Null') {
                                        action_entory.update_value_or_template = '\'\'';
                                    }
                                    action_entory.workflow_criteria = '';
                                    if(rule.formula){
                                        action_entory.workflow_criteria = rule.formula[0];
                                    }else if(rule.criteriaItems){
                                        if(rule.booleanFilter)
                                            action_entory.workflow_criteria = rule.booleanFilter[0] + '\r';
                                        _.each(rule.criteriaItems, function(criteria){
                                            action_entory.workflow_criteria = action_entory.workflow_criteria + criteria.field[0] + ' '+ criteria.operation[0] + ' ' + (criteria.value? criteria.value[0] : '\'\'') + '\r';
                                        });
                                    }
                                }else if(action.type[0] === 'Alert'){
                                    target_alert = _.find(alerts, function(e){return e.fullName[0] === action.name[0]});
                                    if(action_entory){
                                        action_entory.workflow_fullName = decodeURI(rule.fullName[0]);
                                        action_entory.workflow_active = util.desc(rule.active[0]);
                                        action_entory.workflow_trigger_code = trigger_code[rule.triggerType[0]];
                                        action_entory.workflow_description = rule.description ? rule.description[0] : '';

                                        action_entory.fullName = target_alert.fullName[0];
                                        action_entory.field_or_recipients = target_alert.recipients ? target_alert.recipients[0].type + ':' + target_alert.recipients[0].recipient : '';
                                        action_entory.name_or_description = target_alert.description[0];
                                        action_entory.type = action_type[action.type[0]];
                                        action_entory.update_value_or_template = target_alert.template[0];

                                        if(rule.formula){
                                            action_entory.workflow_criteria = rule.formula[0];
                                        }else if(rule.booleanFilter){
                                            action_entory.workflow_criteria = rule.booleanFilter[0] + '\r';
                                            _.each(rule.criteriaItems, function(criteria){
                                                action_entory.workflow_criteria = action_entory.workflow_criteria + criteria.field[0] + ' '+ criteria.operation[0] + ' ' + (criteria.value? criteria.value[0] : '\'\'') + '\r';
                                            });
                                        }

                                    }
                                }
                                object_rules.push(action_entory);
                            });
                        });
                        work_flows[fileStat.name.replace('.workflow','')] = object_rules;
                    });
                    next();
                });
            }
        );
        walker.on("end",
            function () {
                resolve(work_flows);
            }
        );
        walker.on("error",
            function (e) {
                reject(e);
            }
        );
    });
};

/**
 * * get_valid_profiles
 * * 有効プロファイルを取得する
 * * カスタムオブジェクトに全く権限を持たないプロファイルを除外する
 * @returns {Array}
 */
SalesforceSpec.prototype.get_valid_profiles = function(){
    var spec_this = this;
    var valid_profiles = [];
    _.each(Object.keys(this.metadata.profile_crud), function(profilename){
         var cruds = spec_this.metadata.profile_crud[profilename];
        if(Object.keys(cruds).length > 0){
            valid_profiles.push(profilename);
        }
    });
    return valid_profiles;
}
/**
 * * get_objs
 * * カスタムオブジェクトのAPI参照名と表示名を取得する
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_objs = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        spec_this.conn.describeGlobal(function(err, res){
            if (err)
                reject(err);
            var objs = [];
            _.each(res.sobjects, function(e){
                objs.push({apiname: e.name, label: e.label});
            })
            resolve(objs);
        });
    });   
}

/**
 * * desc_objects
 * * カスタムオブジェクトの一覧を返却する
 * @returns {Promise}
 */
SalesforceSpec.prototype.desc_objects = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        spec_this.conn.describeGlobal(function(err, res){
            if (err)
                reject(err);
            var objs = [];
            _.each(res.sobjects, function(e){
                if(e.custom && !e.customSetting) {
                    objs.push(e.name);
                }
            })
            resolve(objs);
        });
    });
}

/**
 * * desc_layouts
 * * ページレイアウトの一覧を返却する。
 * @returns {Promise}
 */
SalesforceSpec.prototype.desc_layouts = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        spec_this.conn.metadata.list([{type: 'Layout'}],function(err,res){
            if (err)
                reject(err);
            var layouts = [];
            _.each(res, function(e){
                layouts.push(decodeURI(e.fullName));
            })
            resolve(layouts);
        });
    });
}

/**
 * * desc_profiles
 * * プロファイルの一覧を返却する
 * @returns {Promise}
 */
SalesforceSpec.prototype.desc_profiles = function(){
    var spec_this = this;
    return new Promise(function(resolve, reject){
        spec_this.conn.metadata.list([{type: 'Profile'}],function(err,res){
            if (err)
                reject(err);
            var profiles = [];
            _.each(res, function(e){
                profiles.push(decodeURI(e.fullName));
            })
            resolve(profiles);
        });
    });
}

/**
 * * get_labelname
 * * API参照名から、カスタムオブジェクトのラベルを取得
 * @param apiname
 * @returns {*}
 */
SalesforceSpec.prototype.get_labelname = function(apiname){
    var target = _.find(this.metadata.objs, function(e){return e.apiname === apiname});
    target = target ? target.label : apiname;
    return target;
}

module.exports = SalesforceSpec;