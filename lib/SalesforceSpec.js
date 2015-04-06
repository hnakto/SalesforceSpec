/**
 * * SalesforceSpec
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/08/2015
 */

var Promise = require('bluebird');
var jsforce = require('jsforce');
var fs = Promise.promisifyAll(require("fs"));
var _ = require('underscore');
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var AdmZip = require('adm-zip');
var temp = require('temp').track();
var yaml = require('js-yaml');
var Log = require('log');
var log = new Log(Log.DEBUG);

require('dotenv').load();

var config = load_config();

var SalesforceSpec = function() {};

SalesforceSpec.prototype.initialize = function(){
    log.info('SalesforceSpec:initialize');
    var spec = this;
    var username = process.env.SALESFORCE_USERNAME;
    var password = process.env.SALESFORCE_PASSWORD;
    var host = 'https://' + process.env.SALESFORCE_HOST;
    var temporary_directory = config.temporary_directory;
    this.conn = new jsforce.Connection({loginUrl: host});
    
    return this.conn.login(username,password)
        .then(function(){
            return spec.retrieve_package();
        }).then(function(){
            return spec.parse_all_metadata();
        });
};

SalesforceSpec.prototype.retrieve_package = function(){
    log.info('SalesforceSpec:retrieve_package');
    var spec = this;
    spec.ext_dir = temp.mkdirSync();
    return Promise.props({
            objs:spec.list_metadata('CustomObject'),
            layouts:spec.list_metadata('Layout'),
            profiles:spec.list_metadata('Profile')
        }).then(function(meta){
            var package = {types:[],version:process.env.SALESFORCE_VERSION};
            _.each(meta.objs, function(e){ package.types.push({'name': 'CustomObject', 'members': e});});
            _.each(meta.objs, function(e){ package.types.push({'name': 'Workflow', 'members': e});});
            _.each(meta.layouts, function(e){package.types.push({'name': 'Layout', 'members': e});});
            _.each(meta.profiles, function(e){package.types.push({'name': 'Profile', 'members': e});});

            return new Promise(function(resolve, reject){
                spec.conn.metadata.pollTimeout = 100000;
                log.info('SalesforceSpec:temp directory is successfully created');
                var pipe_result = spec.conn.metadata.retrieve({unpackaged: package})
                    .stream()
                    .pipe(fs.createWriteStream(spec.ext_dir + '/unpackaged.zip'))
                    .on('finish',function(){
                        new AdmZip(spec.ext_dir + '/unpackaged.zip').extractAllTo(spec.ext_dir + '/', true);
                        resolve();
                    }).on('error',reject);
            });
    });
}

SalesforceSpec.prototype.parse_all_metadata = function(){
    var spec = this;
    return Promise.props({
        layouts: spec.parse_directory('layouts'),
        objects: spec.parse_directory('objects'),
        profiles: spec.parse_directory('profiles'),
        workflows: spec.parse_directory('workflows')
    }).then(function(meta){
        meta.objects = _.filter(meta.objects, function(object){
            var checked1 = (object.file_name.indexOf('__c') !== -1) && (object.file_content.CustomObject.customSettingsType === undefined);
            var checked2 = _.find(config.target_standard_object, function(e){ return e === object.file_name;});
            return (checked1 || checked2)
        });
        meta.profiles = _.filter(meta.profiles, function(profile){
            var checked1 = _.find(config.target_license, function(e){ return e === profile.file_content.Profile.userLicense[0];});
            var checked2 = _.find(config.target_profile, function(e){ return e === profile.file_name;});
            return (checked1 || checked2)
        });
        spec.layouts = meta.layouts;
        spec.objects = meta.objects;
        spec.object_names = _.map(meta.objects,function(e){return e.file_name;});
        spec.profiles = meta.profiles;
        spec.workflows = meta.workflows;
        return;
    });
}

SalesforceSpec.prototype.parse_directory = function(metadata_type){
    var spec = this;
    log.info('parse_directory. metadata_type:' + metadata_type);
    return fs.readdirAsync(spec.ext_dir + '/unpackaged/' + metadata_type + '/')
    .map(function(file_name_with_extension){
        var file_name = file_name_with_extension.match(/(.+)(\.[^.]+$)/)[1];
        return fs.readFileAsync(spec.ext_dir + '/unpackaged/' + metadata_type + '/' + file_name_with_extension, 'utf8')
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

SalesforceSpec.prototype.object_label = function(){
    var object_label = {};
    _.each(this.objects, function(obj){
        object_label[obj.file_name] = obj.file_content.CustomObject.label === undefined ? obj.file_name : obj.file_content.CustomObject.label[0];
    });
    return object_label;
};

SalesforceSpec.prototype.custom_object_summary = function(){
    var all_object = [];
    _.each(this.objects, function(obj){
        if(obj.file_content.CustomObject.label){
            all_object.push({
                label:obj.file_content.CustomObject.label[0],
                fullName:obj.file_name,
                description:obj.file_content.CustomObject.description ? obj.file_content.CustomObject.description[0] : ''
            });
        }
    });
    return all_object;
};

SalesforceSpec.prototype.field_type = function(){
    var all_fields = {};
    _.each(this.objects, function(obj){
        var object_name = obj.file_name;
        _.each(obj.file_content.CustomObject.fields, function(field){
            all_fields[object_name + '.' + field.fullName[0]] = field.type === undefined ? '' : type_label[field.type[0]];
        });
    });
    return all_fields;
};

SalesforceSpec.prototype.field_label = function(){
    var all_fields = {};
    _.each(this.objects, function(obj){
        var object_name = obj.file_name;
        _.each(obj.file_content.CustomObject.fields, function(field){
            all_fields[object_name + '.' + field.fullName[0]] = field.label === undefined ? field.fullName[0] :  field.label[0];
        });
    });
    return all_fields;
};

SalesforceSpec.prototype.page_layout_list = function(){
    var layout_list = [];
    _.each(this.layouts, function(layout) {
        layout_list.push(layout.file_name);  
    });
    var all_layouts = {};
    _.each(this.objects, function(obj){
        all_layouts[obj.file_name] = [];
        _.each(layout_list, function(layout_name){
            var is_layout_thisobject = new RegExp('^'+obj.file_name);
            if(is_layout_thisobject.test(layout_name)){
                all_layouts[obj.file_name].push(layout_name.replace(obj.file_name+'-',''));
            }
        });
    });
    return all_layouts;
};

SalesforceSpec.prototype.page_layouts = function(){
    var layout_summary = {};
    _.each(this.layouts, function(layout){
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
    _.each(this.objects, function(obj){
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
    _.each(this.objects, function(obj){
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
};

SalesforceSpec.prototype.validation_rule = function(){
    var validations = {};
    _.each(this.objects, function(obj){
        var rules = [];
        if(obj.file_content.CustomObject.validationRules){
            _.each(obj.file_content.CustomObject.validationRules, function(e){
                var rule_entry = {};
                rule_entry.fullName = e.fullName ? e.fullName[0] : null;
                rule_entry.active = _on_mark(e.active);
                rule_entry.description = e.description ? e.description[0] : null;
                rule_entry.errorConditionFormula = e.errorConditionFormula ? e.errorConditionFormula[0] : null;
                rule_entry.errorMessage = e.errorMessage ? e.errorMessage[0] : null;
                rule_entry.objectname = obj.file_name;
                rule_entry.errorDisplayField = e.errorDisplayField ? e.errorDisplayField[0] : 'ページの最上部';
                rules.push(rule_entry);
            });
            validations[obj.file_name] = rules;
        }
        
    });
    return validations;
};

SalesforceSpec.prototype.record_type = function(){
    var recordtypes = {};
    _.each(this.objects, function(obj){
        var types = [];
        if(obj.file_content.CustomObject.recordTypes){
            _.each(obj.file_content.CustomObject.recordTypes, function(e){
                type = {};
                type.fullName = e.fullName? e.fullName[0] : null;
                type.active = e.active? e.active[0] : null;
                type.description = e.description? e.description[0] : null;
                type.label = e.label? e.label[0] : null;
                types.push(type);
            });
            recordtypes[obj.file_name] = types;
        }
        
    });
    return recordtypes;
};

SalesforceSpec.prototype.recordtype_picklist = function(){
    var recordtype_picklist = {};
    _.each(this.objects, function(obj) {
        if(obj.file_content.CustomObject.recordTypes){
            var map = {};
            _.each(obj.file_content.CustomObject.recordTypes, function (record_type) {
                var picklistValues = {};
                _.each(record_type.picklistValues, function (e) {
                    var pick_list_name = e.picklist;
                    var values = '';
                    _.each(e.values, function (v) {
                        //values = values + v.fullName + '\n';
                        values = values + v.fullName;
                    });
                    picklistValues[pick_list_name] = values;
                });
                map[record_type.fullName] = picklistValues;
            });
            recordtype_picklist[obj.file_name] = map;
        }
    });
    return recordtype_picklist;
};

SalesforceSpec.prototype.custom_field = function(){
    var spec = this;
    var object_fields = {};
    _.each(this.objects, function(obj) {
        if(obj.file_content.CustomObject.fields === undefined)
            return;
        var fields = [];
        _.each(obj.file_content.CustomObject.fields, function(field){
            var field_entry = {};
            field_entry.label = field.label ? field.label[0] : field.fullName[0];
            field_entry.apiname = field.fullName[0];
            field_entry.type = field.type ? type_label[field.type[0]] : '';
            if(field_entry.type === type_label['Lookup']
                || field_entry.type === type_label['MasterDetail']){
                field_entry.type = field_entry.type + (field.referenceTo ? '(' + spec.label_name(field.referenceTo[0]) + ')' : '');
            }else if(field.formula){
                field_entry.type = '数式(' + field_entry.type + ')';
            }
            field_entry.size = field['length']? field['length'][0] : (field.precision? '(' + field.precision[0] + ',' + field.scale[0] + ')' : '');
            field_entry.externalId = _on_mark(field.externalId);
            field_entry.formula = field.formula? field.formula[0] : null;
            field_entry.required = _on_mark(field.required);
            field_entry.unique = _on_mark(field.unique);
            field_entry.trackHistory = _on_mark(field.trackHistory);
            field_entry.trackTrending = _on_mark(field.trackTrending);
            field_entry.defaultValue = field.defaultValue? field.defaultValue[0] : '';
            field_entry.picklistValues = '';
            if(field.picklist){
                _.each(field.picklist[0].picklistValues, function(value){
                    field_entry.picklistValues = field_entry.picklistValues + value.fullName[0] + '\n';
                });
            }
            field_entry.description = field.description? field.description[0] : '';
            fields.push(field_entry);
        });
        object_fields[obj.file_name] = fields;
    });
    return object_fields;
};

SalesforceSpec.prototype.object_permission = function(){
    var spec = this;
    var all_permissions = {};
    _.each(this.profiles, function(profile) {
        if (profile.file_content.Profile.objectPermissions === undefined)
            return;
        var permissions = {};
        _.each(profile.file_content.Profile.objectPermissions, function(permission){
            var permission_entry = {};
            permission_entry.allowCreate = _on_mark(permission.allowCreate);
            permission_entry.allowDelete = _on_mark(permission.allowDelete);
            permission_entry.allowEdit = _on_mark(permission.allowEdit);
            permission_entry.allowRead = _on_mark(permission.allowRead);
            permission_entry.modifyAllRecords = _on_mark(permission.modifyAllRecords);
            permission_entry.viewAllRecords = _on_mark(permission.viewAllRecords);
            permissions[permission.object] = permission_entry;
        });
        all_permissions[decodeURI(profile.file_name)] = permissions;
    });
    return all_permissions;
};

SalesforceSpec.prototype.field_permission = function() {
    var spec = this;
    var all_field_permission = {};
    _.each(this.profiles, function(profile) {
        var permissions = {};
        if(profile.file_content.Profile.fieldPermissions === undefined)
            return;
        _.each(profile.file_content.Profile.fieldPermissions, function(permission){
            permissions[permission.field[0].replace('\'','')] = {};
            permissions[permission.field[0].replace('\'','')].readable = (permission.readable[0] === 'true') ? '●' : '';
            permissions[permission.field[0].replace('\'','')].readonly = (permission.editable[0] === 'true') ? '' : '●';
        });
        all_field_permission[decodeURI(profile.file_name)] = permissions;
    });
    return all_field_permission;
};

SalesforceSpec.prototype.workflow = function() {
    var spec = this;
    var trigger_code = {onCreateOnly: '1', onAllChanges: '2', onCreateOrTriggeringUpdate:'3'};
    var action_type = {FieldUpdate:'項目自動更新', Alert:'メールアラート'};
    var all_work_flow = {};
    _.each(this.workflows, function(workflow) {
        var alerts = workflow.file_content.Workflow.alerts;
        var fieldUpdates = workflow.file_content.Workflow.fieldUpdates;
        var rules = workflow.file_content.Workflow.rules;
        var object_rules = [];
        _.each(rules, function(rule){
            _.each(rule.actions, function(action){
                var action_entory = {};
                //対応するアクションは項目自動更新とメールアラートのみ
                if(action.type[0] === 'FieldUpdate'){
                    target_update = _.find(fieldUpdates, function(e){return e.fullName[0] === action.name[0]});
                    action_entory.workflow_fullName = decodeURI(rule.fullName[0]);
                    action_entory.workflow_active = _on_mark(rule.active[0]);
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
                        action_entory.workflow_active = _on_mark(rule.active[0]);
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
        all_work_flow[workflow.file_name] = object_rules;
    });
    return all_work_flow;
};

SalesforceSpec.prototype.valid_profile = function(){
    var spec = this;
    var profiles = [];
    _.each(Object.keys(this.object_permission()), function(profilename){
         var cruds = spec.object_permission()[profilename];
        if(Object.keys(cruds).length > 0){
            profiles.push(profilename);
        }
    });
    return profiles;
};

SalesforceSpec.prototype.list_metadata = function(metadata_type){
    var spec = this;
    return Promise.resolve()
        .then(function(){
            return spec.conn.metadata.list([{type: metadata_type}]);
        }).map(function(e){
            return e.fullName;
        });
};

SalesforceSpec.prototype.label_name = function(apiname){
    var target = _.find(this.object_names, function(e){return e.apiname === apiname});
    target = target ? target.label : apiname;
    return target;
};

var type_label = {
    Email:'メール',Html:'リッチテキストエリア',LongTextArea:'ロングテキストエリア',Number:'数値',
    Picklist:'選択リスト',MultiselectPicklist:'複数選択リスト',Location:'位置',Currency:'通貨',
    Phone:'電話番号',Date:'日付',DateTime:'日付/時間',Lookup:'参照関係',MasterDetail:'主従関係',Checkbox:'チェックボックス',
    Text:'テキスト',EncryptedText:'テキスト(暗号化)',TextArea:'テキストエリア',Percent:'パーセント',Url:'URL',
    AutoNumber:'自動採番項目',Summary:'積上集計'
};

function _on_mark(flag){
    if(flag === undefined || flag === null || flag[0] === undefined
        || flag[0] === null || flag[0] === false || flag[0] === 'false'){
        return '';
    }else{
        return '●';
    }
}

function load_config(){
    var config = yaml.safeLoad(fs.readFileSync('./yaml/config.yml', 'utf8'));
    return config;
}
module.exports = SalesforceSpec;