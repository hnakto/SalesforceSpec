/**
 * * SalesforceSpec
 * *
 * @author Satoshi Haga(satoshi.haga.github@gmail.com)
 * @data 03/08/2015
 */

var Promise = require('bluebird');
var tmp = require('tmp');
var jsforce = require('jsforce');
var fs = require('fs');
var _ = require('underscore');
var xml2js = require('xml2js');
var parser = new xml2js.Parser();
var walk = require('walk');
var AdmZip = require('adm-zip');
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
};

/**
 * * retrieve_package
 * *  1. Create temporary directory 
 * *  2. Retrieve metadata
 * *  3. Unzip metadata file
 * @returns {Promise} path of unzipped directory
 */
SalesforceSpec.prototype.retrieve_package = function(){

    var spec_this = this;

    return new Promise(function(resolve, reject){
        
        var package = {};
        var conn = new jsforce.Connection({loginUrl: spec_this.salesforce_host});

        fs.readFile(spec_this.config_file, function (err, data) {
            parser.parseString(data, function (err, package_data) {
                package.types = [];
                _.each(package_data.Package.types, function(e){
                    package.types.push({'name': e.name[0], 'members': e.members[0]});
                });
                package.version = '30.0';
            });
            tmp.dir({dir:spec_this.temporary_directory},function(err, dir) {
                if (err)
                    reject(err);
                conn.login(spec_this.salesforce_username, spec_this.salesforce_password, function(err, res) {
                    if (err)
                        reject(err);
                    var pipe_result = conn.metadata.retrieve({unpackaged: package})
                        .stream()
                        .pipe(
                        fs.createWriteStream(
                                dir +'/'+ "unpackaged.zip"
                        )
                    ).on('finish',function(chunk){

                            var zip = new AdmZip('./' + dir + '/unpackaged.zip');
                            zip.extractAllTo('./' + dir + '/', true);

                            resolve('./' + dir +'/'+ 'unpackaged/');
                        }).on('error',function(e){
                            reject(e);
                        });
                });
            });
        });
    });
}

/**
 * * get_validation_rules
 * @param dir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_validation_rules = function(dir){
    return new Promise(function(resolve, reject){
        var walker = walk.walk(dir);
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
                                rule_entry.active = e.active ? e.active[0] : null;
                                rule_entry.description = e.description ? e.description[0] : null;
                                rule_entry.errorConditionFormula = e.errorConditionFormula ? e.errorConditionFormula[0] : null;
                                rule_entry.errorMessage = e.errorMessage ? e.errorMessage[0] : null;
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
 * *get_recordtypes
 * @param dir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_recordtypes = function(dir){
    return new Promise(function(resolve, reject){
        var walker = walk.walk(dir);
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
 * *get_recordtype_picklist
 * @param dir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_recordtype_picklist = function(dir){
    return new Promise(function(resolve, reject){
        var walker = walk.walk(dir);
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
 * *get_layout_assignment
 * @param basedir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_layout_assignment = function(basedir){
    return new Promise(function(resolve, reject){
        //get object-recordtype
        var walker = walk.walk(basedir+'/objects/');
        var layout_assignment = {};

        new SalesforceSpec().get_recordtypes(basedir+'/objects/')
        .then(function(recordtypes){
            var recordtype_name = {};
            _.each(_.keys(recordtypes), function(objectname){
                _.each(recordtypes[objectname],function(record_type){
                    recordtype_name[objectname+'.'+record_type.fullName] = {name:objectname, label:record_type.label};
                });
            });

            var walker = walk.walk(basedir+'/profiles/');
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
                    resolve(obj1);
                }
            );
            
        }).catch(function(err){
            reject(err);
        });
    });
}


/**
 * *get_fields
 * @param dir
 * @returns {Promise}
 */
SalesforceSpec.prototype.get_fields = function(dir){
    return new Promise(function(resolve, reject){
        var walker = walk.walk(dir);
        var object_fields = {};
        walker.on("file",
            function(root, fileStat, next){
                fs.readFile(root+'/'+fileStat.name, function (err, data) {
                    parser.parseString(data, function (err, result) {
                        var fields = [];
                        if(result.CustomObject.fields){
                            _.each(result.CustomObject.fields, function(field){
                                var field_entry = {};field_entry.type
                                field_entry.label = field.label[0];
                                field_entry.apiname = field.fullName[0];
                                field_entry.type = field.type[0];
                                if(field_entry.type === 'Lookup'
                                    || field_entry.type === 'MasterDetail'){
                                    field_entry.type = field_entry.type + '(' + field.referenceTo[0] + ')';
                                }else if(field_entry.formula){
                                    field_entry.type = 'Formula(' + field_entry.type + ')';
                                }
                                field_entry.size = field['length']? field['length'][0] : (field.precision? '(' + field.precision[0] + ',' + field.scale[0] + ')' : '');
                                field_entry.externalId = field.externalId[0];
                                field_entry.required = field.required? field.required[0] : '';
                                field_entry.picklistValues = '';
                                _.each(field.picklistValues, function(value){
                                    field_entry.picklistValues = field_entry.picklistValues + value + '\r\n';
                                });
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

module.exports = SalesforceSpec;