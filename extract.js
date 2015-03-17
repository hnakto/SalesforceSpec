/**
 * Extract Salesforce metadata to Spreadsheet
 *
 * @author Satoshi Haga
 * @date 3/1/2015
 */
var Promise = require('bluebird');
var SalesforceSpec = require('./lib/SalesforceSpec');
var _ = require('underscore');

var spec = new SalesforceSpec();

spec.retrieve_package()
.then(function(ext_dir){

    //Validation rules
    spec.get_validation_rules(ext_dir + '/objects/').then(function(rules){
        //console.log(rules);
        //console.log(Object.keys(rules));
        _.each(Object.keys(rules), function(obj){
            _.each(rules[obj], function(each_rule){
                each_rule.obj_name = obj;
                console.log(each_rule);
            })
        });
    }).catch(function(err){
        console.log(err);
    });
}).catch(function(err){
    console.log(err);
});
