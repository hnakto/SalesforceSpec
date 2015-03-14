/**
 * Extract Salesforce metadata to Spreadsheet
 *
 * @author Satoshi Haga
 * @date 3/1/2015
 */
var Promise = require('bluebird');
var SalesforceSpec = require('./lib/SalesforceSpec');
require('dotenv').load();

var spec = new SalesforceSpec();
spec.retrieve_package(
    process.env.TEMPORARY_DIRECTORY,
    process.env.SALESFORCE_USERNAME,
    process.env.SALESFORCE_PASSWORD,
    process.env.CONFIG_FILE
).then(function(ext_dir){

    //Validation rules
    spec.get_fields(ext_dir + '/objects/').then(function(recordtypes){
        console.log(recordtypes);
    }).catch(function(err){
        console.log(err);
    });
}).catch(function(err){
    console.log(err);
});
