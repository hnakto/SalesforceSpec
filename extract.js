/**
 * Extract Salesforce metadata to Spreadsheet
 *
 * @author Satoshi Haga
 * @date 3/1/2015
 */
var Promise = require('bluebird');
var SalesforceSpec = require('./lib/SalesforceSpec');

var spec = new SalesforceSpec();

spec.retrieve_package()
.then(function(ext_dir){

    //Validation rules
    spec.get_fields(ext_dir + '/objects/').then(function(recordtypes){
        console.log(recordtypes);
    }).catch(function(err){
        console.log(err);
    });
}).catch(function(err){
    console.log(err);
});
