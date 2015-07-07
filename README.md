SalesforceSpec
============
Batch script that creates the followings specification files of your salesforce organization.
* Custom Objects
* Object Permissions
* Custom Fields
* Field Level Securities
* Page Layouts
* Workflow Rules
* Validation Rules


### 1.git clone

```
git clone git@github.com:hagasatoshi/SalesforceSpec.git
cd SalesforceSpec
```

### 2.create credential file

```
touch .env
```

write information to .env as follows.

```bash
SALESFORCE_USERNAME= user name of your salesforce organization
SALESFORCE_PASSWORD= password of your salesforce organization. Please append security token if required
SALESFORCE_HOST= hostname of your salesforce organization, for example ap.salesforce.com
SALESFORCE_VERSION= version of your salesforce organization
```

### 3.run app

```bash
node app.js
```

### note
* you can change configuration of application by editing yaml/config.yml <br/>
Please check the page https://github.com/hagasatoshi/SalesforceSpec/tree/master/yaml for detail.



