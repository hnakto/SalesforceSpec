SalesforceSpec
============
Metadata APIで取得した組織情報を、Excelファイルに出力する。

####対応帳票
* カスタムオブジェクト一覧
* オブジェクト権限一覧
* カスタム項目一覧
* 項目レベル権限一覧
* レイアウト一覧
* ワークフロー一覧
* 入力規則一覧

Usage
------------

### npm install

```bash
npm install
```

### Salesforce Credentials

```bash
touch .env
```

.envにSalesforceのCredentialを記載

```bash
SALESFORCE_USERNAME=your salesforce user name
SALESFORCE_PASSWORD=your salesforce password + security token(if required)
SALESFORCE_HOST=hostname of target organization. e.g. ap.salesforce.com
SALESFORCE_VERSION=salesforce version of target organization. e.g. 33.0
```

### Run app

```bash
node app.js
```

### Note
* yaml/config.ymlを編集することにより、アプリケーションの設定値が変更可能。<br/>
詳細は https://github.com/hagasatoshi/SalesforceSpec/tree/master/yaml を参照。
* 帳票の仕様については https://github.com/hagasatoshi/SalesforceSpec/tree/master/template を参照。



