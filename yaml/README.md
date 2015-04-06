# config.yml

yamlファイルの設定値の仕様は下記の通り。

* system_name: 帳票ヘッダーのシステム名
* created_by: 帳票ヘッダーの新規作成者
* template_directory, template_file/*: テンプレートの配置ディレクトリ / ファイル名
* output_directory, output_file/*: 出力ファイルのディレクトリ / ファイル名
* target_license: 対象プロファイルのライセンス名。このライセンスのプロファイルは全て対象となる。
* target_profile: 対象プロファイル名。個別に対象プロファイルを指定する。<br/>
e.g. 初期指定の設定は、"Force.com - App Subscription"のプロファイル全てと、システム管理者<br/>
* target_standard_object: 対象の標準オブジェクト。カスタムオブジェクトはデフォルト全て対象となる。<br/>
e.g. 初期指定の設定は、全てのカスタムオブジェクトと、AccountとContact
