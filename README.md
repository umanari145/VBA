# VBA

汎用的に使えそうなマクロの登録(DDL関係など)

- makeDDL
    ExcelのテンプレートからDDLを一気に作成する。<br>
    テーブルが複数あることを前提にし、テーブル一覧でチェックをつけたもののDDLを自動生成

- makeSQL
    Excelのテンプレートからinsertを作成する。<br>

- updateSQL
    Excelのテンプレートからupdateを作成する。<br>

- eachMakeSQL
    Excel出力のサブ関数(テーブル単位のSQLの作成)<br>

- util
    汎用的なモジュールなど。サンプルシート参照。flgでチェック(1)をつけたものを処理

## Excelの説明

- tableList DDLテーブルリストの対象のテーブル
- 顧客、顧客詳細などDDL作成の各シート
- dataList データ作成用リストのテーブル
- t_customer,t_communicationなどの各々のデータを入れているデータ