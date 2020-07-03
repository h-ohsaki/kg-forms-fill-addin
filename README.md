# NAME

kg-forms-fill-addin - Microsoft Forms の回答ファイルに学生の学籍番号を記入する Excel アドイン

# DESCRIPTION

**kg-forms-fill-addin** は、XLSX 形式で保存された Microsoft Forms の回答ファイ
ルに、回答者の学籍番号を記入する Excel アドインです。

Microsoft Forms では回答者の ID を記録することは可能なのですが、回答者の学籍番
号がわからないため採点等の時に不便です。**kg-forms-fill-addin** は回答者の学籍
番号を記入するためだけの Excel アドインです。

Microsoft Forms からダウンロードした XLSX 形式の「回答ファイル」に対して実行し
ます。D 列に記録されている回答者の ID (Microsoft 365 の ID) を別の名簿ファイル
から探索し、学生名簿に ID が存在すれば C 列に学籍番号を記入します。C 列 (もと
もとは回答完了時刻が記録されています) を書き換えることに注意してください。名簿
ファイルは LUNA (関西学院向けにカスタマイズされたBlackboard)の「名簿ダウンロー
ド」からダウンロードされた、タブ区切りのテキストファイル (ただし拡張子は .XLS)
を前提としています。

Microsoft Forms が書き出す回答ファイルの形式や LUNA が書き出す学生名簿の形式に
依存したプログラムになっています。ファイルの形式が変更になった場合は適宜プログ
ラムを書き換えてご利用ください。

# REQUIREMENTS

- Microsoft Excel

# INSTALLATION

1. kg-forms-fill.xlam を Excel アドインのディレクトリに保存

例えば、c:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kg-forms-fill.xlam
にコピーします。Excel アドインのフォルダはお使いの環境にあわせてください。

2. Kg-Forms-Fill アドインを有効にする

ファイル　→ オプション　→ アドイン → 設定で、Kg-Forms-Fill にチェックを入れ
る。

3. Kg-Forms-Fill をクイックアクセスツールバーから呼出せるように設定する

クイックアクセスツールバーのユーザー設定　→ その他のコマンド → マクロで、
forms_fill_code を選択して「追加」する。ウィンドウのタイトルバーにアドインを実
行するための小さなアイコンが表示されます。

# USAGE

1. Microsoft Forms からダウンロードした回答ファイル (\*.xlsx) を開きます。

2. クイックアクセスツールバーをクリックして、Kg-Forms-Fill アドインを実行しま
   す。

3. ファイル選択ダイアログが表示されますので、LUNA からダウンロードした名簿ファ
   イル (meibo-\*.xls) を選択してください。

回答ファイルのシートの C 列に D 列の ID に対応する学籍番号が記入されます。

# AVAILABILITY

最新版の **kg-forms-fill-addin** は
https://github.com/h-ohsaki/kg-forms-fill-addin から入手できます。

# AUTHOR

Hiroyuki Ohsaki (ohsaki[atmark]lsnl.jp)
