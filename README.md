# NAME

forms-rename - Microsoft Forms で回収したファイルのファイル名を正規化する

# DESCRIPTION

**forms-rename** は Microsoft Forms で回収したファイルを、質問番号と提出者の
ID および氏名に基づいた規則的な名前に変更する Python スクリプトです。

Microsoft Forms では、質問の種別として「ファイルのアップロード」を選択すれば、
回答者からファイルを受け取ることが可能です。回答者によってアップロードされたファ
イルは、それぞれの質問ごとに異なる SharePoint (OneDrive) 上のフォルダに格納さ
れます。

このプログラムは、Microsoft Forms で「ファイルのアップロード」によって回収され
たファイルを、整理しやすいファイル名に変更します。

※ 本来は名前を変更 (リネーム) すべきですが、現在は安全のため元のファイルを残
したまま複製 (コピー) しています。不要な場合は元のファイルを手動で削除してくだ
さい。

例えば、Microsoft Forms にアップロードされたファイルを、SharePoint (OneDrive)
から一括ダウンロードすると、以下のようなファイルになります。

- ./第2回レポート課題/Question/レポート課題_カンガク タロウ.pdf
- ./第2回レポート課題/Question/レポート課題_カンガク タロウ 1.pdf
- ./第2回レポート課題/Question/レポート課題(修正版)_カンガク タロウ.pdf
- ./第2回レポート課題/Question/report_ヤマダ ジロウ.docx
- ./第2回レポート課題/Question 1/感想_カンガク タロウ.pdf
- ./第2回レポート課題/Question 1/report 2_ヤマダ ジロウ.docx
- ./第2回レポート課題/Question 1/report 2 変更あり_ヤマダ ジロウ.docx
- ./第2回レポート課題/Question 1/report 2 変更あり_ヤマダ ジロウ 1.docx
- ./第2回レポート課題/Question 1/report 2 変更あり_ヤマダ ジロウ 2.docx

同じファイル名で複数回アップロードした場合には、ファイル名の末尾に 「 1」や
「 2」等が付与されます。しかし、同じ提出者が別のファイル名でアップロードすると、
また異なったファイル名になってしまいます。同一の提出者が複数回提出した場合、ど
のファイルが最も新しいファイルなのかがファイル名からは判別できません。

**forms-rename** は、これらのファイルを

- q1-xyz12345-カンガク_タロウ.pdf
- q1-xyz12345-カンガク_タロウ-2.pdf
- q1-xyz12345-カンガク_タロウ-3.pdf
- q1-abc98765-ヤマダ_ジロウ.docx
- q2-xyz12345-カンガク_タロウ.pdf
- q2-abc98765-ヤマダ_ジロウ.docx
- q2-abc98765-ヤマダ_ジロウ-2.docx
- q2-abc98765-ヤマダ_ジロウ-3.docx

のような正規化されたファイル名でコピーします。「設問番号-ユーザID-名前.拡張子」
というファイル名です。同一の提出者が複数回アップロードした場合は、提出時刻が古
いものから新しいものへと、順番に -2、-3、-4 等が拡張子の直前に付与されます。

従って、最後に提出した (最新の) ファイルのみをチェックするのであれば、

- q1-xyz12345-カンガク_タロウ-3.pdf
- q1-abc98765-ヤマダ_ジロウ.docx
- q2-xyz12345-カンガク_タロウ.pdf
- q2-abc98765-ヤマダ_ジロウ-3.docx

だけをチェックすればよいことになります。

# REQUIREMENTS

- Python 3
- perlcompat (https://pypi.org/project/perlcompat/)

## 実行ファイルを使用する場合

Linux (x64) または Windows 10 (x64) 向けの実行ファイルを使用する場合は Python
の実行環境等は不要です。

# INSTALLATION

```sh
pip3 install perlcompat
./forms-rename
```
## Linux (x64) または Windows 10 (x64) で実行ファイルを使用する場合

Linux (x64) および Windows 10 (x64) 用の実行ファイルを用意してあります。
PyInstaller で単一の実行ファイルに変換したものです。以下の実行ファイルをダウン
ロードして、そのまま (エクスプローラでダブルクリックする等して) 実行してくださ
い。

https://github.com/h-ohsaki/forms-rename/raw/master/linux-x64/forms-rename-1.0 (準備中)

https://github.com/h-ohsaki/forms-rename/raw/master/win10-x64/forms-rename-1.0 (準備中)


# USAGE

まず、Microsoft Forms に提出されたファイルを SharePoint (OneDrive) からダウン
ロードして、カレントディレクトリ以下に展開してください。

**forms-rename** を実行すると、ファイル選択ダイアログが開きますので、Microsoft
Forms からダウンロードした回答ファイル (Excel ファイル) を選択してください。

作業完了を知らせるダイアログ等は何も表示されませんが、カレントディレクトリに、
正規化した名前でアップロードされたファイルがコピーされています。

# AVAILABILITY

最新版の **forms-rename** は https://github.com/h-ohsaki/forms-rename から入手
できます。

# AUTHOR

Hiroyuki Ohsaki (ohsaki[atmark]lsnl.jp)
