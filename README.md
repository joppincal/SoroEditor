# そろエディタ SoroEditor

複数列が横に揃って並ぶテキストエディタです。

列ごとに情報を分けて記述ができます。

ツールバーやステータスバー、フォント、テーマカラーなどを自由にカスタマイズできます。

列の動きを同期して見えている行をそろえることができます。

## スクリーンショット
- lumenテーマ（デフォルト）
![Screenshot-1](/image/screenshot_lumen.png)
- superheroテーマ
![Screenshot-2](/image/screenshot_superhero.png)

このソフトウェアは
[israel-dryer/ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap)
を利用しており、テーマはttkbootstrapのものです。

テーマについて詳しくは [リンク先](https://ttkbootstrap.readthedocs.io/en/latest/themes/)
を参照してください。

## 使用例

- 台本制作
- スケジュール帳
- etc...

## ショートカットキー

- Ctrl+O:           ファイルを開く
- Ctrl+R:           前回使用したファイルを開く
- Ctrl+S:           上書き保存
- Ctrl+Shift+S:     名前をつけて保存
- Ctrl+Shift+E:     エクスポート
- Ctrl+Shift+I:     インポート
- F5:               再読み込み

- Ctrl+X:           カット
- Ctrl+C:           コピー
- Ctrl+V:           ペースト
- Ctrl+A:           すべて選択
- Ctrl+L:           1行選択
- Enter:            1行追加（下）
- Ctrl+Enter:       1行追加（上）
- Shift+Enter:      通常の改行
- Ctrl+Delete:      1行削除
- Ctrl+Z:           取り消し
- Ctrl+Shift+Z:     取り消しを戻す

- Ctrl+F:           検索
- Ctrl+Shift+F:     置換

- Ctrl+T:           定型文
- Ctrl+1-0:         各定型文の入力

- Ctrl+B:           付箋

- Ctrl+Shift+P:     設定

- Ctrl+Q, Alt+<:    左の列に移動
- Ctrl+W, Alt+>:    右の列に移動

## プロジェクトファイルについて

プロジェクトファイルはYAML形式のファイルの拡張子を変更したものです（*.cep）

一般のテキストエディタでも閲覧・編集が可能です

## 設定ファイルについて

設定ファイルはYAML形式です（setting.yaml）

一般のテキストエディタでも閲覧・編集が可能です

実行ファイル（SoroEditor.exe または SoroEditor.py）と同じフォルダに置いてください

## Releasesについて

[Nuitka](https://github.com/Nuitka/Nuitka)によって作成したWindows向けexeです

コマンドは以下の通り
```PowerShell
nuitka --mingw64 --onefile --enable-plugin=tk-inter --include-data-files=ThirdPartyNotices.txt= --include-data-files=hello.txt= --include-data-dir=src=src --disable-console --clang --windows-icon-from-ico="src/icon/icon.ico” SoroEditor.py
```
