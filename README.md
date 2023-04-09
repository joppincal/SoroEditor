# そろエディタ SoroEditor

複数列が横に揃って並ぶテキストエディタです。

列ごとに情報を分けて記述ができます。

ツールバーやステータスバー、フォント、テーマカラーなどを自由にカスタマイズできます。

列の動きを同期して見えている行をそろえることができます。

## スクリーンショット
- lumenテーマ（デフォルト）
![Screenshot-1](/image/screenshot.2023-03-20%2010.56.05.png)
- 使用例
![Screenshot-2](/image/screenshot.2023-03-20%2010.56.29.png)
- superheroテーマ
![Screenshot-3](/image/screenshot.2023-03-20%2010.57.12.png)

このソフトウェアは
[israel-dryer/ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap)
を利用しており、テーマはttkbootstrapのものです。

テーマについて詳しくは [リンク先](https://ttkbootstrap.readthedocs.io/en/latest/themes/)
を参照してください。

## 使用例

- 台本制作
- スケジュール帳

## ショートカットキー

- Ctrl+O:           ファイルを開く
- Ctrl+S:           上書き保存
- Ctrl+Shift+S:     名前をつけて保存
- Ctrl+Z:           取り消し
- Ctrl+Shift+Z:     取り消しを戻す
- Ctrl+Enter:       1行追加（下）
- Ctrl+Shift+Enter: 1行追加（上）
- Ctrl+L:           1行選択
- Ctrl+Q, Alt+<:   左の列に移動
- Ctrl+W, Alt+>:   右の列に移動

## プロジェクトファイルについて

プロジェクトファイルはYAML形式のファイルの拡張子を変更したものです（*.cep）

一般のテキストエディタでも閲覧・編集が可能です

## 設定ファイルについて

設定ファイルはYAML形式です（setting.yaml）

一般のテキストエディタでも閲覧・編集が可能です

実行ファイル（Soro.exe または Soro.py）と同じフォルダに置いてください