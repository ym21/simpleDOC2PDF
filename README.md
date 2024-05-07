# Simple PDF

## 概要

指定したディレクトリ内の Word,Excel ファイルを PDF に変換し、1 つの PDF ファイルに結合します。

## インストール

[リリース](https://github.com/ym21/simpleDOC2PDF/releases/)から"simpalePDF.zip"をダウンロードし、解凍してください。

## 使い方

1. simplePDF.exe を起動
2. ディレクトリ選択ダイアログが開いたら、変換したいディレクトリを選択
3. ダイアログが表示されたら、空白ページを挿入するか選択
4. 完了ダイアログが表示されたら完了

## 注意事項

- Microsoft Office2007 以上(Word,Excel)のインストールが必須です。
- 実行する前に、Word,Excel のアプリケーションを終了してください。
- 指定したディレクトリ直下のすべての Word,Excel,PDF 文書が変換の対象です。
- ファイル名昇順で結合されます。ファイル名を "001document" のように、ゼロ埋め数字を頭に付けることをお勧めします。
- 空白ページの挿入は各文書が偶数ページになるように挿入され、文書サイズは考慮されません。
- Excel 文書は、選択されているシートのみが変換対象です。事前にシート、印刷範囲を設定して下さい。
- 変換・結合された PDF は "まとめ.pdf" として保存され、同名ファイルは上書きされます。
- "（任意の文字列）+まとめ.pdf" は結合から除外されます。
- **_Word,Excel で同名のファイルがあるとバグります_**(今後修正したい)

## ライセンス

This project is licensed under the MIT License, see the LICENSE.txt file for details
