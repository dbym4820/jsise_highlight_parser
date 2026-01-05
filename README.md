# 採録論文ハイライト PDF パーサー

教育システム情報学会誌の「採録論文ハイライト」PDFから論文情報を抽出し，Excelファイルに出力するスクリプトです．

## 抽出される情報

| カラム | 説明 |
|--------|------|
| 種別 | 一般論文・実践論文・ショートノート・実践速報 |
| 開始ページ | 論文の開始ページ番号 |
| 終了ページ | 論文の終了ページ番号 |
| DOI | J-STAGEから自動取得（例：`https://doi.org/10.14926/jsise.43.26`） |
| タイトル | 論文タイトル |
| 著者 | 筆頭著者名と所属 |
| カテゴリ | 論文のカテゴリ |
| 分野 | 論文の分野 |
| 問い | 本論文の扱う「問い」 |
| 面白い点 | 「本論文のここが面白い！」 |

## 動作環境

- Python 3.9 以上
- macOS / Windows / Linux

## セットアップと実行手順

### 1. PDFファイルのダウンロード

1. [教育システム情報学会誌のJ-STAGEページ](https://www.jstage.jst.go.jp/browse/jsise/list/-char/ja)にアクセス
2. 該当する巻号を選択（例：43巻1号）
3. 「採録論文ハイライト」のPDFをダウンロード
4. ダウンロードしたPDFを任意のフォルダに保存（例：`~/Desktop/hilight/`）

### 2. ターミナルを開く

**macOSの場合：**
- `Command + Space` で Spotlight を開き，「ターミナル」と入力して Enter
- または，Finder → アプリケーション → ユーティリティ → ターミナル

**Windowsの場合：**
- `Win + R` で「ファイル名を指定して実行」を開き，`cmd` と入力して Enter
- または，PowerShell を使用

### 3. 作業ディレクトリに移動

```bash
# PDFを保存したフォルダに移動
cd ~/Desktop/hilight
```

### 4. 必要なライブラリをインストール

```bash
# 必須ライブラリ
pip install pymupdf openpyxl

# DOI自動取得を使用する場合（推奨）
pip install requests beautifulsoup4
```

> **注意**: `requests` と `beautifulsoup4` がない場合でも動作しますが，DOIは開始ページ番号から推測されます．

### 5. スクリプトを実行

```bash
# 基本的な使い方（出力ファイルはPDFと同名の.xlsx）
python parse_highlight.py 43_430114.pdf

# 出力ファイル名を指定する場合
python parse_highlight.py 43_430114.pdf output.xlsx
```

### 6. 出力結果

実行すると以下のような出力が表示されます：

```
PDFを読み込み中: 43_430114.pdf
検出した巻号: Vol.43 No.1
検出した論文数: 4
J-STAGEからDOI情報を取得中: https://www.jstage.jst.go.jp/browse/jsise/43/1/_contents/-char/ja
  取得したDOI数: 12
  1. [一般論文] 学習者の受講状態センシングに基づくインタラクティブロボット講... [✓]
  2. [一般論文] 学習支援ロボットのロールがプレゼンテーションセルフレビューに... [✓]
  3. [ショートノート] 精神障害領域作業療法プログラムの改善段階が学習成果に及ぼす影... [✓]
  4. [実践速報] デザイン・コンペティションを活用した授業の設計と実践̶̶デザ... [✓]
Excelファイルを保存しました: 43_430114.xlsx
```

生成されたExcelファイル（`43_430114.xlsx`）を開いて結果を確認してください．

## トラブルシューティング

### Python がインストールされていない場合

```bash
# macOS (Homebrew)
brew install python

# Windows
# https://www.python.org/downloads/ からインストーラをダウンロード
```

### pip がない場合

```bash
python -m ensurepip --upgrade
```

### ライブラリのインストールでエラーが出る場合

```bash
# Python 3 を明示的に指定
python3 -m pip install pymupdf openpyxl requests beautifulsoup4
```

### DOI取得に失敗する場合

インターネット接続を確認してください．接続できない環境では，DOIは開始ページ番号から自動推測されます（教育システム情報学会誌の場合，推測DOIは正確です）．

## ファイル構成

```
hilight/
├── parse_highlight.py   # メインスクリプト
├── README.md            # このファイル
├── 43_430114.pdf        # 入力PDFファイル（例）
└── 43_430114.xlsx       # 出力Excelファイル（例）
```

## ライセンス

MIT License
