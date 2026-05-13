# rename_pdf

PDFの内容を読み取り、業務で扱いやすいファイル名へ自動リネームするPython製GUIツールです。

住宅設備の発注関連PDFを対象に、書類種別の判定、現場情報の抽出、確定日の計算、PDFへの追記、投函用PDFの結合までをまとめて処理します。手作業でPDFを開いて内容を確認し、ファイル名を付け替え、関連書類をまとめる作業を減らすために作成しました。

## 特徴

- PDFテキストを解析し、書類種別を自動判定
- 店舗コード、管理番号、現場名などの情報を抽出
- 抽出結果をもとにPDFファイル名を自動生成
- 出荷情報と休日データから確定日を計算
- 必要に応じてPDFへ確定日を追記
- 最終確認票と仕様明細書を投函用PDFとして結合
- tkinter GUIで対象フォルダを選択して実行
- `pytest`、`ruff`、`mypy`、GitHub Actionsで品質確認

## 想定ユースケース

大量のPDFを受け取り、内容に応じてファイル名を整えたり、関連書類をまとめたりする定型業務向けのツールです。

特に次のような作業を自動化します。

- PDFを1件ずつ開いて現場名や管理番号を確認する
- 書類種別ごとに命名ルールを変えてリネームする
- 休日を考慮して確定日を計算する
- 投函用に関連PDFを結合する

## 対応している書類

- ユニットバスルーム納期最終確認票
- ユニットバスルームご発注確認票
- 御見積書
- 仕様変更確認票
- キャンセル確認票

## 技術スタック

- Python 3.9+
- tkinter
- pdfminer.six
- PyPDF2
- reportlab
- python-dateutil
- cx_Freeze
- pytest / ruff / mypy

## セットアップ

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

PowerShellの場合:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

開発用ツールも含めてセットアップする場合:

```bash
make setup-dev
```

## 使い方

このツールはGUIアプリケーションです。開発環境から実行する場合は、次のコマンドで起動します。

```bash
PYTHONPATH=src python src/renamePdf.py
```

PowerShellの場合:

```powershell
$env:PYTHONPATH = "src"
python src/renamePdf.py
```

GUIが起動したら、処理対象のPDFが入っているフォルダを選択して実行します。

### 事前に必要なファイル

休日計算には `data/休日.csv` を使用します。ビルド済みアプリとして配布する場合は、`setup.py` の設定により `休日.csv` が実行ファイル側へ同梱されます。

CSVは次の形式を想定しています。

| date | holiday_name |
| --- | --- |
| 2024/01/01 | 元日 |
| 2024/01/08 | 成人の日 |
| 2024/02/11 | 建国記念の日 |

## 処理の流れ

1. GUIで対象フォルダを選択
2. フォルダ内のPDFを読み込み
3. PDF内の文字情報から書類種別を判定
4. 書類種別ごとの抽出ロジックで必要情報を取得
5. 命名ルールに沿ってファイル名を生成
6. 必要なPDFへ確定日を書き込み
7. 関連PDFを投函用ファイルとして結合

## アーキテクチャ

`src/` 配下は、テストしやすいように責務ごとに分割しています。

```text
src/
├── classifiers.py          # PDFの書類種別判定
├── services.py             # フォルダ処理、リネーム、PDF結合の orchestration
├── models.py               # 処理結果やPDFテキスト要素のデータ構造
├── extractors/             # PDFテキスト抽出
├── processors/             # 書類種別ごとの情報抽出と命名ルール
└── writers/                # ファイル操作とPDF書き込み
```

GUIエントリポイントは `src/renamePdf.py` です。

## ビルド

`cx_Freeze` を使ってWindows向け実行ファイルを作成できます。

```bash
python setup.py build
```

ビルド時には、次のファイルが同梱されます。

- `icons/icon.ico`
- `resources/resource.res`
- `data/休日.csv`

## テストと品質確認

開発環境では `Makefile` から主要な品質確認を実行できます。

```bash
make test        # pytest
make lint        # ruff
make type-check  # mypy
make check       # lint + type-check + test
make ci          # format + lint + type-check + coverage付きtest
```

テストはPDF処理の中核ロジックを対象にし、GUIやWindows固有の実行環境に依存しない形で実行できるようにしています。

## このプロジェクトで示していること

- 現場の定型業務を観察し、実用的な自動化ツールに落とし込む力
- PDF解析、ファイル操作、GUI、ビルド配布を組み合わせたPython実装
- 後からテストしやすくするための段階的なモジュール分割
- `pytest`、`ruff`、`mypy`、CIを使った保守性改善

## 注意事項

- 対象PDFのレイアウトや文言が変わると、抽出ロジックの調整が必要になる場合があります。
- PDFへ文字を書き込む処理では、実行環境に利用可能なフォントが必要です。
- 実運用前には、対象業務のPDFサンプルでリネーム結果と結合結果を確認してください。
