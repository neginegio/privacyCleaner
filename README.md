# hoso Privacy Cleaner

Windows 上で動く、ローカル完結型の Excel 匿名化ツールです。Python で作られており、PII 検出には Microsoft Presidio の `PatternRecognizer` と日本語向けカスタム Recognizer を使います。exe を起動するとネイティブな Windows アプリ画面が開き、Excel ファイルを投入して、検出結果を人が確認してから Excel として出力できます。

## 起動方法

開発実行:

```cmd
pip install -r requirements.txt
run_python_app.cmd
```

exe を作る場合:

```cmd
build_exe.cmd
```

ビルド後の実行ファイル:

```text
dist\hosoPrivacyCleaner\hosoPrivacyCleaner.exe
```

通常利用時は `Start-ExcelPrivacyCleaner.cmd` を実行してください。CSV出力対応版のネイティブ exe がある場合はその版を起動し、ない場合は既存 exe または Python 版を起動します。

## 対応していること

- 外部クラウドへ送信しないローカル処理
- PySide6 / Qt によるネイティブ Windows 画面
- `.xlsx` / `.xlsm` の読み込み
- 複数シートの検査
- Presidio カスタム Recognizer による日本語の氏名・住所らしき文字列の検出
- 氏名・住所系の列見出しによる列単位検査
- 自由記述欄の中に含まれる氏名・住所の検査
- 検出結果を画面上で確認し、変換対象をチェックで選択
- 検出結果一覧を CSV として出力
- 同じ氏名は同じ仮名、同じ住所は同じ住所番号へ統一
- 変換後の値を画面上で手修正可能
- 原本を上書きせず、`元ファイル名_匿名化_yyyyMMdd_HHmmss.xlsx` として保存
- 処理後に一時コピーを削除
- 変換履歴を画面上に表示

## テスト

自己完結型テスト(合成データのみで完結し、GitHub Actions でも実行されます):

```cmd
pip install -r requirements-dev.txt
pytest tests -v
```

`tools/test_*.py` は実データ(git 管理外の実 PDF/Excel サンプルなど)や Windows の日本語フォント、GUI に依存するテストです。ローカルにサンプルデータを用意した上で個別に実行してください:

```cmd
python tools/test_analysis_mode.py
```

これらは `tests/test_local_data.py` から `RUN_LOCAL_DATA_TESTS=1` を立てた場合のみ pytest からも実行できます(サンプルデータが無ければ自動的に SKIP されます)。CI は `pytest tests -m "not local_data"` を実行し、これらは対象外です。

## 注意

- Python 3.10 以上が必要です。exe 化後は利用者 PC に Python は不要です。
- 初回ビルド時は `pip install` により依存パッケージを取得します。取得後のアプリ処理はローカルで完結します。
- Presidio を使っていますが、日本語検出は標準 NLP モデルではなく、このアプリに同梱した日本語向けカスタム Recognizer が中心です。
- 検出は誤検出・検出漏れを前提に、保存前に必ず一覧を確認してください。
- 旧 `.xls` は Python 版では対象外です。必要な場合は Excel で `.xlsx` に保存し直してから使ってください。
- マクロ付きブックを読み込んだ場合は `.xlsm` として保存します。ただし、重要なマクロを含むファイルは事前にコピーで検証してください。
