# ITB- Excel ファイル進捗チェッカー

指定フォルダ配下の `ITB-` で始まる Excel ファイルを検証し、結果を Excel に出力するツールです。

## 全体の流れ

```mermaid
flowchart LR
    A[フォルダを選択] --> B[ITB- Excel を検索]
    B --> C[各シートを検証]
    C --> D[検証結果_YYYYMMDD_HHMMSS.xlsx\nを出力]
```

## 前提条件

- Python 3.10 以上
- PowerShell

Python 未インストールの場合は [python.org](https://www.python.org/downloads/) からダウンロードしてください。  
インストール時に **「Add Python to PATH」にチェック** を入れてください。

## セットアップ

```mermaid
flowchart TD
    S1["① 仮想環境の作成\npython -m venv .venv"] --> S2["② 仮想環境の有効化\n.\.venv\Scripts\Activate.ps1"]
    S2 --> S3["③ パッケージインストール\npip install -r requirements.txt"]
```

```powershell
cd <プロジェクトフォルダのパス>
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> **注意:** `このシステムではスクリプトの実行が無効になっています` というエラーが出る場合は、先に以下を実行してください。
>
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

有効化されるとプロンプトの先頭に `(.venv)` と表示されます。

## 実行方法

仮想環境を有効化した状態で以下を実行します。

```powershell
python checker.py
```

フォルダ選択ダイアログが表示されるので、検証対象のフォルダを選択してください。  
検証結果は選択フォルダ内に `検証結果_YYYYMMDD_HHMMSS.xlsx` として出力されます。

## 検証ルール

```mermaid
flowchart TD
    subgraph header["ヘッダー領域チェック"]
        H1["D5: シナリオ番号"]
        H2["E5: タイトル"]
        H3["E8: 概要"]
        H4["T5: テストケース作成日"]
        H5["T6: テストケース作成者"]
        H6["T7: テスト実施者"]
        H7["T8: テスト検証者"]
    end

    subgraph rows["19行目以降 行チェック"]
        R1["C列に値あり & D~V列が空 → 検証内容のみ記載"]
        R3["C列に値あり & O列が非ハイフン & Q列が空\n→ 実施日（予定）未入力"]
        R4["C列に値あり & O列が非ハイフン & S列が空\n→ 検証日（予定）未入力"]
        R5["C列に値あり & O列がハイフン & V列が空\n→ 対象外理由の記載なし"]
    end

    subgraph sheet["シートチェック"]
        S1["ITB- シートが1件のみであること"]
    end
```

| # | チェック内容 | エラーメッセージ |
|---|------------|---------------|
| 1 | D5: シナリオ番号 | シナリオ番号が未入力です |
| 2 | E5: タイトル | タイトルが未入力です |
| 3 | E8: 概要 | 概要が未入力です |
| 4 | T5: テストケース作成日 | テストケース作成日が未入力です |
| 5 | T6: テストケース作成者 | テストケース作成者が未入力です |
| 6 | T7: テスト実施者 | テスト実施者が未入力です |
| 7 | T8: テスト検証者 | テスト検証者が未入力です |
| 8 | C列に値あり & D\~V列が空 | 検証内容のみ記載のケースがあります |
| 9 | C列に値あり & O列が非ハイフン & Q列が空 | ケースIDがあるのに実施日（予定）が未入力です |
| 10 | C列に値あり & O列が非ハイフン & S列が空 | ケースIDがあるのに検証日（予定）が未入力です |
| 11 | C列に値あり & O列がハイフン & V列が空 | 実施対象外ケースの場合は欠陥内容／備考欄に理由を記載してください |
| 12 | ITB- で始まるシートが1件のみであること | ITB- で始まるシートが N 件存在します |

## 出力ファイルの形式

```mermaid
block-beta
    columns 5
    A["A\nNo."] B["B\nエラー内容"] C["C\nシート名"] D["D\nファイル名"] E["E\nフルパス"]
```

1ファイルで複数のエラーがある場合、エラーの数だけ行が出力されます。

## exe ファイルの作成

```mermaid
flowchart LR
    B["pyinstaller でビルド"] --> O["dist/ITB_Checker.exe"]
    O --> D["任意の場所にコピーして配布"]
```

### ビルドコマンド

```powershell
pyinstaller --onefile --windowed --name "ITB_Checker" checker.py
```

### 出力先

```
dist\ITB_Checker.exe
```

この exe ファイルは Python がインストールされていない PC でもそのまま実行できます。  
`dist\ITB_Checker.exe` を任意の場所にコピーして配布してください。ダブルクリックで起動します。
