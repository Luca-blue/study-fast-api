# study-fast-api

FastAPI 入門のハンズオン教材です。スライドを見ながら、`sample/` のファイルを開いて編集していきます。

## スライド

`slides/fastapi-intro.html` をブラウザで開いてください（28枚・1ファイル完結）。PowerPoint 版は `slides/fastapi-intro.pptx`（内容は同じ、すべて編集可能）。

- `←` `→` / `Space` / マウスホイール / スワイプでページ送り、`Home` `End` で最初・最後へ
- URL の `#13` のように番号を付けると、そのスライドから開く
- `E` キー（または画面左上にマウスを乗せる）で編集モード。`Ctrl/Cmd + S` で HTML を保存
- 印刷（PDF 保存）すると 1 スライド 1 ページで出力

| 章 | 内容 |
|---|---|
| 01 | Web API のしくみ（クライアント・Web API・DB、リクエスト/レスポンス） |
| 02 | 準備（ライブラリの役目、教材の取得とインストール、`sample/` の進め方） |
| 03 | 書いてみる（最小コード、起動、PC 内のサーバーにブラウザでアクセス、パラメータ） |
| 04 | HTTP メソッド（GET / POST / PUT / DELETE） |
| 05 | データを受け取る（Pydantic、ToDo API の関数を書く、/docs で試す） |
| 06 | 画面をつなぐ（完成品の画面で動作確認、fetch の重要なところを読む） |
| 07 | SQLite につなぐ（DB 版を読む、5章の関数との比較、再起動テスト） |

## 準備

```bash
git clone https://github.com/luca-blue/study-fast-api.git
cd study-fast-api
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install fastapi uvicorn
```

## 教材フォルダ

| ファイル | 最初の状態 | やること | 章 |
|---|---|---|---|
| `sample/01_hello/main.py` | 空 | すべて自分で書く | 3章 |
| `sample/02_todo/main.py` | import〜todos 辞書まで用意済み | API の関数だけ書く | 5章 |
| `sample/frontend/index.html` | 完成品（02・03 共通の画面） | 動かして、重要なところを読む | 6章 |
| `sample/03_todo_sqlite/` | 完成品 | 動かして、5章との違いを読む | 7章 |

各フォルダで `uvicorn main:app --reload` を実行し、http://127.0.0.1:8000/ （API は `/docs`）を開きます。

`answer/` には `01_hello` と `02_todo` の書き終わった状態のコードがあります。動かないときに見比べてください。
